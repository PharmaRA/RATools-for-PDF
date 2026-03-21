import os
import csv
import re
from datetime import datetime
from pathlib import Path
from PySide6.QtWidgets import QFileDialog, QMessageBox
from PySide6.QtCore import Qt, QThread, Signal
from pdf_processor import PDFProcessor


# ==========================================
# 线程一：主处理工作线程 (执行合规净化与重构)
# ==========================================
class ProcessWorker(QThread):
    progress = Signal(int, str, str)
    all_completed = Signal()

    def __init__(self, file_list, options, output_dir):
        super().__init__()
        self.file_list = file_list
        self.options = options
        self.output_dir = output_dir

    def run(self):
        rename_ectd = "文件名修改为符合电子申报/eCTD要求的格式" in self.options
        for row_idx, file_path in enumerate(self.file_list):
            self.progress.emit(row_idx, "⏳ 处理中...", "blue")
            base_name = os.path.basename(file_path)

            if rename_ectd:
                name, ext = os.path.splitext(base_name)
                name = name.lower().replace(" ", "-")
                name = re.sub(r'[^a-z0-9_-]', '', name)
                if not name: name = f"doc_{row_idx + 1:03d}"
                base_name = f"{name}{ext.lower()}"

            out_path = os.path.join(self.output_dir, base_name)
            success, msg = PDFProcessor.process_document(file_path, out_path, self.options)
            color = "green" if success else "red"
            self.progress.emit(row_idx, msg, color)

        self.all_completed.emit()


# ==========================================
# 线程二：IO 交互工作线程 (负责专职导入与导出)
# ==========================================
class IOActionWorker(QThread):
    progress = Signal(int, str, str)
    finished_action = Signal()

    def __init__(self, mode, file_list, target_dir, output_dir=None):
        super().__init__()
        self.mode = mode  # export_bookmarks, import_bookmarks, etc.
        self.file_list = file_list
        self.target_dir = target_dir
        self.output_dir = output_dir  # 仅导入时使用

    def run(self):
        for row_idx, file_path in enumerate(self.file_list):
            base_name = os.path.basename(file_path)
            name_no_ext, _ = os.path.splitext(base_name)

            self.progress.emit(row_idx, "⏳ 数据提取/注入中...", "blue")
            success, msg = False, ""

            try:
                if self.mode == 'export_bookmarks':
                    csv_path = os.path.join(self.target_dir, f"{name_no_ext}_bookmarks.csv")
                    PDFProcessor.export_bookmarks(file_path, csv_path)
                    success, msg = True, "✅ 导出书签成功"

                elif self.mode == 'import_bookmarks':
                    csv_path = os.path.join(self.target_dir, f"{name_no_ext}_bookmarks.csv")
                    if not os.path.exists(csv_path):
                        success, msg = False, "⚠️ 未找到匹配的CSV"
                    else:
                        out_pdf = os.path.join(self.output_dir, base_name)
                        PDFProcessor.import_bookmarks(file_path, csv_path, out_pdf)
                        success, msg = True, "✅ 导入书签成功"

                elif self.mode == 'export_links':
                    json_path = os.path.join(self.target_dir, f"{name_no_ext}_links.json")
                    PDFProcessor.export_links(file_path, json_path)
                    success, msg = True, "✅ 导出链接成功"

                elif self.mode == 'import_links':
                    json_path = os.path.join(self.target_dir, f"{name_no_ext}_links.json")
                    if not os.path.exists(json_path):
                        success, msg = False, "⚠️ 未找到匹配的JSON"
                    else:
                        out_pdf = os.path.join(self.output_dir, base_name)
                        PDFProcessor.import_links(file_path, json_path, out_pdf)
                        success, msg = True, "✅ 导入链接成功"

            except Exception as e:
                success, msg = False, f"❌ 失败: {str(e)}"

            color = "green" if success else "red"
            if "未找到匹配" in msg:
                color = "orange"  # 使用特殊颜色警告缺失文件
            self.progress.emit(row_idx, msg, color)

        self.finished_action.emit()


# ==========================================
# 主控制器：响应界面操作并调度业务逻辑
# ==========================================
class MainController:
    def __init__(self, view):
        self.view = view
        self.loaded_files = set()
        self.file_list = []
        self.process_logs = []
        self.setup_connections()

    def setup_connections(self):
        self.view.add_folder_btn.clicked.connect(self.open_folder_dialog)
        self.view.drop_zone.files_dropped.connect(self.process_dropped_paths)
        self.view.btn_clear.clicked.connect(self.clear_table)
        self.view.drop_zone.mousePressEvent = lambda event: self.open_file_dialog()

        self.view.btn_start.clicked.connect(self.start_processing)
        self.view.btn_log.clicked.connect(self.show_log_dialog)

        # 绑定高级数据交换 (IO) 按钮事件
        self.view.btn_export_bookmarks.clicked.connect(lambda: self.handle_io_action('export_bookmarks'))
        self.view.btn_import_bookmarks.clicked.connect(lambda: self.handle_io_action('import_bookmarks'))
        self.view.btn_export_links.clicked.connect(lambda: self.handle_io_action('export_links'))
        self.view.btn_import_links.clicked.connect(lambda: self.handle_io_action('import_links'))

        self.setup_exclusive_options()

    def setup_exclusive_options(self):
        cb_a4 = self.view.all_checkboxes.get("一键批量将页面切换成A4")
        cb_letter = self.view.all_checkboxes.get("一键批量将页面切换成Letter")
        if cb_a4 and cb_letter:
            cb_a4.toggled.connect(lambda checked: cb_letter.setChecked(False) if checked else None)
            cb_letter.toggled.connect(lambda checked: cb_a4.setChecked(False) if checked else None)

    # =============== 数据载入区域 ===============
    def open_folder_dialog(self):
        folder_path = QFileDialog.getExistingDirectory(self.view, "选择包含 PDF 的文件夹")
        if folder_path: self.process_dropped_paths([folder_path])

    def open_file_dialog(self):
        file_paths, _ = QFileDialog.getOpenFileNames(self.view, "选择 PDF 文件", filter="PDF Files (*.pdf)")
        if file_paths: self.process_dropped_paths(file_paths)

    def process_dropped_paths(self, paths):
        pdf_files = []
        for path_str in paths:
            path_obj = Path(path_str)
            if path_obj.is_file() and path_obj.suffix.lower() == '.pdf':
                pdf_files.append(path_obj)
            elif path_obj.is_dir():
                pdf_files.extend(path_obj.rglob('*.pdf'))
        self.add_files_to_table(pdf_files)

    def add_files_to_table(self, file_paths):
        for path_obj in file_paths:
            str_path = str(path_obj)
            if str_path in self.loaded_files:
                if str_path in self.file_list:
                    row_idx = self.file_list.index(str_path)
                    self.view.update_table_row_status(row_idx, "等待中", Qt.darkGray)
                continue
            self.loaded_files.add(str_path)
            self.file_list.append(str_path)
            self.view.add_table_row(f"📄 {path_obj.name}", str_path, "等待中")
        self.update_ui_counters()

    def clear_table(self):
        self.loaded_files.clear()
        self.file_list.clear()
        self.process_logs.clear()
        self.view.clear_table_ui()
        self.update_ui_counters()

    def update_ui_counters(self):
        count = len(self.loaded_files)
        self.view.update_counters_ui(count)

    # =============== 高级数据 IO 操作分发器 ===============
    def handle_io_action(self, mode):
        if not self.file_list:
            QMessageBox.warning(self.view, "提示", "待处理列表为空，请先导入 PDF 文件！")
            return

        is_export = "export" in mode
        data_type = "CSV" if "bookmarks" in mode else "JSON"
        action_name = "导出" if is_export else "导入"

        dir_path = QFileDialog.getExistingDirectory(self.view, f"请选择 {data_type} 数据{action_name}的目录")
        if not dir_path:
            return

        # 若是导入操作，需要额外设立一个新文件夹存放处理后的 PDF
        out_dir = None
        if not is_export:
            first_file = Path(self.file_list[0])
            out_dir = first_file.parent / f"RATools_{action_name}完成"
            out_dir.mkdir(exist_ok=True)
            out_dir = str(out_dir)

        # 启动单独的 IO 线程处理
        self.io_worker = IOActionWorker(mode, self.file_list, dir_path, out_dir)
        self.io_worker.progress.connect(self.update_processing_status)
        self.io_worker.finished_action.connect(
            lambda: QMessageBox.information(self.view, "操作完成", f"批量{action_name}任务已结束！"))
        self.io_worker.start()

    # =============== 核心批量处理流程控制 ===============
    def start_processing(self):
        if not self.file_list:
            QMessageBox.warning(self.view, "提示", "待处理列表为空，请先导入 PDF 文件！")
            return

        selected_options = self.view.get_selected_options()
        if not selected_options:
            QMessageBox.warning(self.view, "提示", "请在右侧设置面板至少勾选一项处理规则！")
            return

        first_file = Path(self.file_list[0])
        out_dir = first_file.parent / "RATools_Output"
        out_dir.mkdir(exist_ok=True)

        self.view.btn_start.setEnabled(False)
        self.view.btn_start.setText("▶ 处理中...")

        self.worker = ProcessWorker(self.file_list, selected_options, str(out_dir))
        self.worker.progress.connect(self.update_processing_status)
        self.worker.all_completed.connect(self.processing_finished)
        self.worker.start()

    def update_processing_status(self, row_idx, status_msg, color_name):
        color_map = {
            "blue": Qt.blue,
            "green": Qt.darkGreen,
            "red": Qt.red,
            "orange": Qt.darkYellow  # 专为缺失的匹配文件使用醒目的橘黄色
        }
        qt_color = color_map.get(color_name, Qt.black)
        self.view.update_table_row_status(row_idx, status_msg, qt_color)

        if not status_msg.startswith("⏳"):
            time_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            file_path = self.file_list[row_idx]
            file_name = os.path.basename(file_path)
            self.process_logs.append((time_str, file_name, file_path, status_msg))

    def processing_finished(self):
        self.view.btn_start.setEnabled(True)
        self.view.btn_start.setText("▶ 开始批量处理")
        QMessageBox.information(self.view, "处理完成",
                                "所有的 PDF 文件常规处理完毕！\n您可以前往 RATools_Output 文件夹查看结果。")

    # =============== 日志查看与导出业务 ===============
    def show_log_dialog(self):
        from view import LogDialog
        dialog = LogDialog(self.view)

        if not self.process_logs:
            dialog.text_edit.setPlainText("暂无处理日志。请先添加 PDF 文件并进行操作。")
            dialog.btn_export.setEnabled(False)
        else:
            lines = []
            for log in self.process_logs:
                lines.append(f"[{log[0]}] {log[1]}\n   ↳ 路径: {log[2]}\n   ↳ 结果: {log[3]}")
            dialog.text_edit.setPlainText("\n\n".join(lines))
            dialog.btn_export.clicked.connect(lambda: self.export_logs_to_csv(dialog))

        dialog.exec()

    def export_logs_to_csv(self, dialog):
        save_path, _ = QFileDialog.getSaveFileName(self.view, "导出处理日志", "RATools_处理日志.csv",
                                                   "CSV Files (*.csv)")
        if not save_path:
            return

        try:
            with open(save_path, 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f)
                writer.writerow(["处理时间", "文件名", "文件路径", "处理结果"])
                for log in self.process_logs:
                    writer.writerow(log)

            QMessageBox.information(self.view, "成功", f"日志已成功导出至：\n{save_path}")
            dialog.accept()
        except Exception as e:
            QMessageBox.critical(self.view, "错误", f"导出日志失败：{str(e)}")