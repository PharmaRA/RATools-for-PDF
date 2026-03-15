import os
import csv
from datetime import datetime
from pathlib import Path
from PySide6.QtWidgets import QFileDialog, QMessageBox
from PySide6.QtCore import Qt, QThread, Signal
from pdf_processor import PDFProcessor


# ==========================================
# 后台工作线程：防止 PDF 处理期间主界面无响应 (ANR)
# ==========================================
class ProcessWorker(QThread):
    progress = Signal(int, str, str)  # 信号: 行号, 状态文本, 颜色名称

    # 【修复重点】：避免使用 'finished' 命名，因为它会与 QThread 内置信号冲突导致执行两次
    all_completed = Signal()

    def __init__(self, file_list, options, output_dir):
        super().__init__()
        self.file_list = file_list
        self.options = options
        self.output_dir = output_dir

    def run(self):
        for row_idx, file_path in enumerate(self.file_list):
            self.progress.emit(row_idx, "⏳ 处理中...", "blue")

            base_name = os.path.basename(file_path)
            out_path = os.path.join(self.output_dir, f"{base_name}")

            # 调用核心处理引擎 (已重命名为统一入口 process_document)
            success, msg = PDFProcessor.process_document(file_path, out_path, self.options)

            color = "green" if success else "red"
            self.progress.emit(row_idx, msg, color)

        # 发送自定义完成信号
        self.all_completed.emit()


# ==========================================
# 主控制器：响应界面操作并调度业务逻辑
# ==========================================
class MainController:
    def __init__(self, view):
        self.view = view
        self.loaded_files = set()
        self.file_list = []  # 保持文件的加入顺序，对应表格里的 row_index
        self.process_logs = []  # 存储处理日志记录
        self.setup_connections()

    def setup_connections(self):
        self.view.add_folder_btn.clicked.connect(self.open_folder_dialog)
        self.view.drop_zone.files_dropped.connect(self.process_dropped_paths)
        self.view.btn_clear.clicked.connect(self.clear_table)
        self.view.drop_zone.mousePressEvent = lambda event: self.open_file_dialog()

        self.view.btn_start.clicked.connect(self.start_processing)
        self.view.btn_log.clicked.connect(self.show_log_dialog)

        # 处理配置选项的互斥联动逻辑
        self.setup_exclusive_options()

    def setup_exclusive_options(self):
        """设置互相冲突的选项联动（如 A4 和 Letter 只能二选一）"""
        cb_a4 = self.view.all_checkboxes.get("一键批量将页面切换成A4")
        cb_letter = self.view.all_checkboxes.get("一键批量将页面切换成Letter")

        if cb_a4 and cb_letter:
            # 勾选 A4 时，自动取消勾选 Letter
            cb_a4.toggled.connect(lambda checked: cb_letter.setChecked(False) if checked else None)
            # 勾选 Letter 时，自动取消勾选 A4
            cb_letter.toggled.connect(lambda checked: cb_a4.setChecked(False) if checked else None)

    def open_folder_dialog(self):
        folder_path = QFileDialog.getExistingDirectory(self.view, "选择包含 PDF 的文件夹")
        if folder_path:
            self.process_dropped_paths([folder_path])

    def open_file_dialog(self):
        file_paths, _ = QFileDialog.getOpenFileNames(self.view, "选择 PDF 文件", filter="PDF Files (*.pdf)")
        if file_paths:
            self.process_dropped_paths(file_paths)

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
                continue
            self.loaded_files.add(str_path)
            self.file_list.append(str_path)
            self.view.add_table_row(f"📄 {path_obj.name}", str_path, "等待中")
        self.update_ui_counters()

    def clear_table(self):
        self.loaded_files.clear()
        self.file_list.clear()
        self.process_logs.clear()  # 清空旧日志
        self.view.clear_table_ui()
        self.update_ui_counters()

    def update_ui_counters(self):
        count = len(self.loaded_files)
        self.view.update_counters_ui(count)

    # ==========================================
    # 核心批量处理流程控制
    # ==========================================
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

        # 绑定重命名后的信号
        self.worker.all_completed.connect(self.processing_finished)

        self.worker.start()

    def update_processing_status(self, row_idx, status_msg, color_name):
        color_map = {
            "blue": Qt.blue,
            "green": Qt.darkGreen,
            "red": Qt.red
        }
        qt_color = color_map.get(color_name, Qt.black)
        self.view.update_table_row_status(row_idx, status_msg, qt_color)

        # 记录最终的处理日志（排除带有⏳的过度状态，只记录成功或失败）
        if not status_msg.startswith("⏳"):
            time_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            file_path = self.file_list[row_idx]
            file_name = os.path.basename(file_path)
            self.process_logs.append((time_str, file_name, file_path, status_msg))

    def processing_finished(self):
        self.view.btn_start.setEnabled(True)
        self.view.btn_start.setText("▶ 开始批量处理")
        QMessageBox.information(self.view, "处理完成",
                                "所有的 PDF 文件处理完毕！\n您可以前往 RATools_Output 文件夹查看结果。")

    # ==========================================
    # 日志查看与导出业务
    # ==========================================
    def show_log_dialog(self):
        from view import LogDialog
        dialog = LogDialog(self.view)

        if not self.process_logs:
            dialog.text_edit.setPlainText("暂无处理日志。请先添加 PDF 文件并点击“开始批量处理”。")
            dialog.btn_export.setEnabled(False)
        else:
            lines = []
            for log in self.process_logs:
                # log: (时间, 文件名, 路径, 结果)
                lines.append(f"[{log[0]}] {log[1]}\n   ↳ 路径: {log[2]}\n   ↳ 结果: {log[3]}")
            dialog.text_edit.setPlainText("\n\n".join(lines))

            # 绑定导出 CSV 事件
            dialog.btn_export.clicked.connect(lambda: self.export_logs_to_csv(dialog))

        dialog.exec()

    def export_logs_to_csv(self, dialog):
        save_path, _ = QFileDialog.getSaveFileName(self.view, "导出处理日志", "RATools_处理日志.csv",
                                                   "CSV Files (*.csv)")
        if not save_path:
            return

        try:
            # 使用 utf-8-sig 编码，确保 Excel 打开 CSV 文件时中文字符不会乱码
            with open(save_path, 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f)
                writer.writerow(["处理时间", "文件名", "文件路径", "处理结果"])
                for log in self.process_logs:
                    writer.writerow(log)

            QMessageBox.information(self.view, "成功", f"日志已成功导出至：\n{save_path}")
            dialog.accept()  # 导出成功后自动关闭日志弹窗
        except Exception as e:
            QMessageBox.critical(self.view, "错误", f"导出日志失败：{str(e)}")