import os
import re
import platform
import subprocess
from pathlib import Path
from datetime import datetime
from PySide6.QtWidgets import QFileDialog
from PySide6.QtCore import QObject, QThread, Signal, Qt
from PySide6.QtGui import QColor

from pdf_processor import PDFProcessor
from view import LogDialog


class ProcessWorker(QThread):
    """
    后台处理线程：负责核心的 PDF 批量规则应用，防止 UI 卡死
    """
    progress = Signal(int, str, str)  # row_index, status_text, log_message
    finished_all = Signal(str)  # summary
    error = Signal(str)  # error_msg

    def __init__(self, files, options, output_dir, common_base="", overwrite_original=False):
        super().__init__()
        self.files = files
        self.options = options
        self.output_dir = output_dir
        self.common_base = common_base
        self.overwrite_original = overwrite_original

    def run(self):
        try:
            success_count = 0
            rename_ectd = "文件名修改为符合电子申报/eCTD要求的格式" in self.options

            for i, file_path in enumerate(self.files):
                base_name = os.path.basename(file_path)
                self.progress.emit(i, "正在处理...", f"[{datetime.now().strftime('%H:%M:%S')}] 开始处理: {base_name}")

                # eCTD 命名合规处理
                if rename_ectd:
                    name, ext = os.path.splitext(base_name)
                    name = name.lower().replace(" ", "-")
                    name = re.sub(r'[^a-z0-9_-]', '', name)
                    if not name: name = f"doc_{i + 1:03d}"
                    base_name = f"{name}{ext.lower()}"

                # 决定输出路径 (支持保留原有文件夹层级结构)
                if self.overwrite_original:
                    out_path = file_path + ".tmp_overwrite.pdf"
                else:
                    if self.common_base:
                        # 计算当前文件相对于公共根目录的相对路径
                        file_dir = os.path.dirname(os.path.abspath(file_path))
                        rel_dir = os.path.relpath(file_dir, self.common_base)
                        if rel_dir == '.':
                            target_dir = self.output_dir
                        else:
                            target_dir = os.path.join(self.output_dir, rel_dir)
                    else:
                        # 降级方案：如果不在同盘符，直接扁平化输出到目标文件夹
                        target_dir = self.output_dir

                    # 自动创建不存在的层级文件夹
                    os.makedirs(target_dir, exist_ok=True)
                    out_path = os.path.join(target_dir, base_name)

                # 精准调用 pdf_processor.py 中的静态方法
                success, msg = PDFProcessor.process_document(file_path, out_path, self.options)

                # 覆盖原文件逻辑
                if success and self.overwrite_original:
                    try:
                        os.replace(out_path, file_path)
                        out_path = file_path
                    except Exception as e:
                        success = False
                        msg = f"覆盖原文件失败: {str(e)}"
                        if os.path.exists(out_path):
                            os.remove(out_path)

                if success:
                    success_count += 1
                    status = "处理完成"
                else:
                    status = "处理失败"

                self.progress.emit(i, status, f"   ↳ 结果: {msg}")

            summary = f"处理结束。共成功处理 {success_count} / {len(self.files)} 个文件。"
            self.finished_all.emit(summary)

        except Exception as e:
            self.error.emit(str(e))


class IOActionWorker(QThread):
    """
    高级 IO 操作后台线程：处理书签、链接等需要长时间读写的批量导入/导出操作
    """
    progress = Signal(int, str, str)
    finished_action = Signal(str)
    error_action = Signal(str)

    def __init__(self, action_type, files, target_dir, output_dir=None):
        super().__init__()
        self.action_type = action_type
        self.files = files
        self.target_dir = target_dir
        self.output_dir = output_dir

    def run(self):
        try:
            for row_idx, file_path in enumerate(self.files):
                base_name = os.path.basename(file_path)
                name_no_ext, _ = os.path.splitext(base_name)

                self.progress.emit(row_idx, "正在执行...",
                                   f"[{datetime.now().strftime('%H:%M:%S')}] 正在处理: {base_name}")
                success, msg = False, ""

                # 精准调用 pdf_processor.py 中的底层数据交换静态方法
                if self.action_type == 'export_bookmarks':
                    csv_path = os.path.join(self.target_dir, f"{name_no_ext}_bookmarks.csv")
                    PDFProcessor.export_bookmarks(file_path, csv_path)
                    success, msg = True, "✅ 导出书签成功"

                elif self.action_type == 'import_bookmarks':
                    csv_path = os.path.join(self.target_dir, f"{name_no_ext}_bookmarks.csv")
                    if not os.path.exists(csv_path):
                        success, msg = False, "⚠️ 未找到匹配的CSV文件"
                    else:
                        out_pdf = os.path.join(self.output_dir, base_name)
                        PDFProcessor.import_bookmarks(file_path, csv_path, out_pdf)
                        success, msg = True, "✅ 导入书签成功"

                elif self.action_type == 'export_links':
                    json_path = os.path.join(self.target_dir, f"{name_no_ext}_links.json")
                    PDFProcessor.export_links(file_path, json_path)
                    success, msg = True, "✅ 导出链接成功"

                elif self.action_type == 'import_links':
                    json_path = os.path.join(self.target_dir, f"{name_no_ext}_links.json")
                    if not os.path.exists(json_path):
                        success, msg = False, "⚠️ 未找到匹配的JSON文件"
                    else:
                        out_pdf = os.path.join(self.output_dir, base_name)
                        PDFProcessor.import_links(file_path, json_path, out_pdf)
                        success, msg = True, "✅ 导入链接成功"

                status = "操作成功" if success else "操作失败"
                if "未找到匹配" in msg:
                    status = "未匹配跳过"
                self.progress.emit(row_idx, status, f"   ↳ 结果: {msg}")

            action_name = "导出" if "export" in self.action_type else "导入"
            self.finished_action.emit(f"批量高级 '{action_name}' 任务执行完成。")

        except Exception as e:
            self.error_action.emit(str(e))


class MainController(QObject):
    def __init__(self, view):
        super().__init__()
        self.view = view
        self.loaded_files = []
        self.process_logs = ""
        self.last_output_dir = ""

        self.setup_connections()

    def setup_connections(self):
        # 基础文件列表交互：绑定拖拽与点击添加文件
        self.view.drop_zone.files_dropped.connect(self.add_files)
        self.view.drop_zone.mousePressEvent = self.open_file_dialog
        self.view.add_folder_btn.clicked.connect(self.add_folder)
        self.view.btn_clear.clicked.connect(self.clear_list)

        # 核心功能交互
        self.view.btn_start.clicked.connect(self.start_processing)
        self.view.btn_log.clicked.connect(self.show_log_dialog)

        # 高级数据 IO 交互
        self.view.btn_export_bookmarks.clicked.connect(lambda: self.handle_io_action('export_bookmarks'))
        self.view.btn_import_bookmarks.clicked.connect(lambda: self.handle_io_action('import_bookmarks'))
        self.view.btn_export_links.clicked.connect(lambda: self.handle_io_action('export_links'))
        self.view.btn_import_links.clicked.connect(lambda: self.handle_io_action('import_links'))

        # 互斥选项配置：A4 和 Letter 只能选其一
        self.setup_exclusive_options()

    def setup_exclusive_options(self):
        cb_a4 = self.view.all_checkboxes.get("一键批量将页面切换成A4")
        cb_letter = self.view.all_checkboxes.get("一键批量将页面切换成Letter")
        if cb_a4 and cb_letter:
            cb_a4.toggled.connect(lambda checked: cb_letter.setChecked(False) if checked else None)
            cb_letter.toggled.connect(lambda checked: cb_a4.setChecked(False) if checked else None)

    def open_file_dialog(self, event):
        if event.button() == Qt.LeftButton:
            file_paths, _ = QFileDialog.getOpenFileNames(
                self.view,
                "选择 PDF 文件",
                "",
                "PDF Files (*.pdf);;All Files (*)"
            )
            if file_paths:
                self.add_files(file_paths)

    def add_files(self, paths):
        new_count = 0

        # 提取所有实际的 pdf 路径（支持解析传入的文件夹）
        valid_pdf_paths = []
        for p in paths:
            if os.path.isfile(p) and p.lower().endswith('.pdf'):
                valid_pdf_paths.append(p)
            elif os.path.isdir(p):
                for root, _, files in os.walk(p):
                    for file in files:
                        if file.lower().endswith('.pdf'):
                            valid_pdf_paths.append(os.path.join(root, file))

        for path in valid_pdf_paths:
            if path not in self.loaded_files:
                self.loaded_files.append(path)
                name = os.path.basename(path)
                self.view.add_table_row(name, path, "等待处理")
                new_count += 1

        self.view.update_counters_ui(len(self.loaded_files))

        if new_count == 0 and paths:
            self.view.show_info_message("ℹ️ 提示", "添加的文件或文件夹中没有新的 PDF 文件，或文件已存在于列表中。")

    def add_folder(self):
        folder_path = QFileDialog.getExistingDirectory(self.view, "选择包含 PDF 的文件夹")
        if folder_path:
            # 既然 add_files 现在支持解析文件夹了，直接复用逻辑即可
            self.add_files([folder_path])

    def clear_list(self):
        if not self.loaded_files:
            return

        if self.view.show_confirm_message("🗑️ 确认清空", "您确定要清空待处理列表吗？"):
            self.loaded_files.clear()
            self.view.clear_table_ui()
            self.view.update_counters_ui(0)
            self.process_logs = ""

    def start_processing(self):
        if not self.loaded_files:
            self.view.show_warning_message("⚠️ 警告", "请至少添加一个 PDF 文件！")
            return

        selected_options = self.view.get_selected_options()
        if not selected_options:
            self.view.show_warning_message("⚠️ 警告", "请至少在右侧勾选一个处理规则！")
            return

        overwrite_cb = self.view.all_checkboxes.get("覆盖原始文件 (不推荐)")
        overwrite_original = overwrite_cb.isChecked() if overwrite_cb else False

        out_dir = ""
        common_base = ""

        if overwrite_original:
            if not self.view.show_confirm_message("⚠️ 危险操作确认",
                                                  "您勾选了【覆盖原始文件】。\n此操作不可逆，强烈建议您在操作前备份文件！\n\n是否继续？"):
                return
        else:
            # 弹出对话框，让用户主动选择想要保存的根目录
            user_selected_dir = QFileDialog.getExistingDirectory(self.view, "选择输出文件保存的根目录")
            if not user_selected_dir:
                return  # 用户取消了选择

            out_dir = os.path.join(user_selected_dir, "RATools_Output")
            self.last_output_dir = out_dir

            try:
                # 提取所有文件所在的绝对路径目录，并计算它们共有的最长路径前缀
                dirs = [os.path.dirname(os.path.abspath(f)) for f in self.loaded_files]
                common_base = os.path.commonpath(dirs)
            except ValueError:
                # 异常处理：例如文件分别位于 Windows 的 C 盘和 D 盘，没有共同路径
                common_base = ""

        self.view.btn_start.setEnabled(False)
        self.view.btn_start.setText("处理中...")

        self.worker = ProcessWorker(self.loaded_files, selected_options, out_dir, common_base, overwrite_original)
        self.worker.progress.connect(self.update_progress)
        self.worker.finished_all.connect(self.processing_finished)
        self.worker.error.connect(self.processing_error)
        self.worker.start()

    def handle_io_action(self, action_type):
        if not self.loaded_files:
            self.view.show_warning_message("⚠️ 警告", "请先添加目标 PDF 文件！")
            return

        is_export = "export" in action_type
        data_type = "CSV" if "bookmarks" in action_type else "JSON"
        action_name = "导出" if is_export else "导入"

        dir_path = QFileDialog.getExistingDirectory(self.view, f"请选择 {data_type} 数据{action_name}的目录")
        if not dir_path:
            return

        out_dir = None
        if not is_export:
            first_file = Path(self.loaded_files[0])
            out_dir_path = first_file.parent / f"RATools_{action_name}完成"
            out_dir_path.mkdir(exist_ok=True)
            out_dir = str(out_dir_path)

        self.io_worker = IOActionWorker(action_type, self.loaded_files, dir_path, out_dir)
        self.io_worker.progress.connect(self.update_progress)
        self.io_worker.finished_action.connect(self.on_io_action_finished)
        self.io_worker.error_action.connect(self.on_io_action_error)
        self.io_worker.start()

    def update_progress(self, row_index, status_text, log_msg):
        if status_text in ["处理完成", "操作成功"]:
            color = QColor(16, 185, 129)  # 绿色
        elif status_text in ["处理失败", "操作失败"]:
            color = QColor(239, 68, 68)  # 红色
        elif status_text == "未匹配跳过":
            color = QColor(245, 158, 11)  # 橙黄色警告
        else:
            color = QColor(37, 99, 235)  # 蓝色处理中

        self.view.update_table_row_status(row_index, status_text, color)
        self.process_logs += f"{log_msg}\n"

    def processing_finished(self, summary):
        self.process_logs += f"\n=== 批量处理结束 ===\n{summary}\n"
        self.view.btn_start.setEnabled(True)
        self.view.btn_start.setText("▶ 开始批量处理")

        self.view.show_success_message("✅ 处理完成", "所有 PDF 文件的批量处理任务已结束！")

        # 如果勾选了自动打开输出文件夹，利用新记录的 self.last_output_dir 进行跳转
        auto_open_cb = self.view.all_checkboxes.get("处理完成后自动打开输出文件夹")
        if auto_open_cb and auto_open_cb.isChecked() and self.loaded_files:
            overwrite_cb = self.view.all_checkboxes.get("覆盖原始文件 (不推荐)")
            if overwrite_cb and not overwrite_cb.isChecked():
                if hasattr(self, 'last_output_dir') and self.last_output_dir and os.path.exists(self.last_output_dir):
                    self._open_directory(self.last_output_dir)

    def processing_error(self, error_msg):
        self.process_logs += f"\n[致命错误] {error_msg}\n"
        self.view.btn_start.setEnabled(True)
        self.view.btn_start.setText("▶ 开始批量处理")
        self.view.show_error_message("❌ 处理异常", f"处理过程中发生错误：\n{error_msg}")

    def on_io_action_finished(self, result_msg):
        self.process_logs += f"\n{result_msg}\n"
        self.view.show_success_message("✅ 操作成功", result_msg)

    def on_io_action_error(self, error_msg):
        self.process_logs += f"\n[IO错误] {error_msg}\n"
        self.view.show_error_message("❌ 操作失败", error_msg)

    def show_log_dialog(self):
        if not hasattr(self, 'log_dialog'):
            self.log_dialog = LogDialog(self.view)
            self.log_dialog.btn_export.clicked.connect(self.export_logs)

        self.log_dialog.text_edit.setText(self.process_logs if self.process_logs else "暂无处理日志...")
        self.log_dialog.show()
        self.log_dialog.raise_()
        self.log_dialog.activateWindow()

    def export_logs(self):
        if not self.process_logs:
            self.view.show_warning_message("⚠️ 提示", "目前暂无任何日志可供导出！")
            return

        file_path, _ = QFileDialog.getSaveFileName(self.view, "导出处理日志", "process_logs.txt",
                                                   "Text Files (*.txt);;All Files (*)")
        if file_path:
            try:
                with open(file_path, "w", encoding="utf-8") as f:
                    f.write(self.process_logs)
                self.view.show_success_message("✅ 导出成功", "处理日志已成功保存！")
            except Exception as e:
                self.view.show_error_message("❌ 导出失败", f"文件保存失败：\n{str(e)}")

    def _open_directory(self, dir_path):
        """跨平台打开目录"""
        sys_plat = platform.system()
        try:
            if sys_plat == "Windows":
                os.startfile(dir_path)
            elif sys_plat == "Darwin":
                subprocess.Popen(["open", dir_path])
            else:
                subprocess.Popen(["xdg-open", dir_path])
        except Exception as e:
            self.process_logs += f"\n[警告] 自动打开文件夹失败：{str(e)}\n"