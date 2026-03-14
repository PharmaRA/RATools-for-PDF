import os
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

            # 调用核心处理引擎
            success, msg = PDFProcessor.process_initial_view(file_path, out_path, self.options)

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
        self.setup_connections()

    def setup_connections(self):
        self.view.add_folder_btn.clicked.connect(self.open_folder_dialog)
        self.view.drop_zone.files_dropped.connect(self.process_dropped_paths)
        self.view.btn_clear.clicked.connect(self.clear_table)
        self.view.drop_zone.mousePressEvent = lambda event: self.open_file_dialog()

        self.view.btn_start.clicked.connect(self.start_processing)

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

        # 【修复重点】：绑定重命名后的信号
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

    def processing_finished(self):
        self.view.btn_start.setEnabled(True)
        self.view.btn_start.setText("▶ 开始批量处理")
        QMessageBox.information(self.view, "处理完成",
                                "所有的 PDF 文件处理完毕！\n您可以前往 RATools_Output 文件夹查看结果。")