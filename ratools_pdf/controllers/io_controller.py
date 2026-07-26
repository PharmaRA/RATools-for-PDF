"""书签/链接批量导入导出子控制器（IO 向导流程）。"""

from pathlib import Path

from PySide6.QtCore import QObject
from PySide6.QtWidgets import QFileDialog

from ratools_pdf.controllers.io_actions import (
    _build_io_preview_rows,
    _io_action_metadata,
    _normalize_io_action_types,
    common_base_dir,
)
from ratools_pdf.controllers.workers import IOActionWorker
from ratools_pdf.ui.dialogs import IODataWizardDialog


class IOController(QObject):
    def __init__(self, host, view, parent=None):
        super().__init__(parent)
        self.host = host
        self.view = view
        self.io_worker = None

    def _loaded_files_common_base(self):
        return common_base_dir(self.host.loaded_files)

    def handle_io_wizard(self, default_data_kind):
        host, view = self.host, self.view
        if not host.loaded_files:
            view.show_warning_message("⚠️ 警告", "请先添加目标 PDF 文件！")
            return

        common_base = self._loaded_files_common_base()

        def build_preview(action_type, dir_path):
            return _build_io_preview_rows(host.loaded_files, action_type, dir_path, common_base)

        dialog = IODataWizardDialog(
            data_kind=default_data_kind,
            file_count=len(host.loaded_files),
            preview_callback=build_preview,
            parent=view,
        )
        if not dialog.exec():
            return

        self.handle_io_action(
            dialog.get_action_types(),
            dir_path=dialog.get_selected_directory(),
            common_base=common_base,
            confirmed=True,
            link_scope=dialog.get_link_scope(),
            link_mode=dialog.get_link_mode(),
        )

    def handle_io_action(self, action_type, dir_path=None, common_base=None, confirmed=False,
                         link_scope="all", link_mode="overwrite"):
        host, view = self.host, self.view
        if not host.loaded_files:
            view.show_warning_message("⚠️ 警告", "请先添加目标 PDF 文件！")
            return

        action_types = _normalize_io_action_types(action_type)
        meta = _io_action_metadata(action_types[0])
        is_export = meta["is_export"]
        data_type = "CSV/JSON" if len(action_types) > 1 else meta["data_type"]
        action_name = meta["action_name"]

        if common_base is None:
            common_base = self._loaded_files_common_base()

        if dir_path is None:
            dir_path = QFileDialog.getExistingDirectory(view, f"请选择 {data_type} 数据{action_name}的目录")
        if not dir_path:
            return

        if confirmed and not is_export:
            rows = _build_io_preview_rows(host.loaded_files, action_types, dir_path, common_base)
            if not any(row["status"] == "已匹配" for row in rows):
                view.show_warning_message("⚠️ 未找到数据文件", f"所选目录中没有匹配的 {data_type} 文件。")
                return

        out_dir = None
        if not is_export:
            first_file = Path(host.loaded_files[0])
            out_dir_path = first_file.parent / f"RATools_{action_name}完成"
            out_dir_path.mkdir(exist_ok=True)
            out_dir = str(out_dir_path)

        self.io_worker = IOActionWorker(
            action_types, host.loaded_files, dir_path, out_dir, common_base,
            link_scope=link_scope, link_mode=link_mode,
        )
        self.io_worker.progress.connect(host.update_progress)
        self.io_worker.finished_action.connect(self.on_io_action_finished)
        self.io_worker.error_action.connect(self.on_io_action_error)
        self.io_worker.start()

    def on_io_action_finished(self, result_msg):
        self.host.process_logs += f"\n{result_msg}\n"
        self.view.show_success_message("✅ 操作完成", result_msg)

    def on_io_action_error(self, error_msg):
        self.host.process_logs += f"\n[IO操作错误] {error_msg}\n"
        self.view.show_error_message("❌ 操作失败", f"批量高级操作执行失败：\n{error_msg}")
