"""日志查看与导出子控制器：日志对话框、处理日志 CSV/TXT 导出、预检结果 CSV 导出。"""

import csv
import os
from datetime import datetime

from PySide6.QtCore import QObject
from PySide6.QtWidgets import QFileDialog

from ratools_pdf.controllers.io_actions import common_base_dir
from ratools_pdf.controllers.log_export import _select_log_rows_for_export
from ratools_pdf.ui.dialogs import LogDialog


class LogController(QObject):
    def __init__(self, host, view, parent=None):
        super().__init__(parent)
        self.host = host
        self.view = view
        self.log_dialog = None

    def show_log_dialog(self):
        host = self.host
        if self.log_dialog is None:
            self.log_dialog = LogDialog(self.view)
            self.log_dialog.btn_export.clicked.connect(self.export_logs)
            self.log_dialog.btn_export_precheck.clicked.connect(self.export_precheck_results)

        self.log_dialog.set_log_data(
            host.process_logs if host.process_logs else "暂无处理日志...",
            host.process_log_rows,
        )
        self.log_dialog.btn_export_precheck.setEnabled(bool(host.last_precheck_results))
        self.log_dialog.btn_export_precheck.setToolTip(
            "导出最近一次批量预检结果" if host.last_precheck_results else "请先执行一次批量预检"
        )
        self.log_dialog.show()
        self.log_dialog.raise_()
        self.log_dialog.activateWindow()

    def _default_export_dir(self, prefer_last_output=True):
        host = self.host
        if prefer_last_output and host.last_output_dir and os.path.isdir(host.last_output_dir):
            return host.last_output_dir
        default_output_dir = self.view.settings_dialog.default_output_edit.text().strip()
        if default_output_dir and os.path.isdir(default_output_dir):
            return default_output_dir
        if host.loaded_files:
            return common_base_dir(
                host.loaded_files,
                fallback=os.path.dirname(os.path.abspath(host.loaded_files[0])),
            )
        return ""

    def export_logs(self):
        host = self.host
        if not host.process_logs:
            self.view.show_warning_message("⚠️ 提示", "目前暂无任何日志可供导出！")
            return

        default_dir = self._default_export_dir(prefer_last_output=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        default_filename = f"RATools_process_logs_{timestamp}.csv"
        default_path = os.path.join(default_dir, default_filename) if default_dir else default_filename

        file_path, selected_filter = QFileDialog.getSaveFileName(
            self.view,
            "导出处理日志",
            default_path,
            "CSV Summary (*.csv);;Text Files (*.txt);;All Files (*)"
        )
        if not file_path:
            return
        try:
            export_csv = file_path.lower().endswith('.csv') or selected_filter.startswith("CSV")

            if export_csv and not file_path.lower().endswith('.csv'):
                file_path += '.csv'
            if not export_csv and selected_filter.startswith("Text") and not file_path.lower().endswith('.txt'):
                file_path += '.txt'

            if export_csv:
                with open(file_path, "w", encoding="utf-8-sig", newline="") as f:
                    writer = csv.DictWriter(f, fieldnames=["time", "file_original", "file_output", "status", "success", "duration_sec", "changes"])
                    writer.writeheader()
                    writer.writerows(_select_log_rows_for_export(host.process_log_rows, host.process_logs))
            else:
                with open(file_path, "w", encoding="utf-8") as f:
                    f.write(host.process_logs)
            self.view.show_success_message("✅ 导出成功", "处理日志已成功保存！")
        except Exception as e:
            self.view.show_error_message("❌ 导出失败", f"文件保存失败：\n{str(e)}")

    def export_precheck_results(self):
        host = self.host
        if not host.last_precheck_results:
            self.view.show_warning_message("⚠️ 提示", "请先执行一次批量预检，再导出预检结果。")
            return

        default_dir = self._default_export_dir(prefer_last_output=False)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        default_filename = f"RATools_precheck_results_{timestamp}.csv"
        default_path = os.path.join(default_dir, default_filename) if default_dir else default_filename

        file_path, _selected_filter = QFileDialog.getSaveFileName(
            self.view,
            "导出预检结果",
            default_path,
            "CSV Files (*.csv);;All Files (*)",
        )
        if not file_path:
            return
        if not file_path.lower().endswith(".csv"):
            file_path += ".csv"

        try:
            fieldnames = [
                "file_name",
                "file_path",
                "status",
                "suggestions",
                "suggestion_ids",
                "error",
                "font_summary",
                "font_details",
                "annotation_summary",
                "annotation_details",
                "broken_reference_summary",
                "broken_reference_details",
            ]
            with open(file_path, "w", encoding="utf-8-sig", newline="") as f:
                writer = csv.DictWriter(f, fieldnames=fieldnames)
                writer.writeheader()
                rows = []
                for row in host.last_precheck_results:
                    export_row = {key: row.get(key, "") for key in fieldnames}
                    rows.append(export_row)
                writer.writerows(rows)
            self.view.show_success_message("✅ 导出成功", "预检结果已成功保存！")
        except Exception as e:
            self.view.show_error_message("❌ 导出失败", f"文件保存失败：\n{str(e)}")
