"""批量预检子控制器：预检执行、建议应用、建议文件筛选。

通过 host（MainController）访问共享状态：处理 worker 忙碌判断、队列、
日志缓冲、进度路由与"仅处理建议文件"的处理入口。
"""

from PySide6.QtCore import QObject

from ratools_pdf.controllers.workers import PreCheckWorker


class PrecheckController(QObject):
    def __init__(self, host, view, parent=None):
        super().__init__(parent)
        self.host = host
        self.view = view
        self.precheck_worker = None
        self.precheck_files = []
        self.last_precheck_results = []
        self.last_precheck_suggested_files = []
        self.precheck_result_current = False

    def is_running(self):
        return bool(self.precheck_worker and self.precheck_worker.isRunning())

    def mark_stale(self):
        """队列变化或处理完成后，上一轮预检结果不再可信。"""
        self.precheck_result_current = False
        self.view.btn_precheck.setProperty("precheckResultCurrent", False)
        self.view.btn_precheck.show()

    def start_precheck(self):
        host, view = self.host, self.view
        if host.worker and host.worker.isRunning():
            view.show_warning_message("⚠️ 正在处理", "批量处理进行中，无法执行预检。")
            return
        if host._is_detection_running():
            view.show_warning_message("⚠️ 正在检测", "检测进行中，请稍候再执行预检。")
            return
        if self.is_running():
            return
        if not host.loaded_files:
            view.show_warning_message("⚠️ 警告", "请至少添加一个 PDF 文件！")
            return

        self.precheck_files = list(host.loaded_files)
        self.last_precheck_results = []
        self.last_precheck_suggested_files = []
        self.precheck_result_current = False
        view.btn_precheck.setProperty("precheckResultCurrent", False)
        view.btn_apply_precheck.setProperty("hasPrecheckSuggestions", False)
        view.btn_apply_precheck.hide()
        view.btn_process_precheck_suggested.setProperty("hasPrecheckSuggestedFiles", False)
        view.btn_process_precheck_suggested.hide()
        host.process_logs += f"\n{'=' * 56}\n批量预检开始\n{'=' * 56}\n"
        view.btn_precheck.setEnabled(False)
        view.btn_precheck.setText("预检中...")
        view.btn_precheck.setProperty("precheckMode", True)
        view.btn_start.setEnabled(False)

        self.precheck_worker = PreCheckWorker(self.precheck_files)
        self.precheck_worker.progress.connect(host.update_progress)
        self.precheck_worker.result_ready.connect(self._record_precheck_result)
        self.precheck_worker.finished_precheck.connect(self.precheck_finished)
        self.precheck_worker.error_precheck.connect(self.precheck_error)
        self.precheck_worker.finished.connect(self.precheck_worker.deleteLater)
        self.precheck_worker.start()

    def _record_precheck_result(self, row):
        self.last_precheck_results.append(dict(row))
        suggestion_ids = str(row.get("suggestion_ids", "") or "")
        if suggestion_ids.strip():
            self.view.btn_apply_precheck.setProperty("hasPrecheckSuggestions", True)
        if row.get("status") == "建议处理" and row.get("file_path") and suggestion_ids.strip():
            file_path = row.get("file_path")
            if file_path not in self.last_precheck_suggested_files:
                self.last_precheck_suggested_files.append(file_path)
            self.view.btn_process_precheck_suggested.setProperty("hasPrecheckSuggestedFiles", True)
        self.view.refresh_selection_summary()

    def apply_precheck_suggestions(self):
        view = self.view
        suggested_options = []
        seen = set()
        for row in self.last_precheck_results:
            raw_ids = str(row.get("suggestion_ids", "") or "")
            for option_id in [item.strip() for item in raw_ids.split(",") if item.strip()]:
                if option_id in seen:
                    continue
                if option_id in view.all_checkboxes:
                    suggested_options.append(option_id)
                    seen.add(option_id)

        if not suggested_options:
            view.show_warning_message("⚠️ 无建议项", "最近一次预检没有可自动应用的建议处理项。")
            return

        view.is_applying_preset = True
        try:
            for option_id in suggested_options:
                view.all_checkboxes[option_id].setChecked(True)
        finally:
            view.is_applying_preset = False

        view.active_preset_key = None
        view._set_preset_button_state(None)
        view.custom_selection_before_preset = set(view.get_selected_options())
        view.persist_all_settings()
        view.refresh_selection_summary()
        view.show_success_message("✅ 已应用", f"已自动勾选 {len(suggested_options)} 条预检建议规则。")

    def start_precheck_suggested_processing(self):
        host = self.host
        suggested_files = [path for path in self.last_precheck_suggested_files if path in host.loaded_files]
        if not suggested_files:
            self.view.show_warning_message("⚠️ 无建议文件", "最近一次预检没有可处理的建议文件。")
            return
        host.start_processing(processing_files=suggested_files)

    def _reset_precheck_ui(self):
        """预检结束/异常后的按钮与状态复位（finished 与 error 共用）。"""
        self.view.btn_precheck.setText("🔎 预检")
        self.view.btn_precheck.setProperty("precheckMode", False)
        self.view.btn_start.setEnabled(True)
        self.precheck_files = []
        self.precheck_worker = None

    def precheck_finished(self, summary):
        self.host.process_logs += f"\n{'=' * 56}\n批量预检结束\n{summary}\n{'=' * 56}\n"
        self._reset_precheck_ui()
        self.precheck_result_current = True
        self.view.btn_precheck.setProperty("precheckResultCurrent", True)
        self.view.btn_precheck.hide()
        self.view.refresh_selection_summary()
        self.view.show_info_message("🔎 预检完成", summary)

    def precheck_error(self, error_msg):
        self.host.process_logs += f"\n{'!' * 56}\n[预检错误] {error_msg}\n{'!' * 56}\n"
        self._reset_precheck_ui()
        self.view.refresh_selection_summary()
        self.view.show_error_message("❌ 预检失败", f"预检过程中发生错误：\n{error_msg}")
