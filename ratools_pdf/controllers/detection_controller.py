"""文档检测子控制器：批注 / 失效引用文本的只读扫描。

通过 host（MainController）访问共享状态：忙碌互斥、队列、日志缓冲与
进度路由。检测本身不修改任何 PDF。
"""

from PySide6.QtCore import QObject

from ratools_pdf.controllers.workers import DetectionWorker


class DetectionController(QObject):
    KIND_MAP = {
        "annotations": "annotation",
        "broken_refs": "broken_reference",
    }

    KIND_TITLES = {
        "annotation": "批注检测",
        "broken_reference": "失效引用/链接文本检测",
    }

    def __init__(self, host, view, parent=None):
        super().__init__(parent)
        self.host = host
        self.view = view
        self.detection_worker = None
        self.detection_files = []
        self.last_detection_results = []

    def is_running(self):
        return bool(self.detection_worker and self.detection_worker.isRunning())

    def start_detection(self, ui_kind):
        host, view = self.host, self.view
        detection_kind = self.KIND_MAP.get(ui_kind, "annotation")
        kind_title = self.KIND_TITLES[detection_kind]

        if host.worker and host.worker.isRunning():
            view.show_warning_message("⚠️ 正在处理", f"批量处理进行中，无法执行{kind_title}。")
            return
        if host._is_precheck_running():
            view.show_warning_message("⚠️ 正在预检", f"预检进行中，无法执行{kind_title}。")
            return
        if self.is_running():
            view.show_warning_message("⚠️ 正在检测", "已有检测任务进行中，请稍候。")
            return
        if not host.loaded_files:
            view.show_warning_message("⚠️ 警告", "请至少添加一个 PDF 文件！")
            return

        self.detection_files = list(host.loaded_files)
        self.last_detection_results = []
        host.process_logs += f"\n{'=' * 56}\n{kind_title}开始\n{'=' * 56}\n"

        view.btn_detect_annotations.setEnabled(False)
        view.btn_detect_broken_refs.setEnabled(False)

        self.detection_worker = DetectionWorker(self.detection_files, detection_kind)
        self.detection_worker.progress.connect(host.update_progress)
        self.detection_worker.result_ready.connect(self._record_detection_result)
        self.detection_worker.finished_detection.connect(
            lambda summary, results, title=kind_title: self.detection_finished(summary, results, title)
        )
        self.detection_worker.error_detection.connect(
            lambda error_msg, title=kind_title: self.detection_error(error_msg, title)
        )
        self.detection_worker.finished.connect(self.detection_worker.deleteLater)
        self.detection_worker.start()

    def _record_detection_result(self, row):
        self.last_detection_results.append(dict(row))

    def _restore_detection_buttons(self):
        self.view.btn_detect_annotations.setEnabled(True)
        self.view.btn_detect_broken_refs.setEnabled(True)

    def _reset_state(self):
        self._restore_detection_buttons()
        self.detection_files = []
        self.detection_worker = None

    def detection_finished(self, summary, results, kind_title):
        self.host.process_logs += f"\n{'=' * 56}\n{kind_title}结束\n{summary}\n{'=' * 56}\n"
        self._reset_state()

        hit_rows = [row for row in results if row.get("status") == "发现问题"]
        failed_rows = [row for row in results if row.get("status") == "检测失败"]

        message_lines = [summary, ""]
        if hit_rows:
            message_lines.append("发现问题的文件：")
            for row in hit_rows:
                detail = row.get("summary", "") or ""
                if row.get("details"):
                    detail = f"{detail}（{row['details']}）" if detail else row["details"]
                message_lines.append(f"• {row.get('file_name', '')}：{detail}")
        else:
            message_lines.append("未在任何文件中发现相关内容。")
        if failed_rows:
            message_lines.append("")
            message_lines.append("检测失败的文件：")
            for row in failed_rows:
                message_lines.append(f"• {row.get('file_name', '')}：{row.get('error', '')}")

        self.view.show_info_message(f"🔍 {kind_title}完成", "\n".join(message_lines))

    def detection_error(self, error_msg, kind_title):
        self.host.process_logs += f"\n{'!' * 56}\n[{kind_title}错误] {error_msg}\n{'!' * 56}\n"
        self._reset_state()
        self.view.show_error_message(f"❌ {kind_title}失败", f"检测过程中发生错误：\n{error_msg}")
