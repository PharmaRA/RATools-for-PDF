"""更新检查子控制器：手动/静默检查、结果展示、发布页跳转。

ENABLE_UPDATE_CHECK 的判断集中在本模块；NoUpdate 变体下所有入口静默返回。
"""

import webbrowser
from datetime import date

from PySide6.QtCore import QObject

from ratools_pdf.config.features import ENABLE_UPDATE_CHECK
from ratools_pdf.controllers.workers import UpdateCheckWorker


class UpdateController(QObject):
    def __init__(self, view, parent=None):
        super().__init__(parent)
        self.view = view
        self.update_worker = None

    # ---- 接线 ----

    def wire_about_dialog(self):
        if not ENABLE_UPDATE_CHECK:
            return
        dialog = getattr(self.view, "about_dialog", None)
        if not dialog or getattr(dialog, "_update_buttons_wired", False):
            return

        dialog.btn_check_updates.clicked.connect(self.check_updates_manually)
        dialog.btn_open_release.clicked.connect(self.open_release_url)
        dialog._update_buttons_wired = True

    # ---- 入口 ----

    def check_updates_manually(self):
        if not ENABLE_UPDATE_CHECK:
            return
        self.wire_about_dialog()
        dialog = getattr(self.view, "about_dialog", None)
        started = self._start_update_check(silent=False)
        if dialog and started:
            dialog.set_update_checking()
        elif dialog:
            dialog.set_update_result("已有更新检查正在进行，请稍后再试。")

    def check_updates_on_startup(self):
        if not ENABLE_UPDATE_CHECK:
            return
        settings = getattr(self.view, "app_settings", None)
        if not settings:
            return

        today = date.today().isoformat()
        if settings.value("Update/LastSilentCheckDate") == today:
            return

        if self._start_update_check(silent=True):
            settings.setValue("Update/LastSilentCheckDate", today)

    # ---- worker 生命周期 ----

    def _start_update_check(self, silent=False):
        if not ENABLE_UPDATE_CHECK:
            return False
        worker = self.update_worker
        if worker and worker.isRunning():
            return False

        worker = UpdateCheckWorker(silent=silent, parent=self)
        worker.finished_check.connect(self._handle_update_result)
        worker.finished.connect(worker.deleteLater)
        worker.finished.connect(lambda checked_worker=worker: self._clear_update_worker(checked_worker))
        self.update_worker = worker
        worker.start()
        return True

    def _clear_update_worker(self, worker):
        if self.update_worker is worker:
            self.update_worker = None

    def shutdown(self):
        worker = self.update_worker
        if worker and worker.isRunning():
            worker.wait(9000)

    # ---- 结果处理 ----

    def open_release_url(self, release_url=""):
        if not release_url:
            dialog = getattr(self.view, "about_dialog", None)
            release_url = getattr(dialog, "latest_release_url", "") if dialog else ""
        if not release_url:
            return

        try:
            opened = webbrowser.open(release_url)
        except Exception as exc:
            self.view.show_error_message("打开失败", f"无法打开发布页：{exc}")
            return

        if not opened:
            self.view.show_error_message("打开失败", "无法打开发布页：浏览器拒绝打开链接")

    def _handle_silent_update_result(self, result):
        if not result.ok or not result.has_update or not result.is_major or not result.latest_release:
            return

        settings = getattr(self.view, "app_settings", None)
        if not settings:
            return

        release = result.latest_release
        if settings.value("Update/IgnoredVersion", "") == release.version_text:
            return

        today = date.today().isoformat()
        prompt_marker = f"{today}:{release.version_text}"
        if settings.value("Update/LastPromptedVersion", "") == prompt_marker:
            return

        settings.setValue("Update/LastPromptedVersion", prompt_marker)
        action = self.view.show_major_update_prompt(result.current_version, release)

        if action == "open":
            self.open_release_url(release.html_url)
        elif action == "ignore":
            settings.setValue("Update/IgnoredVersion", release.version_text)

    def _handle_update_result(self, result, silent):
        if silent:
            self._handle_silent_update_result(result)
            return

        dialog = getattr(self.view, "about_dialog", None)
        if not dialog:
            return

        if not result.ok:
            dialog.set_update_result(f"检查更新失败：{result.error}")
            return

        if not result.has_update or not result.latest_release:
            dialog.set_update_result(f"当前已是最新版本：{result.current_version}")
            return

        release = result.latest_release
        published_at = release.published_at or "未知"
        message = (
            f"发现新版本：{release.version_text}\n"
            f"当前版本：{result.current_version}\n"
            f"发布标题：{release.title}\n"
            f"发布时间：{published_at}"
        )
        dialog.set_update_result(message, release.html_url)
