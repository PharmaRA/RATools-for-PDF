
from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtWidgets import (
    QHBoxLayout, QLabel, QPushButton,
)

from ratools_pdf.ui.log_view_model import (
    build_log_summary_items,
    filter_log_summary_items,
)
from ratools_pdf.ui.dialogs.base import FramelessDraggableDialog


class _FontRecheckWorker(QThread):
    result_ready = Signal(bool, str)

    def __init__(self, callback, parent):
        super().__init__(parent)
        self.callback = callback

    def run(self):
        try:
            ok, message = self.callback()
        except Exception as exc:
            ok, message = False, f"字体复检失败（{type(exc).__name__}），未确认风险已消除。"
        self.result_ready.emit(bool(ok), str(message))


class ManualFontEmbeddingDialog(FramelessDraggableDialog):
    def __init__(self, pdf_paths, open_callback, parent=None, *, capability_text="", recheck_callback=None):
        super().__init__("🛠 手动嵌入缺失字体", parent)
        self.resize(520, 360)
        self.pdf_paths = list(pdf_paths or [])
        self.open_callback = open_callback
        self.recheck_callback = recheck_callback
        self._recheck_worker = None

        intro = QLabel(
            "请在 Acrobat 中手动执行印前检查来嵌入缺失字体。\n"
            "点击下方按钮后会打开选中的 PDF，对话框会保持打开，便于你对照步骤操作。"
        )
        intro.setWordWrap(True)
        intro.setObjectName("dialogBody")
        self.content_layout.addWidget(intro)

        self.capability_label = QLabel(capability_text or "自动字体修复不可用；当前仅支持手动交接，保存后需复检。")
        self.capability_label.setWordWrap(True)
        self.capability_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        self.capability_label.setObjectName("dialogMuted")
        self.content_layout.addWidget(self.capability_label)

        steps = QLabel(
            "操作路径：\n"
            "1. 所有工具 > 印刷制作 > 印前检查\n"
            "2. 选择“嵌入缺失的字体”\n"
            "3. 点击修复并保存\n"
            "4. 点击“保存后复检字体”，确认未嵌入、替代风险、状态未知三类风险"
        )
        steps.setWordWrap(True)
        steps.setTextInteractionFlags(Qt.TextSelectableByMouse)
        steps.setObjectName("dialogCodeBlock")
        self.content_layout.addWidget(steps)

        file_count = len(self.pdf_paths)
        self.status_label = QLabel(f"已选中 {file_count} 个 PDF。")
        self.status_label.setWordWrap(True)
        self.status_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        self.status_label.setObjectName("dialogMuted")
        self.content_layout.addWidget(self.status_label)
        self.content_layout.addStretch()

        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(10)
        btn_layout.addStretch()

        self.btn_open_acrobat = QPushButton("打开 Acrobat")
        self.btn_open_acrobat.setObjectName("dialogPrimaryBtn")
        self.btn_open_acrobat.setFixedHeight(32)
        self.btn_open_acrobat.clicked.connect(self.open_acrobat)

        self.btn_recheck = QPushButton("保存后复检字体")
        self.btn_recheck.setObjectName("dialogSecondaryBtn")
        self.btn_recheck.setFixedHeight(32)
        self.btn_recheck.setEnabled(bool(recheck_callback))
        self.btn_recheck.clicked.connect(self.start_recheck)

        self.btn_close = QPushButton("关闭")
        self.btn_close.setObjectName("dialogSecondaryBtn")
        self.btn_close.setFixedHeight(32)
        self.btn_close.clicked.connect(self.accept)

        btn_layout.addWidget(self.btn_open_acrobat)
        btn_layout.addWidget(self.btn_recheck)
        btn_layout.addWidget(self.btn_close)
        self.content_layout.addLayout(btn_layout)

    def open_acrobat(self):
        if not self.open_callback:
            self.status_label.setText("无法打开 Acrobat：缺少打开回调。")
            return
        self.btn_open_acrobat.setEnabled(False)
        try:
            ok, message = self.open_callback()
        except Exception as exc:
            ok, message = False, f"无法打开 Acrobat：{exc}"
        self._set_status(message, "pending" if ok else "error")
        self.btn_open_acrobat.setEnabled(True)

    def _set_status(self, message, state):
        self.status_label.setText(message)
        self.status_label.setObjectName("dialogStatus")
        self.status_label.setProperty("state", state)
        self.status_label.style().unpolish(self.status_label)
        self.status_label.style().polish(self.status_label)

    def start_recheck(self):
        if self._recheck_worker is not None or not self.recheck_callback:
            return
        self.btn_open_acrobat.setEnabled(False)
        self.btn_recheck.setEnabled(False)
        self.btn_close.setEnabled(False)
        self._set_status("正在重新读取 PDF，检查三类字体风险……", "pending")
        worker = _FontRecheckWorker(self.recheck_callback, self)
        self._recheck_worker = worker
        worker.result_ready.connect(self._show_recheck_result)
        worker.finished.connect(self._finish_recheck)
        worker.start()

    def _show_recheck_result(self, ok, message):
        self._set_status(message, "success" if ok else "error")

    def _finish_recheck(self):
        worker = self._recheck_worker
        self._recheck_worker = None
        self.btn_open_acrobat.setEnabled(True)
        self.btn_recheck.setEnabled(bool(self.recheck_callback))
        self.btn_close.setEnabled(True)
        if worker is not None:
            worker.deleteLater()

    def closeEvent(self, event):
        if self._recheck_worker is not None:
            event.ignore()
            return
        super().closeEvent(event)

    def accept(self):
        if self._recheck_worker is None:
            super().accept()

    def reject(self):
        if self._recheck_worker is None:
            super().reject()
