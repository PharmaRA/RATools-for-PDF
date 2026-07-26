from PySide6.QtCore import QEvent, QSize, Qt, Signal
from PySide6.QtWidgets import QLabel


class ElidedLabel(QLabel):
    """单行标签：空间不足时按 ElideRight 截断，并仅在截断时挂完整文本 tooltip。

    sizeHint 始终按完整文本上报，因此窗口变宽时布局会把空间还给标签；
    minimumSizeHint 归零，挤压时标签先让位，避免把同行按钮顶出窗口。
    """

    def __init__(self, text="", parent=None):
        super().__init__(parent)
        self._full_text = ""
        self.setText(text)

    def setText(self, text):
        self._full_text = text or ""
        self._update_elided()

    def text(self):
        return self._full_text

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._update_elided()

    def changeEvent(self, event):
        super().changeEvent(event)
        if event.type() in (QEvent.FontChange, QEvent.StyleChange):
            self._update_elided()

    def _update_elided(self):
        metrics = self.fontMetrics()
        available = max(0, self.width() - 2)
        elided = metrics.elidedText(self._full_text, Qt.ElideRight, available)
        super().setText(elided)
        self.setToolTip(self._full_text if elided != self._full_text else "")

    def sizeHint(self):
        base = super().sizeHint()
        return QSize(self.fontMetrics().horizontalAdvance(self._full_text) + 4, base.height())

    def minimumSizeHint(self):
        return QSize(0, super().minimumSizeHint().height())


class DropZoneLabel(QLabel):
    files_dropped = Signal(list)

    def __init__(self, text):
        super().__init__(text)
        self.setAcceptDrops(True)

    def _set_drag_active(self, active):
        # 拖拽高亮通过动态属性驱动中央 QSS (#dropZone[dragActive="true"])，
        # 不再本地硬编码颜色。
        self.setProperty("dragActive", "true" if active else "false")
        self.style().unpolish(self)
        self.style().polish(self)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            self._set_drag_active(True)
        else:
            event.ignore()

    def dragLeaveEvent(self, event):
        self._set_drag_active(False)

    def dropEvent(self, event):
        self._set_drag_active(False)
        urls = event.mimeData().urls()
        paths = [url.toLocalFile() for url in urls if url.isLocalFile()]
        if paths:
            self.files_dropped.emit(paths)
