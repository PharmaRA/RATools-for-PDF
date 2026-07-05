from PySide6.QtCore import Signal
from PySide6.QtWidgets import QLabel


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
