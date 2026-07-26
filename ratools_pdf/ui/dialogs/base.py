
from PySide6.QtCore import Qt
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QDialog, QFrame, QGraphicsDropShadowEffect, QHBoxLayout, QLabel, QPushButton, QVBoxLayout, QWidget,
)

from ratools_pdf.ui.log_view_model import (
    build_log_summary_items,
    filter_log_summary_items,
)
from ratools_pdf.ui import win32
from ratools_pdf.ui.platform import is_win11, should_use_manual_dialog_shadow
from ratools_pdf.ui.theme import active_palette


class FramelessDraggableDialog(QDialog):
    def __init__(self, title_text, parent=None):
        super().__init__(parent)
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.Dialog | Qt.NoDropShadowWindowHint)
        self.setAttribute(Qt.WA_TranslucentBackground)  # 支持圆角透明背景

        #  如果是 Win11，就调用去边框方法
        if is_win11():
            self._remove_win11_transparent_border()

        # 视觉样式全部来自应用级中央 QSS (theme.py)，此处不再本地硬编码颜色。

        self.main_layout = QVBoxLayout(self)

        if is_win11():
            # Win11 会自动贴着窗口边缘画阴影。必须把边距设为 0，否则阴影会画在 18px 的透明空气外围
            self.main_layout.setContentsMargins(0, 0, 0, 0)
        else:
            # Win10 需要手动渲染阴影，保留 18px 留给内部画阴影用
            self.main_layout.setContentsMargins(18, 18, 18, 18)

        self.main_layout.setSpacing(0)

        # 整体圆角和边框容器
        self.bg_frame = QFrame()
        self.bg_frame.setObjectName("dialogBg")

        if should_use_manual_dialog_shadow():
            shadow = QGraphicsDropShadowEffect(self)
            shadow.setBlurRadius(36)
            shadow.setOffset(0, 8)
            shadow.setColor(QColor(*active_palette().shadow_rgba))
            self.bg_frame.setGraphicsEffect(shadow)

        bg_layout = QVBoxLayout(self.bg_frame)
        bg_layout.setContentsMargins(0, 0, 0, 0)
        bg_layout.setSpacing(0)

        # 顶部自定义标题栏
        self.title_bar = QFrame()
        self.title_bar.setObjectName("dialogTitleBar")
        self.title_bar.setFixedHeight(40)
        tb_layout = QHBoxLayout(self.title_bar)
        tb_layout.setContentsMargins(16, 0, 8, 0)

        self.title_lbl = QLabel(title_text)
        self.title_lbl.setObjectName("dialogTitle")
        tb_layout.addWidget(self.title_lbl)
        tb_layout.addStretch()

        self.btn_close = QPushButton("✕")
        self.btn_close.setObjectName("dialogCloseBtn")
        self.btn_close.setFixedSize(30, 30)
        self.btn_close.clicked.connect(self.reject)
        tb_layout.addWidget(self.btn_close)

        bg_layout.addWidget(self.title_bar)

        # 内部内容区
        self.content_widget = QWidget()
        self.content_widget.setObjectName("dialogContent")
        self.content_layout = QVBoxLayout(self.content_widget)
        self.content_layout.setContentsMargins(24, 24, 24, 24)
        bg_layout.addWidget(self.content_widget)

        self.main_layout.addWidget(self.bg_frame)

    def mousePressEvent(self, event):
        """接管鼠标按下事件：若点击在实际标题栏区域内，则记录起始坐标"""
        title_top = self.bg_frame.y() + self.title_bar.y()
        title_bottom = title_top + self.title_bar.height()
        mouse_y = event.position().y()

        if event.button() == Qt.LeftButton and title_top <= mouse_y <= title_bottom:
            self.drag_pos = event.globalPosition().toPoint() - self.frameGeometry().topLeft()
            event.accept()

    def mouseMoveEvent(self, event):
        """接管鼠标移动事件：应用拖拽偏移"""
        if event.buttons() == Qt.LeftButton and hasattr(self, 'drag_pos'):
            self.move(event.globalPosition().toPoint() - self.drag_pos)
            event.accept()

    def mouseReleaseEvent(self, event):
        """接管释放事件：清除记录"""
        if hasattr(self, 'drag_pos'):
            del self.drag_pos
            event.accept()

    def _remove_win11_transparent_border(self):
        """专门处理 Win11 下强制附加的透明边框、圆角和阴影"""
        try:
            win32.remove_win11_window_decorations(int(self.winId()))
        except Exception:
            pass


# ================== 具体的业务对话框 ==================
