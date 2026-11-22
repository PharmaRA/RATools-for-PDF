from __future__ import annotations

from PySide6.QtCore import QEvent, QRectF, QSize, Qt, Signal
from PySide6.QtGui import QColor, QFont, QIcon, QPainter, QPen
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QFrame,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QSizePolicy,
    QStyle,
    QStyledItemDelegate,
    QStyleOptionViewItem,
    QTreeWidget,
    QVBoxLayout,
)

from ratools_pdf.common.status import status_semantic
from ratools_pdf.ui import icons
from ratools_pdf.ui.theme import active_palette


def _extract_local_paths(mime_data) -> list[str]:
    if not mime_data or not mime_data.hasUrls():
        return []
    return [url.toLocalFile() for url in mime_data.urls() if url.isLocalFile()]


def _update_drag_active_property(widget, active: bool):
    widget.setProperty("dragActive", "true" if active else "false")
    widget.style().unpolish(widget)
    widget.style().polish(widget)


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
    """可接受文件/文件夹拖放的标签组件。"""

    files_dropped = Signal(list)

    def __init__(self, text="", parent=None):
        super().__init__(text, parent)
        self.setAcceptDrops(True)

    def _set_drag_active(self, active):
        _update_drag_active_property(self, active)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            self._set_drag_active(True)
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragLeaveEvent(self, event):
        self._set_drag_active(False)

    def dropEvent(self, event):
        self._set_drag_active(False)
        paths = _extract_local_paths(event.mimeData())
        if paths:
            self.files_dropped.emit(paths)


class DropZoneFrame(QFrame):
    """支持整片区域拖放的 Fluent 空队列引导容器。"""

    files_dropped = Signal(list)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("dropZone")
        self.setAcceptDrops(True)

    def _set_drag_active(self, active: bool):
        _update_drag_active_property(self, active)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            self._set_drag_active(True)
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragLeaveEvent(self, event):
        self._set_drag_active(False)

    def dropEvent(self, event):
        self._set_drag_active(False)
        paths = _extract_local_paths(event.mimeData())
        if paths:
            self.files_dropped.emit(paths)


class NoFocusItemDelegate(QStyledItemDelegate):
    """屏蔽树形列表获得焦点时在整行周围绘制的原生活动虚线框。"""

    def paint(self, painter, option, index):
        opt = QStyleOptionViewItem(option)
        opt.state &= ~QStyle.StateFlag.State_HasFocus
        super().paint(painter, opt, index)


class QueueTreeWidget(QTreeWidget):
    """支持直接向已填充的文件树拖放追加 PDF 或文件夹的树组件。"""

    files_dropped = Signal(list)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("queueTree")
        self.setAcceptDrops(True)
        self.setItemDelegate(NoFocusItemDelegate(self))

    def _set_drag_active(self, active: bool):
        _update_drag_active_property(self, active)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            self._set_drag_active(True)
        else:
            super().dragEnterEvent(event)

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            super().dragMoveEvent(event)

    def dragLeaveEvent(self, event):
        self._set_drag_active(False)
        super().dragLeaveEvent(event)

    def dropEvent(self, event):
        self._set_drag_active(False)
        paths = _extract_local_paths(event.mimeData())
        if paths:
            event.acceptProposedAction()
            self.files_dropped.emit(paths)
        else:
            super().dropEvent(event)


class NavModuleButton(QPushButton):
    """左侧模块导航按钮：支持 Fluent 选中左竖条与右侧已选规则数量角标。"""

    def __init__(self, title: str, icon_name: str = "", parent=None):
        super().__init__(title, parent)
        self._icon_name = icon_name
        self._badge_count = 0
        self.setObjectName("navBtn")
        self.setCheckable(True)
        self.setCursor(Qt.PointingHandCursor)
        self.setIconSize(QSize(16, 16))

    def set_badge_count(self, count: int):
        normalized = max(0, int(count))
        if normalized != self._badge_count:
            self._badge_count = normalized
            self.update()

    def badge_count(self) -> int:
        return self._badge_count

    def set_theme_icon(self, palette):
        if not self._icon_name:
            return
        self.setIcon(
            icons.svg_icon(
                self._icon_name,
                palette.text_muted,
                size=16,
                active_color=palette.primary_text,
                disabled_color=palette.text_faint,
            )
        )

    def paintEvent(self, event):
        super().paintEvent(event)
        pal = active_palette()
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)

        if self.isChecked():
            bar_h = max(14, self.height() - 16)
            bar_rect = QRectF(1.0, (self.height() - bar_h) / 2.0, 3.0, bar_h)
            painter.setPen(Qt.PenStyle.NoPen)
            painter.setBrush(QColor(pal.accent_bar))
            painter.drawRoundedRect(bar_rect, 1.5, 1.5)

        if self._badge_count > 0:
            text = str(self._badge_count)
            badge_font = QFont(self.font())
            badge_font.setPixelSize(11)
            badge_font.setWeight(QFont.Weight.Bold)
            painter.setFont(badge_font)
            fm = painter.fontMetrics()
            pill_w = max(22, fm.horizontalAdvance(text) + 12)
            pill_h = 20
            pill_x = self.width() - pill_w - 10
            pill_y = (self.height() - pill_h) / 2.0
            pill_rect = QRectF(pill_x, pill_y, pill_w, pill_h)

            if self.isChecked():
                painter.setPen(Qt.PenStyle.NoPen)
                painter.setBrush(QColor(pal.primary))
                text_color = QColor(pal.text_on_primary)
            else:
                painter.setPen(QPen(QColor(pal.border), 1.0))
                painter.setBrush(QColor(pal.surface))
                text_color = QColor(pal.text_muted)

            painter.drawRoundedRect(pill_rect, 4.0, 4.0)
            painter.setPen(text_color)
            painter.drawText(pill_rect, Qt.AlignmentFlag.AlignCenter, text)

        painter.end()


class RuleOptionCard(QFrame):
    """右侧整行可点击的 Fluent 规则选项卡片。"""

    def __init__(
        self,
        title: str,
        desc: str = "",
        checked: bool = False,
        tag: tuple[str, str] | None = None,
        parent=None,
    ):
        super().__init__(parent)
        self.setObjectName("ruleCard")
        self.setCursor(Qt.PointingHandCursor)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 10, 12, 10)
        layout.setSpacing(4)

        top_row = QHBoxLayout()
        top_row.setContentsMargins(0, 0, 0, 0)
        top_row.setSpacing(8)

        self.checkbox = QCheckBox()
        self.checkbox.setChecked(checked)
        self.checkbox.setFocusPolicy(Qt.NoFocus)

        self.title_label = QLabel(title)
        self.title_label.setWordWrap(True)
        self.title_label.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Minimum)
        self.title_label.setObjectName("optionTitle")

        top_row.addWidget(self.checkbox, 0, Qt.AlignmentFlag.AlignTop)
        top_row.addWidget(self.title_label, 1)

        if tag:
            tag_text, tag_tone = tag
            self.tag_label = QLabel(tag_text)
            self.tag_label.setObjectName("ruleTag")
            self.tag_label.setProperty("tone", tag_tone)
            top_row.addWidget(self.tag_label, 0, Qt.AlignmentFlag.AlignTop)
        else:
            self.tag_label = None

        layout.addLayout(top_row)

        if desc:
            self.desc_label = QLabel(desc)
            self.desc_label.setWordWrap(True)
            self.desc_label.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Minimum)
            self.desc_label.setObjectName("optionDesc")
            desc_row = QHBoxLayout()
            desc_row.setContentsMargins(24, 0, 0, 0)
            desc_row.setSpacing(0)
            desc_row.addWidget(self.desc_label, 1)
            layout.addLayout(desc_row)
        else:
            self.desc_label = None

        self.checkbox.toggled.connect(self._sync_checked_state)
        self._sync_checked_state(checked)

    def _sync_checked_state(self, checked: bool):
        self.setProperty("checked", "true" if checked else "false")
        self.style().unpolish(self)
        self.style().polish(self)

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton and self.checkbox.isEnabled():
            self.checkbox.toggle()
            event.accept()
            return
        super().mousePressEvent(event)


class StatusBadgeDelegate(QStyledItemDelegate):
    """文件树状态列胶囊绘制器：在当前状态列绘制圆角彩色状态胶囊。"""

    def sizeHint(self, option, index):
        base = super().sizeHint(option, index)
        return QSize(base.width(), max(base.height(), 34))

    def paint(self, painter, option, index):
        opt = QStyleOptionViewItem(option)
        self.initStyleOption(opt, index)
        opt.state &= ~QStyle.StateFlag.State_HasFocus
        text = (opt.text or "").strip()

        # 先由样式系统绘制行底色/悬停态/选中态，清除原生文字与图标避免重影
        opt.text = ""
        opt.icon = QIcon()
        style = opt.widget.style() if opt.widget else QApplication.style()
        style.drawControl(QStyle.ControlElement.CE_ItemViewItem, opt, painter, opt.widget)

        if not text:
            return

        pal = active_palette()
        semantic = status_semantic(text)
        if text == "等待处理":
            dot_hex = pal.text_faint
            bg_hex = pal.surface_alt
            fg_hex = pal.text_muted
            border_hex = pal.border
        elif semantic == "positive":
            dot_hex = pal.success
            bg_hex = pal.success_soft
            fg_hex = pal.success_text
            border_hex = ""
        elif semantic == "negative":
            dot_hex = pal.danger
            bg_hex = pal.danger_soft
            fg_hex = pal.danger_text
            border_hex = ""
        elif semantic == "warning":
            dot_hex = pal.warning
            bg_hex = pal.warning_soft
            fg_hex = pal.warning_text
            border_hex = ""
        else:
            dot_hex = pal.info
            bg_hex = pal.info_soft
            fg_hex = pal.info_text
            border_hex = ""

        painter.save()
        painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)

        pill_font = QFont(opt.font)
        pill_font.setPixelSize(12)
        pill_font.setWeight(QFont.Weight.DemiBold)
        painter.setFont(pill_font)
        fm = painter.fontMetrics()

        dot_size = 6
        gap = 6
        h_pad = 9
        pill_h = 22
        max_pill_w = max(40, opt.rect.width() - 12)
        natural_w = h_pad * 2 + dot_size + gap + fm.horizontalAdvance(text)
        pill_w = min(max_pill_w, natural_w)

        pill_x = opt.rect.x() + 6
        pill_y = opt.rect.y() + (opt.rect.height() - pill_h) / 2.0
        pill_rect = QRectF(pill_x, pill_y, pill_w, pill_h)

        painter.setBrush(QColor(bg_hex))
        if border_hex:
            painter.setPen(QPen(QColor(border_hex), 1.0))
        else:
            painter.setPen(Qt.PenStyle.NoPen)
        painter.drawRoundedRect(pill_rect, 4.0, 4.0)

        dot_x = pill_x + h_pad
        dot_y = pill_y + (pill_h - dot_size) / 2.0
        painter.setPen(Qt.PenStyle.NoPen)
        painter.setBrush(QColor(dot_hex))
        painter.drawEllipse(QRectF(dot_x, dot_y, dot_size, dot_size))

        text_x = dot_x + dot_size + gap
        text_w = max(0.0, pill_x + pill_w - h_pad - text_x)
        elided = fm.elidedText(text, Qt.TextElideMode.ElideRight, int(text_w))
        painter.setPen(QColor(fg_hex))
        painter.drawText(
            QRectF(text_x, pill_y, text_w, pill_h),
            Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft,
            elided,
        )
        painter.restore()
