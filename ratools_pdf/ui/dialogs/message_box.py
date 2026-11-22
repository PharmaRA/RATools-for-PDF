from PySide6.QtCore import Qt
from PySide6.QtWidgets import QDialog, QHBoxLayout, QLabel, QPushButton, QSizePolicy

from ratools_pdf.ui import icons
from ratools_pdf.ui.dialogs.base import FramelessDraggableDialog
from ratools_pdf.ui.theme import active_palette


def _strip_leading_emoji(text: str) -> str:
    prefixes = ("ℹ️ ", "✅ ", "⚠️ ", "❌ ", "❓ ", "🔎 ", "📝 ", "🗑️ ")
    for prefix in prefixes:
        if text.startswith(prefix):
            return text[len(prefix):]
    return text


class CustomMessageBox(FramelessDraggableDialog):
    """用于完全替代原生 QMessageBox 的统一 Fluent 提示框。"""

    Yes = 0x00004000
    No = 0x00010000
    Ok = 0x00000400
    Cancel = 0x00400000

    def __init__(self, title_text, message_text, msg_type="info", show_cancel=False, parent=None):
        super().__init__(_strip_leading_emoji(str(title_text or "")), parent)
        self.setMinimumWidth(420)
        self.setMaximumWidth(560)

        pal = active_palette()
        icon_specs = {
            "info": ("info_circle", pal.info_text),
            "success": ("check_circle", pal.success_text),
            "warning": ("alert_triangle", pal.warning_text),
            "error": ("x_circle", pal.danger_text),
            "question": ("help_circle", pal.primary_text),
        }
        variant = msg_type if msg_type in icon_specs else "info"
        icon_name, icon_color = icon_specs[variant]

        content_h_layout = QHBoxLayout()
        content_h_layout.setSpacing(14)

        icon_lbl = QLabel()
        icon_lbl.setObjectName("msgIconBadge")
        icon_lbl.setProperty("variant", variant)
        icon_lbl.setFixedSize(36, 36)
        icon_lbl.setAlignment(Qt.AlignCenter)
        icon_lbl.setPixmap(icons.render_svg_pixmap(icon_name, icon_color, size=20, stroke_width=1.85))

        msg_lbl = QLabel(message_text)
        is_multiline_block = ("\n\n" in message_text) or (" -> " in message_text)
        msg_lbl.setWordWrap(True)
        msg_lbl.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Minimum)
        if is_multiline_block:
            msg_lbl.setTextInteractionFlags(Qt.TextSelectableByMouse)
            msg_lbl.setObjectName("msgTextBlock")
        else:
            msg_lbl.setObjectName("msgText")
        msg_lbl.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)

        content_h_layout.addWidget(icon_lbl, 0, Qt.AlignTop)
        content_h_layout.addWidget(msg_lbl, 1)

        self.content_layout.addLayout(content_h_layout)
        self.content_layout.addStretch()

        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(10)
        btn_layout.addStretch()

        if show_cancel:
            self.btn_cancel = QPushButton("取 消")
            self.btn_cancel.setObjectName("dialogSecondaryBtn")
            self.btn_cancel.setFixedSize(84, 34)
            self.btn_cancel.setCursor(Qt.PointingHandCursor)
            self.btn_cancel.clicked.connect(self.reject)
            btn_layout.addWidget(self.btn_cancel)

        self.btn_ok = QPushButton("确 定")
        self.btn_ok.setFixedSize(84, 34)
        self.btn_ok.setCursor(Qt.PointingHandCursor)
        if msg_type in ["error", "warning"]:
            self.btn_ok.setObjectName("dialogDangerBtn")
        else:
            self.btn_ok.setObjectName("dialogPrimaryBtn")
        self.btn_ok.clicked.connect(self.accept)
        btn_layout.addWidget(self.btn_ok)

        self.content_layout.addLayout(btn_layout)
        self.adjustSize()

    @classmethod
    def information(cls, parent, title, text, *_args):
        cls(title, text, msg_type="info", parent=parent).exec()
        return cls.Ok

    @classmethod
    def warning(cls, parent, title, text, *_args):
        cls(title, text, msg_type="warning", parent=parent).exec()
        return cls.Ok

    @classmethod
    def critical(cls, parent, title, text, *_args):
        cls(title, text, msg_type="error", parent=parent).exec()
        return cls.Ok

    @classmethod
    def question(cls, parent, title, text, *_args):
        dlg = cls(title, text, msg_type="question", show_cancel=True, parent=parent)
        return cls.Yes if dlg.exec() == QDialog.Accepted else cls.No
