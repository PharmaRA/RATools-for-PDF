
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QHBoxLayout, QLabel, QPushButton,
)

from ratools_pdf.ui.log_view_model import (
    build_log_summary_items,
    filter_log_summary_items,
)
from ratools_pdf.ui.dialogs.base import FramelessDraggableDialog


class CustomMessageBox(FramelessDraggableDialog):
    """用于完全替代原生 QMessageBox 的统一提示框"""

    def __init__(self, title_text, message_text, msg_type="info", show_cancel=False, parent=None):
        super().__init__(title_text, parent)
        self.resize(400, 200)

        icons = {
            "info": "ℹ️",
            "success": "✅",
            "warning": "⚠️",
            "error": "❌",
            "question": "❓"
        }
        icon_char = icons.get(msg_type, "ℹ️")

        content_h_layout = QHBoxLayout()
        content_h_layout.setSpacing(16)

        icon_lbl = QLabel(icon_char)
        icon_lbl.setObjectName("msgIcon")
        icon_lbl.setAlignment(Qt.AlignTop)

        msg_lbl = QLabel(message_text)
        is_multiline_block = ("\n\n" in message_text) or (" -> " in message_text)
        msg_lbl.setWordWrap(True)
        if is_multiline_block:
            msg_lbl.setTextInteractionFlags(Qt.TextSelectableByMouse)
            msg_lbl.setObjectName("msgTextBlock")
        else:
            msg_lbl.setObjectName("msgText")
        msg_lbl.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)

        content_h_layout.addWidget(icon_lbl)
        content_h_layout.addWidget(msg_lbl, 1)

        self.content_layout.addLayout(content_h_layout)
        self.content_layout.addStretch()

        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(12)
        btn_layout.addStretch()

        if show_cancel:
            self.btn_cancel = QPushButton("取 消")
            self.btn_cancel.setObjectName("dialogSecondaryBtn")
            self.btn_cancel.setFixedSize(80, 32)
            self.btn_cancel.clicked.connect(self.reject)
            btn_layout.addWidget(self.btn_cancel)

        self.btn_ok = QPushButton("确 定")
        self.btn_ok.setFixedSize(80, 32)
        if msg_type in ["error", "warning"]:
            self.btn_ok.setObjectName("dialogDangerBtn")
        else:
            self.btn_ok.setObjectName("dialogPrimaryBtn")
        self.btn_ok.clicked.connect(self.accept)
        btn_layout.addWidget(self.btn_ok)

        self.content_layout.addLayout(btn_layout)
