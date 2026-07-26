
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QHBoxLayout, QLabel, QPushButton,
)

from ratools_pdf.ui.log_view_model import (
    build_log_summary_items,
    filter_log_summary_items,
)
from ratools_pdf.ui.dialogs.base import FramelessDraggableDialog


class ManualFontEmbeddingDialog(FramelessDraggableDialog):
    def __init__(self, pdf_paths, open_callback, parent=None):
        super().__init__("🛠 手动嵌入缺失字体", parent)
        self.resize(520, 360)
        self.pdf_paths = list(pdf_paths or [])
        self.open_callback = open_callback

        intro = QLabel(
            "请在 Acrobat 中手动执行印前检查来嵌入缺失字体。\n"
            "点击下方按钮后会打开选中的 PDF，对话框会保持打开，便于你对照步骤操作。"
        )
        intro.setWordWrap(True)
        intro.setObjectName("dialogBody")
        self.content_layout.addWidget(intro)

        steps = QLabel(
            "操作路径：\n"
            "1. 所有工具 > 印刷制作 > 印前检查\n"
            "2. 选择“嵌入缺失的字体”\n"
            "3. 点击修复并保存\n"
            "4. 回到 RATools 重新预检确认字体风险是否消失"
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

        self.btn_close = QPushButton("关闭")
        self.btn_close.setObjectName("dialogSecondaryBtn")
        self.btn_close.setFixedHeight(32)
        self.btn_close.clicked.connect(self.accept)

        btn_layout.addWidget(self.btn_open_acrobat)
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
        self.status_label.setText(message)
        self.status_label.setObjectName("dialogStatus")
        self.status_label.setProperty("state", "success" if ok else "error")
        self.status_label.style().unpolish(self.status_label)
        self.status_label.style().polish(self.status_label)
        self.btn_open_acrobat.setEnabled(True)
