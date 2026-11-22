from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QButtonGroup,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QRadioButton,
    QSpinBox,
)

from ratools_pdf.config.compression import MAX_IMAGE_DPI, MIN_IMAGE_DPI
from ratools_pdf.ui.dialogs.base import FramelessDraggableDialog


class DpiSelectionDialog(FramelessDraggableDialog):
    """图像压缩 DPI 选择对话框"""

    def __init__(self, parent=None):
        super().__init__("图像压缩 DPI 设置", parent)
        self.setModal(True)
        self.setFixedWidth(480)
        self.content_layout.setSpacing(14)

        title = QLabel("选择目标 DPI（越低文件越小）")
        title.setObjectName("dialogSubTitle")
        self.content_layout.addWidget(title)

        desc = QLabel(
            "降低 DPI 会减小文件体积，但可能影响图像清晰度。\n"
            "建议根据文档类型选择合适的 DPI 值。"
        )
        desc.setWordWrap(True)
        desc.setObjectName("dialogMuted")
        self.content_layout.addWidget(desc)

        self.btn_group = QButtonGroup(self)

        self.rb_150 = QRadioButton("150 DPI - 适用于纯文字/表格文档")
        self.rb_300 = QRadioButton("300 DPI - 适用于包含图表的文档（推荐）")
        self.rb_custom = QRadioButton("自定义 DPI")
        self.rb_300.setChecked(True)

        self.btn_group.addButton(self.rb_150, 150)
        self.btn_group.addButton(self.rb_300, 300)
        self.btn_group.addButton(self.rb_custom, -1)

        self.content_layout.addWidget(self.rb_150)
        self.content_layout.addWidget(self.rb_300)

        custom_layout = QHBoxLayout()
        custom_layout.addWidget(self.rb_custom)
        self.spin_custom = QSpinBox()
        self.spin_custom.setRange(MIN_IMAGE_DPI, MAX_IMAGE_DPI)
        self.spin_custom.setValue(150)
        self.spin_custom.setSuffix(" DPI")
        self.spin_custom.setEnabled(False)
        self.spin_custom.setMinimumWidth(120)
        custom_layout.addWidget(self.spin_custom)
        custom_layout.addStretch()
        self.content_layout.addLayout(custom_layout)

        self.rb_custom.toggled.connect(self.spin_custom.setEnabled)

        warning = QLabel("注意：降低 DPI 会永久影响图像质量，请谨慎选择。")
        warning.setWordWrap(True)
        warning.setObjectName("dangerHint")
        self.content_layout.addWidget(warning)

        button_layout = QHBoxLayout()
        button_layout.setSpacing(10)
        button_layout.addStretch()

        btn_cancel = QPushButton("取消")
        btn_cancel.setObjectName("dialogSecondaryBtn")
        btn_cancel.setFixedHeight(34)
        btn_cancel.setCursor(Qt.PointingHandCursor)
        btn_cancel.clicked.connect(self.reject)

        btn_ok = QPushButton("确定")
        btn_ok.setObjectName("dialogPrimaryBtn")
        btn_ok.setFixedHeight(34)
        btn_ok.setCursor(Qt.PointingHandCursor)
        btn_ok.clicked.connect(self.accept)
        btn_ok.setDefault(True)

        button_layout.addWidget(btn_cancel)
        button_layout.addWidget(btn_ok)
        self.content_layout.addLayout(button_layout)

    def get_selected_dpi(self):
        """返回用户选择的 DPI 值"""
        if self.rb_150.isChecked():
            return 150
        elif self.rb_300.isChecked():
            return 300
        elif self.rb_custom.isChecked():
            return self.spin_custom.value()
        return 300
