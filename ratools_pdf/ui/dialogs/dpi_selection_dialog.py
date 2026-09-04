from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QButtonGroup,
    QDialog,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QRadioButton,
    QSpinBox,
    QVBoxLayout,
)


class DpiSelectionDialog(QDialog):
    """图像压缩 DPI 选择对话框"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("图像压缩 DPI 设置")
        self.setModal(True)
        self.setFixedWidth(480)

        layout = QVBoxLayout(self)
        layout.setSpacing(16)
        layout.setContentsMargins(24, 24, 24, 24)

        # 标题说明
        title = QLabel("选择目标 DPI（越低文件越小）")
        title.setObjectName("dialogTitle")
        layout.addWidget(title)

        desc = QLabel(
            "降低 DPI 会减小文件体积，但可能影响图像清晰度。\n"
            "建议根据文档类型选择合适的 DPI 值。"
        )
        desc.setWordWrap(True)
        desc.setObjectName("dialogDesc")
        layout.addWidget(desc)

        # 单选按钮组
        self.btn_group = QButtonGroup(self)

        self.rb_150 = QRadioButton("150 DPI - 适用于纯文字/表格文档")
        self.rb_300 = QRadioButton("300 DPI - 适用于包含图表的文档（推荐）")
        self.rb_custom = QRadioButton("自定义 DPI")

        self.rb_300.setChecked(True)  # 默认选中 300 DPI

        self.btn_group.addButton(self.rb_150, 150)
        self.btn_group.addButton(self.rb_300, 300)
        self.btn_group.addButton(self.rb_custom, -1)

        layout.addWidget(self.rb_150)
        layout.addWidget(self.rb_300)

        # 自定义 DPI 输入
        custom_layout = QHBoxLayout()
        custom_layout.addWidget(self.rb_custom)
        self.spin_custom = QSpinBox()
        self.spin_custom.setRange(72, 600)
        self.spin_custom.setValue(150)
        self.spin_custom.setSuffix(" DPI")
        self.spin_custom.setEnabled(False)
        self.spin_custom.setMinimumWidth(120)
        custom_layout.addWidget(self.spin_custom)
        custom_layout.addStretch()
        layout.addLayout(custom_layout)

        # 启用/禁用自定义输入
        self.rb_custom.toggled.connect(self.spin_custom.setEnabled)

        # 警告提示
        warning = QLabel("⚠️ 注意：降低 DPI 会永久影响图像质量，请谨慎选择。")
        warning.setWordWrap(True)
        warning.setObjectName("dialogWarning")
        layout.addWidget(warning)

        # 按钮
        button_layout = QHBoxLayout()
        button_layout.addStretch()

        btn_cancel = QPushButton("取消")
        btn_cancel.setObjectName("secondaryBtn")
        btn_cancel.clicked.connect(self.reject)

        btn_ok = QPushButton("确定")
        btn_ok.setObjectName("primaryBtn")
        btn_ok.clicked.connect(self.accept)
        btn_ok.setDefault(True)

        button_layout.addWidget(btn_cancel)
        button_layout.addWidget(btn_ok)
        layout.addLayout(button_layout)

    def get_selected_dpi(self):
        """返回用户选择的 DPI 值"""
        if self.rb_150.isChecked():
            return 150
        elif self.rb_300.isChecked():
            return 300
        elif self.rb_custom.isChecked():
            return self.spin_custom.value()
        return 300  # 默认
