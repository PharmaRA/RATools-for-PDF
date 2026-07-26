import os

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QAbstractSpinBox, QButtonGroup, QCheckBox, QFileDialog, QFrame, QHBoxLayout, QLabel, QLineEdit, QPushButton, QSpinBox, QToolButton, QVBoxLayout,
)

from ratools_pdf.ui.log_view_model import (
    build_log_summary_items,
    filter_log_summary_items,
)
from ratools_pdf.ui.dialogs.base import FramelessDraggableDialog


class SettingsDialog(FramelessDraggableDialog):
    def __init__(self, parent=None):
        super().__init__("⚙️ 全局设置", parent)
        self.resize(560, 420)

        self.content_layout.setSpacing(16)

        # 复选框视觉由中央 QSS 统一定义；覆盖原文件项用 dangerCheck 标红。
        self.cb_auto_open = QCheckBox("处理完成后自动打开输出文件夹")
        self.cb_auto_open.setChecked(True)

        self.cb_overwrite = QCheckBox("覆盖原始文件 (不推荐)")
        self.cb_overwrite.setChecked(False)
        self.cb_overwrite.setObjectName("dangerCheck")

        self.cb_parallel_processing = QCheckBox("启用并行处理")
        self.cb_parallel_processing.setChecked(False)

        self.parallel_max_workers = max(2, min(os.cpu_count() or 2, 4))
        self.spin_parallel_workers = QSpinBox()
        self.spin_parallel_workers.setRange(2, self.parallel_max_workers)
        self.spin_parallel_workers.setValue(2)
        self.spin_parallel_workers.setEnabled(False)
        self.spin_parallel_workers.setObjectName("settingsWorkerSpin")
        self.spin_parallel_workers.setSuffix(" 个任务")
        self.spin_parallel_workers.setButtonSymbols(QAbstractSpinBox.NoButtons)

        self.parallel_worker_stepper = QFrame()
        self.parallel_worker_stepper.setObjectName("settingsSpinStepper")
        stepper_layout = QVBoxLayout(self.parallel_worker_stepper)
        stepper_layout.setContentsMargins(0, 0, 0, 0)
        stepper_layout.setSpacing(0)

        self.btn_parallel_worker_up = QToolButton()
        self.btn_parallel_worker_down = QToolButton()
        for btn, direction, arrow_type in [
            (self.btn_parallel_worker_up, "up", Qt.UpArrow),
            (self.btn_parallel_worker_down, "down", Qt.DownArrow),
        ]:
            btn.setObjectName("settingsSpinStepBtn")
            btn.setProperty("direction", direction)
            btn.setArrowType(arrow_type)
            btn.setCursor(Qt.PointingHandCursor)
            btn.setFocusPolicy(Qt.NoFocus)
            btn.setEnabled(False)
            stepper_layout.addWidget(btn)

        self.btn_parallel_worker_up.clicked.connect(self.spin_parallel_workers.stepUp)
        self.btn_parallel_worker_down.clicked.connect(self.spin_parallel_workers.stepDown)
        self.cb_parallel_processing.toggled.connect(self.set_parallel_worker_controls_enabled)

        title = QLabel("外观主题")
        title.setObjectName("dialogSectionTitle")
        self.content_layout.addWidget(title)

        theme_card = QFrame()
        theme_card.setObjectName("settingsPathCard")
        theme_layout = QVBoxLayout(theme_card)
        theme_layout.setContentsMargins(12, 12, 12, 12)
        theme_layout.setSpacing(10)

        theme_hint = QLabel("默认跟随系统深浅色，也可手动锁定为亮色或暗色。")
        theme_hint.setWordWrap(True)
        theme_hint.setObjectName("settingsPathHint")
        theme_layout.addWidget(theme_hint)

        theme_row = QHBoxLayout()
        theme_row.setContentsMargins(0, 0, 0, 0)
        theme_row.setSpacing(8)
        self.theme_mode_group = QButtonGroup(self)
        self.theme_mode_group.setExclusive(True)
        self.theme_buttons = {}
        for mode, label in [("system", "跟随系统"), ("light", "亮色"), ("dark", "暗色")]:
            btn = QPushButton(label)
            btn.setObjectName("themeSegBtn")
            btn.setCheckable(True)
            btn.setCursor(Qt.PointingHandCursor)
            btn.setFocusPolicy(Qt.NoFocus)
            self.theme_mode_group.addButton(btn)
            self.theme_buttons[mode] = btn
            theme_row.addWidget(btn)
        theme_row.addStretch()
        theme_layout.addLayout(theme_row)
        self.content_layout.addWidget(theme_card)

        title = QLabel("默认保存位置")
        title.setObjectName("dialogSectionTitle")
        self.content_layout.addWidget(title)

        path_card = QFrame()
        path_card.setObjectName("settingsPathCard")
        path_layout = QVBoxLayout(path_card)
        path_layout.setContentsMargins(12, 12, 12, 12)
        path_layout.setSpacing(10)

        path_hint = QLabel("用于处理输出目录选择和日志导出默认位置")
        path_hint.setObjectName("settingsPathHint")
        path_layout.addWidget(path_hint)

        path_row = QHBoxLayout()
        path_row.setContentsMargins(0, 0, 0, 0)
        path_row.setSpacing(8)
        self.default_output_edit = QLineEdit()
        self.default_output_edit.setPlaceholderText("未设置时，处理输出目录选择将使用系统默认位置")
        self.default_output_edit.setReadOnly(True)
        self.default_output_edit.setObjectName("settingsPathEdit")
        self.default_output_status = QLabel()
        self.default_output_status.setObjectName("settingsPathStatus")
        self.btn_browse_output = QPushButton("浏览...")
        self.btn_browse_output.setObjectName("dialogSecondaryBtn")
        self.btn_clear_output = QPushButton("清除")
        self.btn_clear_output.setObjectName("dialogSecondaryBtn")
        path_row.addWidget(self.default_output_edit, 1)
        path_row.addWidget(self.btn_browse_output)
        path_row.addWidget(self.btn_clear_output)
        path_layout.addLayout(path_row)
        path_layout.addWidget(self.default_output_status)

        self.content_layout.addWidget(path_card)

        title = QLabel("常规选项")
        title.setObjectName("dialogSectionTitle")
        self.content_layout.addWidget(title)
        self.content_layout.addWidget(self.cb_auto_open)
        self.content_layout.addWidget(self.cb_overwrite)

        danger_hint = QLabel("危险操作：直接覆盖源PDF。建议仅在已有备份且确认规则无误后使用。")
        danger_hint.setWordWrap(True)
        danger_hint.setObjectName("dangerHint")
        self.content_layout.addWidget(danger_hint)

        title = QLabel("处理性能")
        title.setObjectName("dialogSectionTitle")
        self.content_layout.addWidget(title)

        parallel_card = QFrame()
        parallel_card.setObjectName("settingsPathCard")
        parallel_layout = QVBoxLayout(parallel_card)
        parallel_layout.setContentsMargins(12, 12, 12, 12)
        parallel_layout.setSpacing(10)

        parallel_hint = QLabel("并行处理会同时处理多个 PDF。建议从 2 个任务开始，异常文件可在队列中选中后单独终止。")
        parallel_hint.setWordWrap(True)
        parallel_hint.setObjectName("settingsPathHint")
        parallel_layout.addWidget(parallel_hint)

        parallel_layout.addWidget(self.cb_parallel_processing)

        worker_row = QHBoxLayout()
        worker_row.setContentsMargins(24, 0, 0, 0)
        worker_row.setSpacing(8)
        worker_label = QLabel("并行数量")
        worker_label.setObjectName("settingsPathHint")
        worker_row.addWidget(worker_label)
        worker_spin_layout = QHBoxLayout()
        worker_spin_layout.setContentsMargins(0, 0, 0, 0)
        worker_spin_layout.setSpacing(0)
        worker_spin_layout.addWidget(self.spin_parallel_workers)
        worker_spin_layout.addWidget(self.parallel_worker_stepper)
        worker_row.addLayout(worker_spin_layout)
        worker_row.addStretch()
        parallel_layout.addLayout(worker_row)

        self.content_layout.addWidget(parallel_card)

        self.content_layout.addStretch()

        btn_close = QPushButton("确 定")
        btn_close.setObjectName("dialogPrimaryBtn")
        btn_close.setFixedHeight(36)
        btn_close.clicked.connect(self.accept)
        self.content_layout.addWidget(btn_close)

        self.btn_browse_output.clicked.connect(self.choose_default_output_dir)
        self.btn_clear_output.clicked.connect(lambda: self.default_output_edit.setText(""))
        self.default_output_edit.textChanged.connect(self.update_default_output_status)
        self.update_default_output_status()

    def choose_default_output_dir(self):
        start_dir = self.default_output_edit.text().strip() or os.path.expanduser("~")
        selected_dir = QFileDialog.getExistingDirectory(self, "选择默认保存位置", start_dir)
        if selected_dir:
            self.default_output_edit.setText(selected_dir)

    def update_default_output_status(self):
        path = self.default_output_edit.text().strip()
        if not path:
            self.default_output_status.setText("未设置，处理输出目录选择将使用系统默认位置。")
            self.default_output_status.setProperty("state", "empty")
        elif os.path.isdir(path):
            self.default_output_status.setText("路径有效，将优先作为处理输出和日志导出的默认位置。")
            self.default_output_status.setProperty("state", "valid")
        else:
            self.default_output_status.setText("路径不存在，请重新选择有效目录。")
            self.default_output_status.setProperty("state", "invalid")

        self.style().unpolish(self.default_output_status)
        self.style().polish(self.default_output_status)

    def get_theme_mode(self):
        for mode, btn in self.theme_buttons.items():
            if btn.isChecked():
                return mode
        return "system"

    def set_theme_mode(self, mode):
        if mode not in self.theme_buttons:
            mode = "system"
        for key, btn in self.theme_buttons.items():
            btn.setChecked(key == mode)

    def set_parallel_worker_controls_enabled(self, enabled):
        self.spin_parallel_workers.setEnabled(enabled)
        self.btn_parallel_worker_up.setEnabled(enabled)
        self.btn_parallel_worker_down.setEnabled(enabled)
