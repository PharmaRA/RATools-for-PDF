import os

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QAbstractSpinBox, QButtonGroup, QCheckBox, QFileDialog, QFrame, QHBoxLayout, QLabel, QLineEdit, QPushButton, QScrollArea, QSizePolicy, QSpinBox, QToolButton, QVBoxLayout, QWidget,
)

from ratools_pdf.ui import theme
from ratools_pdf.ui.log_view_model import (
    build_log_summary_items,
    filter_log_summary_items,
)
from ratools_pdf.ui.dialogs.base import FramelessDraggableDialog


class SettingsDialog(FramelessDraggableDialog):
    def __init__(self, parent=None):
        super().__init__("全局设置", parent)
        self.resize(560, 560)
        self.setMinimumSize(520, 460)

        self.content_layout.setSpacing(14)

        scroll_area = QScrollArea()
        scroll_area.setObjectName("settingsScroll")
        scroll_area.setWidgetResizable(True)
        scroll_area.setFrameShape(QFrame.NoFrame)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll_area.viewport().setAutoFillBackground(True)
        scroll_area.verticalScrollBar().valueChanged.connect(scroll_area.viewport().update)

        body_widget = QWidget()
        body_layout = QVBoxLayout(body_widget)
        body_layout.setContentsMargins(2, 2, 6, 4)
        body_layout.setSpacing(14)

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
        self.spin_parallel_workers.setSuffix(" 个任务")
        (
            worker_spin_layout,
            self.parallel_worker_stepper,
            self.btn_parallel_worker_up,
            self.btn_parallel_worker_down,
        ) = self._build_stepper_spinbox(self.spin_parallel_workers, enabled=False)

        self.cb_parallel_processing.toggled.connect(self.set_parallel_worker_controls_enabled)

        title = QLabel("外观主题")
        title.setObjectName("dialogSectionTitle")
        body_layout.addWidget(title)

        theme_card = QFrame()
        theme_card.setObjectName("settingsPathCard")
        theme_layout = QVBoxLayout(theme_card)
        theme_layout.setContentsMargins(12, 12, 12, 12)
        theme_layout.setSpacing(10)

        theme_hint = QLabel("默认跟随系统深浅色，也可手动锁定为亮色或暗色。")
        theme_hint.setWordWrap(True)
        theme_hint.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Minimum)
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
        body_layout.addWidget(theme_card)

        title = QLabel("界面字号")
        title.setObjectName("dialogSectionTitle")
        body_layout.addWidget(title)

        font_card = QFrame()
        font_card.setObjectName("settingsPathCard")
        font_layout = QVBoxLayout(font_card)
        font_layout.setContentsMargins(12, 12, 12, 12)
        font_layout.setSpacing(10)

        font_hint = QLabel("默认使用微软雅黑 UI，可按屏幕分辨率切换界面字号大小。")
        font_hint.setWordWrap(True)
        font_hint.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Minimum)
        font_hint.setObjectName("settingsPathHint")
        font_layout.addWidget(font_hint)

        scale_row = QHBoxLayout()
        scale_row.setContentsMargins(0, 0, 0, 0)
        scale_row.setSpacing(8)
        self.font_scale_group = QButtonGroup(self)
        self.font_scale_group.setExclusive(True)
        self.font_scale_buttons = {}
        for scale_key in ("compact", "standard", "comfortable"):
            btn = QPushButton(theme.FONT_SCALE_LABELS[scale_key])
            btn.setObjectName("themeSegBtn")
            btn.setCheckable(True)
            btn.setCursor(Qt.PointingHandCursor)
            btn.setFocusPolicy(Qt.NoFocus)
            self.font_scale_group.addButton(btn)
            self.font_scale_buttons[scale_key] = btn
            scale_row.addWidget(btn)
        scale_row.addStretch()
        self.lbl_font_scale_status = QLabel("")
        self.lbl_font_scale_status.setObjectName("settingsPathHint")
        scale_row.addWidget(self.lbl_font_scale_status)
        font_layout.addLayout(scale_row)

        body_layout.addWidget(font_card)

        title = QLabel("默认保存位置")
        title.setObjectName("dialogSectionTitle")
        body_layout.addWidget(title)

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
        self.btn_browse_output.setCursor(Qt.PointingHandCursor)
        self.btn_clear_output = QPushButton("清除")
        self.btn_clear_output.setObjectName("dialogSecondaryBtn")
        self.btn_clear_output.setCursor(Qt.PointingHandCursor)
        path_row.addWidget(self.default_output_edit, 1)
        path_row.addWidget(self.btn_browse_output)
        path_row.addWidget(self.btn_clear_output)
        path_layout.addLayout(path_row)
        path_layout.addWidget(self.default_output_status)

        body_layout.addWidget(path_card)

        title = QLabel("常规选项")
        title.setObjectName("dialogSectionTitle")
        body_layout.addWidget(title)
        body_layout.addWidget(self.cb_auto_open)
        body_layout.addWidget(self.cb_overwrite)

        danger_hint = QLabel("危险操作：直接覆盖源PDF。建议仅在已有备份且确认规则无误后使用。")
        danger_hint.setWordWrap(True)
        danger_hint.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Minimum)
        danger_hint.setObjectName("dangerHint")
        body_layout.addWidget(danger_hint)

        title = QLabel("处理性能")
        title.setObjectName("dialogSectionTitle")
        body_layout.addWidget(title)

        parallel_card = QFrame()
        parallel_card.setObjectName("settingsPathCard")
        parallel_layout = QVBoxLayout(parallel_card)
        parallel_layout.setContentsMargins(12, 12, 12, 12)
        parallel_layout.setSpacing(10)

        parallel_hint = QLabel("并行数量为用户上限，将按可用内存和输入大小保守下调。资源信息不可用时逐个处理；异常文件可在队列中选中后单独终止。")
        parallel_hint.setWordWrap(True)
        parallel_hint.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Minimum)
        parallel_hint.setObjectName("settingsPathHint")
        parallel_layout.addWidget(parallel_hint)

        parallel_layout.addWidget(self.cb_parallel_processing)

        worker_row = QHBoxLayout()
        worker_row.setContentsMargins(24, 0, 0, 0)
        worker_row.setSpacing(8)
        worker_label = QLabel("并行数量")
        worker_label.setObjectName("settingsPathHint")
        worker_row.addWidget(worker_label)
        worker_row.addLayout(worker_spin_layout)
        worker_row.addStretch()
        parallel_layout.addLayout(worker_row)

        perf_sep = QFrame()
        perf_sep.setObjectName("compressionCardSep")
        perf_sep.setFixedHeight(1)
        parallel_layout.addWidget(perf_sep)

        timeout_row = QHBoxLayout()
        timeout_row.setContentsMargins(0, 0, 0, 0)
        timeout_row.setSpacing(8)
        timeout_label = QLabel("每文件超时")
        timeout_label.setObjectName("settingsPathHint")
        timeout_row.addWidget(timeout_label)
        self.spin_file_timeout = QSpinBox()
        self.spin_file_timeout.setRange(0, 86400)
        self.spin_file_timeout.setValue(0)
        self.spin_file_timeout.setSuffix(" 秒")
        self.spin_file_timeout.setSpecialValueText("关闭")
        self.spin_file_timeout.setMinimumWidth(96)
        self.spin_file_timeout.setToolTip("0 为关闭。超时文件保留既有输出，可调整设置后重试；计时包含处理和发布前验证。")
        (
            timeout_spin_layout,
            self.file_timeout_stepper,
            self.btn_file_timeout_up,
            self.btn_file_timeout_down,
        ) = self._build_stepper_spinbox(self.spin_file_timeout, enabled=True)
        timeout_row.addLayout(timeout_spin_layout)
        timeout_row.addStretch()
        parallel_layout.addLayout(timeout_row)

        body_layout.addWidget(parallel_card)
        body_layout.addStretch()

        scroll_area.setWidget(body_widget)
        self.content_layout.addWidget(scroll_area, 1)

        btn_close = QPushButton("确 定")
        btn_close.setObjectName("dialogPrimaryBtn")
        btn_close.setFixedHeight(36)
        btn_close.setCursor(Qt.PointingHandCursor)
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

    def get_font_scale(self) -> str:
        for scale_key, btn in self.font_scale_buttons.items():
            if btn.isChecked():
                return scale_key
        return "standard"

    def set_font_scale(self, scale: str, base_px: int | None = None):
        norm_scale = scale if scale in self.font_scale_buttons else "standard"
        for key, btn in self.font_scale_buttons.items():
            btn.setChecked(key == norm_scale)
        px = base_px if base_px is not None else theme.resolve_typography(norm_scale).fs_base
        scale_label = theme.FONT_SCALE_LABELS.get(norm_scale, "标准")
        self.lbl_font_scale_status.setText(f"当前：{scale_label} ({px}px)")

    def _build_stepper_spinbox(self, spinbox: QSpinBox, *, enabled: bool = True):
        """将 QSpinBox 包装为统一 Fluent 外置步进按钮样式。"""
        spinbox.setObjectName("settingsWorkerSpin")
        spinbox.setButtonSymbols(QAbstractSpinBox.NoButtons)
        spinbox.setFixedHeight(32)

        stepper = QFrame()
        stepper.setObjectName("settingsSpinStepper")
        stepper.setFixedHeight(32)
        stepper.setEnabled(enabled)
        stepper_layout = QVBoxLayout(stepper)
        stepper_layout.setContentsMargins(0, 0, 0, 0)
        stepper_layout.setSpacing(0)

        btn_up = QToolButton()
        btn_down = QToolButton()
        for btn, direction, arrow_type in (
            (btn_up, "up", Qt.UpArrow),
            (btn_down, "down", Qt.DownArrow),
        ):
            btn.setObjectName("settingsSpinStepBtn")
            btn.setProperty("direction", direction)
            btn.setArrowType(arrow_type)
            btn.setCursor(Qt.PointingHandCursor)
            btn.setFocusPolicy(Qt.NoFocus)
            btn.setEnabled(enabled)
            stepper_layout.addWidget(btn)

        btn_up.clicked.connect(spinbox.stepUp)
        btn_down.clicked.connect(spinbox.stepDown)

        spin_layout = QHBoxLayout()
        spin_layout.setContentsMargins(0, 0, 0, 0)
        spin_layout.setSpacing(0)
        spin_layout.addWidget(spinbox)
        spin_layout.addWidget(stepper)
        return spin_layout, stepper, btn_up, btn_down

    def set_parallel_worker_controls_enabled(self, enabled):
        self.spin_parallel_workers.setEnabled(enabled)
        self.parallel_worker_stepper.setEnabled(enabled)
        self.btn_parallel_worker_up.setEnabled(enabled)
        self.btn_parallel_worker_down.setEnabled(enabled)
