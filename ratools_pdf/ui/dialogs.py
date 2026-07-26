import html
import os
import re

from PySide6.QtCore import QPoint, Qt
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QAbstractItemView, QAbstractSpinBox, QButtonGroup, QCheckBox, QDialog, QFileDialog, QFrame, QGraphicsDropShadowEffect,
    QHBoxLayout, QHeaderView, QLabel, QLineEdit, QPushButton, QRadioButton, QScrollArea,
    QSizePolicy, QSplitter, QSpinBox, QTableWidget, QTableWidgetItem, QTextEdit, QToolButton,
    QVBoxLayout, QWidget,
)

from ratools_pdf.config.features import ENABLE_UPDATE_CHECK
from ratools_pdf.config.paths import get_app_dir, get_resource_path
from ratools_pdf.config.version import get_display_version
from ratools_pdf.ui.log_view_model import (
    build_log_summary_items,
    filter_log_summary_items,
)
from ratools_pdf.ui import win32
from ratools_pdf.ui.platform import is_win11, should_use_manual_dialog_shadow
from ratools_pdf.ui.theme import active_palette, log_status_colors


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


class IODataWizardDialog(FramelessDraggableDialog):
    def __init__(self, data_kind="bookmarks", file_count=0, preview_callback=None, parent=None):
        self.data_kind = "links" if data_kind == "links" else "bookmarks"
        self.data_label = "链接" if self.data_kind == "links" else "书签"
        self.data_type = "JSON" if self.data_kind == "links" else "CSV"

        super().__init__(f"{self.data_label}数据导入/导出", parent)
        self.resize(760, 540)
        self.preview_callback = preview_callback
        self.file_count = file_count

        self.content_layout.setSpacing(12)

        title = QLabel(f"配置{self.data_label}数据任务")
        title.setObjectName("dialogHeading")
        self.content_layout.addWidget(title)

        intro = QLabel(
            f"当前队列包含 {file_count} 个 PDF。选择操作方向，再确认目录匹配结果。"
            f"数据文件格式为 {self.data_type}。"
        )
        intro.setWordWrap(True)
        intro.setObjectName("dialogMuted")
        self.content_layout.addWidget(intro)

        options_row = QHBoxLayout()
        options_row.setSpacing(12)

        direction_block = QVBoxLayout()
        direction_block.setSpacing(6)
        direction_title = QLabel("操作方向")
        direction_title.setObjectName("dialogCaption")
        direction_choices = QHBoxLayout()
        direction_choices.setSpacing(8)
        self.direction_group = QButtonGroup(self)
        self.radio_export = QPushButton("导出数据")
        self.radio_import = QPushButton("导入数据")
        self.direction_group.addButton(self.radio_export)
        self.direction_group.addButton(self.radio_import)
        for btn in [self.radio_export, self.radio_import]:
            btn.setCheckable(True)
            btn.setObjectName("choiceToggleBtn")
            btn.setCursor(Qt.PointingHandCursor)
            direction_choices.addWidget(btn)
        direction_block.addWidget(direction_title)
        direction_block.addLayout(direction_choices)

        options_row.addLayout(direction_block, 1)
        self.content_layout.addLayout(options_row)

        # 链接专属选项：链接范围（全部/仅外链）与导入模式（覆盖/增量）。
        self.radio_scope_all = None
        self.radio_scope_external = None
        self.radio_mode_overwrite = None
        self.radio_mode_incremental = None
        if self.data_kind == "links":
            link_options_row = QHBoxLayout()
            link_options_row.setSpacing(12)

            scope_block = QVBoxLayout()
            scope_block.setSpacing(6)
            scope_title = QLabel("链接范围")
            scope_title.setObjectName("dialogCaption")
            scope_choices = QHBoxLayout()
            scope_choices.setSpacing(8)
            self.scope_group = QButtonGroup(self)
            self.radio_scope_all = QPushButton("全部链接")
            self.radio_scope_external = QPushButton("仅外部链接")
            self.radio_scope_external.setToolTip("仅处理网页/邮件 (URI)、跳转其它 PDF (GOTOR) 和打开外部文件 (LAUNCH) 的链接，忽略文档内跳转。")
            self.scope_group.addButton(self.radio_scope_all)
            self.scope_group.addButton(self.radio_scope_external)
            for btn in [self.radio_scope_all, self.radio_scope_external]:
                btn.setCheckable(True)
                btn.setObjectName("choiceToggleBtn")
                btn.setCursor(Qt.PointingHandCursor)
                scope_choices.addWidget(btn)
            scope_block.addWidget(scope_title)
            scope_block.addLayout(scope_choices)

            self.mode_block_widget = QWidget()
            mode_block = QVBoxLayout(self.mode_block_widget)
            mode_block.setContentsMargins(0, 0, 0, 0)
            mode_block.setSpacing(6)
            mode_title = QLabel("导入方式")
            mode_title.setObjectName("dialogCaption")
            mode_choices = QHBoxLayout()
            mode_choices.setSpacing(8)
            self.mode_group = QButtonGroup(self)
            self.radio_mode_overwrite = QPushButton("覆盖")
            self.radio_mode_incremental = QPushButton("增量")
            self.radio_mode_overwrite.setToolTip("先移除 PDF 中所选范围内的现有链接，再写入导入的链接。")
            self.radio_mode_incremental.setToolTip("保留 PDF 中现有链接，仅追加未与现有链接重叠的新链接。")
            self.mode_group.addButton(self.radio_mode_overwrite)
            self.mode_group.addButton(self.radio_mode_incremental)
            for btn in [self.radio_mode_overwrite, self.radio_mode_incremental]:
                btn.setCheckable(True)
                btn.setObjectName("choiceToggleBtn")
                btn.setCursor(Qt.PointingHandCursor)
                mode_choices.addWidget(btn)
            mode_block.addWidget(mode_title)
            mode_block.addLayout(mode_choices)

            link_options_row.addLayout(scope_block, 1)
            link_options_row.addWidget(self.mode_block_widget, 1)
            self.content_layout.addLayout(link_options_row)

            self.radio_scope_all.setChecked(True)
            self.radio_mode_overwrite.setChecked(True)

        directory_card = QFrame()
        directory_card.setObjectName("wizardCard")
        directory_layout = QVBoxLayout(directory_card)
        directory_layout.setContentsMargins(12, 10, 12, 10)
        directory_layout.setSpacing(8)
        directory_title = QLabel("数据目录")
        directory_title.setObjectName("dialogCaption")
        directory_layout.addWidget(directory_title)

        dir_row = QHBoxLayout()
        dir_row.setSpacing(8)
        self.dir_edit = QLineEdit()
        self.dir_edit.setObjectName("settingsPathEdit")
        self.dir_edit.setPlaceholderText(f"选择 {self.data_type} 数据目录")
        self.dir_edit.setFixedHeight(36)
        self.btn_browse = QPushButton("选择目录")
        self.btn_browse.setObjectName("dialogSecondaryBtn")
        self.btn_browse.setFixedHeight(36)
        dir_row.addWidget(self.dir_edit, 1)
        dir_row.addWidget(self.btn_browse)
        directory_layout.addLayout(dir_row)

        self.output_hint = QLabel("")
        self.output_hint.setWordWrap(True)
        self.output_hint.setObjectName("dialogMuted")
        directory_layout.addWidget(self.output_hint)
        self.content_layout.addWidget(directory_card)

        preview_title_row = QHBoxLayout()
        preview_title_row.setSpacing(10)
        preview_title = QLabel("匹配预览")
        preview_title.setObjectName("dialogSubTitle")
        preview_title_row.addWidget(preview_title)
        preview_title_row.addStretch()
        self.content_layout.addLayout(preview_title_row)

        self.summary_label = QLabel("等待选择数据目录")
        self.summary_label.setWordWrap(True)
        self.summary_label.setObjectName("dialogInfoPanel")
        self.content_layout.addWidget(self.summary_label)

        self.preview_empty_label = QLabel("选择数据目录后，将在这里显示每个 PDF 的数据文件匹配结果。")
        self.preview_empty_label.setAlignment(Qt.AlignCenter)
        self.preview_empty_label.setWordWrap(True)
        self.preview_empty_label.setMinimumHeight(150)
        self.preview_empty_label.setObjectName("dialogEmptyPanel")
        self.content_layout.addWidget(self.preview_empty_label, 1)

        self.preview_table = QTableWidget()
        self.preview_table.setObjectName("previewTable")
        self.preview_table.setColumnCount(4)
        self.preview_table.setHorizontalHeaderLabels(["PDF", "类型", "状态", "数据文件"])
        self.preview_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.preview_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.preview_table.setAlternatingRowColors(True)
        self.preview_table.verticalHeader().setVisible(False)
        self.preview_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.preview_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.preview_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.preview_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)
        self.preview_table.setMinimumHeight(170)
        self.content_layout.addWidget(self.preview_table, 1)

        btn_row = QHBoxLayout()
        btn_row.addStretch()
        self.btn_cancel = QPushButton("取消")
        self.btn_cancel.setObjectName("dialogSecondaryBtn")
        self.btn_cancel.setFixedHeight(32)
        self.btn_confirm = QPushButton("开始导出")
        self.btn_confirm.setObjectName("dialogPrimaryBtn")
        self.btn_confirm.setFixedHeight(32)
        self.btn_confirm.setEnabled(False)
        btn_row.addWidget(self.btn_cancel)
        btn_row.addWidget(self.btn_confirm)
        self.content_layout.addLayout(btn_row)

        self.radio_export.setChecked(True)

        for btn in [self.radio_export, self.radio_import]:
            btn.toggled.connect(self.refresh_preview)
        for btn in [self.radio_scope_all, self.radio_scope_external,
                    self.radio_mode_overwrite, self.radio_mode_incremental]:
            if btn is not None:
                btn.toggled.connect(self.refresh_preview)
        self.dir_edit.textChanged.connect(self.refresh_preview)
        self.btn_browse.clicked.connect(self.browse_directory)
        self.btn_cancel.clicked.connect(self.reject)
        self.btn_confirm.clicked.connect(self.accept)
        self.refresh_preview()

    def get_action_type(self):
        action_types = self.get_action_types()
        return action_types[0] if action_types else ""

    def get_action_types(self):
        direction = "import" if self.radio_import.isChecked() else "export"
        return [f"{direction}_{self.data_kind}"]

    def get_link_scope(self):
        if self.radio_scope_external is not None and self.radio_scope_external.isChecked():
            return "external"
        return "all"

    def get_link_mode(self):
        if self.radio_mode_incremental is not None and self.radio_mode_incremental.isChecked():
            return "incremental"
        return "overwrite"

    def get_selected_directory(self):
        return self.dir_edit.text().strip()

    def browse_directory(self):
        action_name = "导入" if self.radio_import.isChecked() else "导出"
        selected = QFileDialog.getExistingDirectory(self, f"请选择 {self.data_type} 数据{action_name}目录")
        if selected:
            self.dir_edit.setText(selected)

    def refresh_preview(self):
        action_types = self.get_action_types()
        data_kind = self.data_label
        data_type = self.data_type
        is_export = self.radio_export.isChecked()
        dir_path = self.get_selected_directory()
        self.btn_confirm.setText("开始导出" if is_export else "开始导入")

        # 导入方式（覆盖/增量）仅在导入方向下有意义。
        if getattr(self, "mode_block_widget", None) is not None:
            self.mode_block_widget.setVisible(not is_export)

        self.output_hint.setText(
            f"将在所选目录生成与 PDF 文件名匹配的 {data_type} 文件，并保留相对目录层级。"
            if is_export
            else "将从所选目录读取匹配的数据文件；导入后的 PDF 保存到 RATools_导入完成。"
        )

        if not dir_path:
            self.preview_table.setRowCount(0)
            self.preview_table.hide()
            self.preview_empty_label.setText("选择数据目录后，将在这里显示每个 PDF 的数据文件匹配结果。")
            self.preview_empty_label.show()
            self.summary_label.setText("等待选择数据目录")
            self.btn_confirm.setEnabled(False)
            return

        try:
            rows = self.preview_callback(action_types, dir_path) if self.preview_callback else []
        except Exception as exc:
            self.preview_table.setRowCount(0)
            self.preview_table.hide()
            self.preview_empty_label.setText(f"无法生成预览：{exc}")
            self.preview_empty_label.show()
            self.summary_label.setText(f"无法生成预览：{exc}")
            self.btn_confirm.setEnabled(False)
            return

        self.preview_empty_label.hide()
        self.preview_table.show()
        self.preview_table.setRowCount(len(rows))
        matched = 0
        for row_idx, row in enumerate(rows):
            status = row.get("status", "")
            if status == "已匹配":
                matched += 1
            values = [row.get("file_name", ""), row.get("data_label", data_kind), status, row.get("data_path", "")]
            for col_idx, value in enumerate(values):
                item = QTableWidgetItem(str(value))
                palette = active_palette()
                if status == "未找到":
                    item.setForeground(QColor(palette.warning_text))
                elif status in ["已匹配", "将生成"]:
                    item.setForeground(QColor(palette.success_text))
                self.preview_table.setItem(row_idx, col_idx, item)

        if is_export:
            self.summary_label.setText(f"将为 {self.file_count} 个 PDF 导出{data_kind}数据，共 {len(rows)} 个数据文件。")
            self.btn_confirm.setEnabled(len(rows) > 0)
        else:
            missing = len(rows) - matched
            self.summary_label.setText(f"已匹配 {matched} / {len(rows)} 个 {data_type} 文件；未找到 {missing} 个，执行时会跳过。")
            self.btn_confirm.setEnabled(matched > 0)


class LogDialog(FramelessDraggableDialog):
    def __init__(self, parent=None):
        super().__init__("📝 处理日志记录", parent)
        self.resize(860, 620)

        self.raw_log_text = ""
        self.filter_mode = "all"
        self.summary_items = []
        self.filtered_items = []

        self.content_layout.setSpacing(12)

        stats_row = QHBoxLayout()
        stats_row.setContentsMargins(0, 0, 0, 0)
        stats_row.setSpacing(10)
        self.stat_total_card, self.stat_total_title, self.stat_total_value = self._create_stat_card(stats_row, "总记录")
        self.stat_success_card, self.stat_success_title, self.stat_success_value = self._create_stat_card(stats_row, "成功")
        self.stat_failure_card, self.stat_failure_title, self.stat_failure_value = self._create_stat_card(stats_row, "失败")
        self.stat_precheck_card, self.stat_precheck_title, self.stat_precheck_value = self._create_stat_card(stats_row, "建议处理")
        self.stat_skip_card, self.stat_skip_title, self.stat_skip_value = self._create_stat_card(stats_row, "跳过")
        self.content_layout.addLayout(stats_row)

        filter_row = QHBoxLayout()
        filter_row.setContentsMargins(0, 0, 0, 0)
        filter_row.setSpacing(8)
        self.filter_buttons = {}
        for mode, label in [
            ("all", "全部"),
            ("success", "仅成功"),
            ("failure", "仅失败"),
            ("precheck", "仅预检"),
            ("skip", "仅跳过"),
        ]:
            btn = QPushButton(label)
            btn.setCheckable(True)
            btn.setObjectName("dialogSecondaryBtn")
            btn.clicked.connect(lambda _checked, current=mode: self.set_filter_mode(current))
            self.filter_buttons[mode] = btn
            filter_row.addWidget(btn)
        filter_row.addStretch()
        self.content_layout.addLayout(filter_row)

        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("搜索文件名、状态或结果关键字")
        self.search_edit.setObjectName("settingsPathEdit")
        self.search_edit.textChanged.connect(self.refresh_view)
        self.content_layout.addWidget(self.search_edit)

        log_splitter = QSplitter(Qt.Vertical)
        log_splitter.setObjectName("logSplitter")

        table_frame = QFrame()
        table_frame.setObjectName("logSummaryFrame")
        table_layout = QVBoxLayout(table_frame)
        table_layout.setContentsMargins(1, 1, 1, 1)
        table_layout.setSpacing(0)

        self.summary_table = QTableWidget()
        self.summary_table.setObjectName("logSummaryTable")
        self.summary_table.setFrameShape(QFrame.NoFrame)
        self.summary_table.setColumnCount(6)
        self.summary_table.setHorizontalHeaderLabels(["时间", "状态", "文件", "耗时", "输出", "结果/变化"])
        self.summary_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.summary_table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.summary_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.summary_table.setAlternatingRowColors(True)
        self.summary_table.setSortingEnabled(True)
        self.summary_table.setShowGrid(False)
        self.summary_table.verticalHeader().setVisible(False)
        self.summary_table.horizontalHeader().setStretchLastSection(True)
        self.summary_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.summary_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.summary_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.summary_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.summary_table.horizontalHeader().setSectionResizeMode(4, QHeaderView.Stretch)
        self.summary_table.horizontalHeader().setSectionResizeMode(5, QHeaderView.Stretch)
        self.summary_table.itemSelectionChanged.connect(self._show_selected_detail)
        table_layout.addWidget(self.summary_table)
        log_splitter.addWidget(table_frame)

        self.text_edit = QTextEdit()
        self.text_edit.setObjectName("logDetailTextEdit")
        self.text_edit.setReadOnly(True)
        self.text_edit.setLineWrapMode(QTextEdit.NoWrap)
        log_splitter.addWidget(self.text_edit)
        log_splitter.setStretchFactor(0, 3)
        log_splitter.setStretchFactor(1, 2)
        self.content_layout.addWidget(log_splitter)

        btn_layout = QHBoxLayout()
        self.btn_export = QPushButton("⬇️ 导出日志")
        self.btn_export.setObjectName("dialogPrimaryBtn")
        self.btn_export_precheck = QPushButton("⬇ 导出预检结果")
        self.btn_export_precheck.setObjectName("dialogPrimaryBtn")
        self.btn_close = QPushButton("关闭")
        self.btn_close.setObjectName("dialogSecondaryBtn")
        self.btn_close.clicked.connect(self.accept)

        btn_layout.addStretch()
        btn_layout.addWidget(self.btn_export_precheck)
        btn_layout.addWidget(self.btn_export)
        btn_layout.addWidget(self.btn_close)
        self.content_layout.addLayout(btn_layout)

        self.set_filter_mode("all")

    def _create_stat_card(self, parent_layout, title):
        card = QFrame()
        card.setObjectName("statCard")
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(10, 6, 10, 6)
        card_layout.setSpacing(4)

        title_label = QLabel(title)
        title_label.setObjectName("statCardTitle")
        value_label = QLabel("0")
        value_label.setObjectName("statCardValue")

        card_layout.addWidget(title_label)
        card_layout.addWidget(value_label)
        parent_layout.addWidget(card)
        return card, title_label, value_label

    @staticmethod
    def _highlight_html(text, query):
        if not query:
            return html.escape(text)
        pattern = re.compile(re.escape(query), re.IGNORECASE)
        last = 0
        parts = []
        for match in pattern.finditer(text):
            parts.append(html.escape(text[last:match.start()]))
            parts.append(f"<mark>{html.escape(match.group(0))}</mark>")
            last = match.end()
        parts.append(html.escape(text[last:]))
        return "".join(parts)

    def set_filter_mode(self, mode):
        self.filter_mode = mode
        for key, btn in self.filter_buttons.items():
            btn.setChecked(key == mode)
        self.refresh_view()

    def set_log_text(self, raw_text):
        self.set_log_data(raw_text, [])

    def set_log_data(self, raw_text, structured_rows=None):
        self.raw_log_text = raw_text or ""
        self.summary_items = build_log_summary_items(self.raw_log_text, structured_rows or [])
        self._update_stats()
        self.refresh_view()

    def _update_stats(self):
        total = len(self.summary_items)
        success = sum(1 for item in self.summary_items if "success" in item["tags"])
        failure = sum(1 for item in self.summary_items if "failure" in item["tags"])
        precheck = sum(1 for item in self.summary_items if item["status"] == "建议处理")
        skip = sum(1 for item in self.summary_items if "skip" in item["tags"])

        self.stat_total_value.setText(str(total))
        self.stat_success_value.setText(str(success))
        self.stat_failure_value.setText(str(failure))
        self.stat_precheck_value.setText(str(precheck))
        self.stat_skip_value.setText(str(skip))

    def refresh_view(self):
        query = self.search_edit.text().strip()
        self.filtered_items = filter_log_summary_items(self.summary_items, self.filter_mode, query)
        self._populate_summary_table(self.filtered_items)

        if not self.filtered_items:
            pal = active_palette()
            self.text_edit.setHtml(
                f"<div style='color:{pal.text_faint}; font-family:Consolas,\'Courier New\',monospace; font-size:12px;'>暂无匹配日志...</div>"
            )
            return

        self.summary_table.selectRow(0)
        self._show_detail_for_item(self.filtered_items[0])

    def _populate_summary_table(self, items):
        self.summary_table.setSortingEnabled(False)
        self.summary_table.setRowCount(len(items))
        for row_index, item in enumerate(items):
            values = [
                item["time"],
                item["status"],
                item["file_name"],
                item["duration_text"],
                item["output_name"],
                item["result"],
            ]
            for column_index, value in enumerate(values):
                table_item = QTableWidgetItem(str(value or ""))
                if column_index == 0:
                    table_item.setData(Qt.UserRole, item)
                if column_index in (0, 1, 3):
                    table_item.setTextAlignment(Qt.AlignCenter)
                if column_index == 1:
                    self._apply_status_table_style(table_item, item.get("tags", set()))
                table_item.setToolTip(str(value or ""))
                self.summary_table.setItem(row_index, column_index, table_item)
        self.summary_table.setSortingEnabled(True)

    @staticmethod
    def _apply_status_table_style(table_item, tags):
        pal = active_palette()
        accent, bg, text = log_status_colors(pal, tags)
        if tags & {"failure", "skip", "precheck", "success"}:
            table_item.setForeground(QColor(text))
            table_item.setBackground(QColor(bg))

    def _show_selected_detail(self):
        selected = self.summary_table.selectedItems()
        if not selected:
            return
        row = selected[0].row()
        first_item = self.summary_table.item(row, 0)
        if first_item:
            self._show_detail_for_item(first_item.data(Qt.UserRole))

    def _show_detail_for_item(self, item):
        if not item:
            return
        query = self.search_edit.text().strip()
        tags = item.get("tags", set())
        pal = active_palette()
        accent, bg, _text = log_status_colors(pal, tags)

        highlighted = self._highlight_html(item.get("detail", ""), query)
        self.text_edit.setHtml(
            f"<div style='padding:10px 12px; background:{bg}; border:1px solid {pal.border}; border-left:4px solid {accent}; border-radius:8px;'>"
            f"<pre style='margin:0; white-space:pre-wrap; color:{pal.text_body}; font-family:Consolas,\"Courier New\",monospace; font-size:12px;'>{highlighted}</pre>"
            "</div>"
        )


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


class AboutDialog(FramelessDraggableDialog):
    def __init__(self, parent=None):
        super().__init__("ℹ️ 关于软件", parent)
        self.resize(520, 460)
        self.content_layout.setSpacing(16)
        self.latest_release_url = ""

        hero_card = QFrame()
        hero_card.setObjectName("aboutHeroCard")
        hero_layout = QVBoxLayout(hero_card)
        hero_layout.setContentsMargins(18, 16, 18, 16)
        hero_layout.setSpacing(6)

        brand_title = QLabel("RATools for PDF")
        brand_title.setObjectName("aboutBrandTitle")
        version_badge = QLabel(get_display_version())
        version_badge.setObjectName("aboutBadge")
        version_badge.setAlignment(Qt.AlignCenter)
        version_badge.setMaximumWidth(110)

        hero_layout.addWidget(brand_title)
        hero_layout.addWidget(version_badge, 0, Qt.AlignLeft)
        self.content_layout.addWidget(hero_card)

        intro_text = QLabel(
            "用于RA递交资料整理的PDF处理工具，"
            "帮助用户以更稳定的方式完成eCTD场景下常见的批量标准化操作。"
        )
        intro_text.setWordWrap(True)
        intro_text.setObjectName("aboutIntro")
        self.content_layout.addWidget(intro_text)

        features_title = QLabel("核心功能")
        features_title.setObjectName("aboutTitle")
        self.content_layout.addWidget(features_title)

        features_text = QLabel(
            "• 批量导入PDF文件或文件夹\n"
            "• 按模块勾选规则，支持中国eCTD/美国eCTD预设\n"
            "• 覆盖文档属性、书签、链接、动态内容与附件等常见合规项\n"
            "• 输出处理日志，便于复核与追踪"
        )
        features_text.setWordWrap(True)
        features_text.setObjectName("aboutText")
        self.content_layout.addWidget(features_text)

        tech_card = QFrame()
        tech_card.setObjectName("aboutInfoCard")
        tech_layout = QVBoxLayout(tech_card)
        tech_layout.setContentsMargins(16, 14, 16, 14)
        tech_layout.setSpacing(4)

        tech_title = QLabel("技术与许可")
        tech_title.setObjectName("aboutTitle")
        tech_detail = QLabel(
            "基于PySide6、PyMuPDF及qpdf等项目构建\n"
            "项目源码遵循GNU AGPL v3开源协议\n"
            "第三方组件许可与源码说明见 THIRD_PARTY_NOTICES.md"
        )
        tech_detail.setWordWrap(True)
        tech_detail.setObjectName("aboutText")
        tech_layout.addWidget(tech_title)
        tech_layout.addWidget(tech_detail)
        self.content_layout.addWidget(tech_card)

        if ENABLE_UPDATE_CHECK:
            update_card = QFrame()
            update_card.setObjectName("aboutInfoCard")
            update_layout = QVBoxLayout(update_card)
            update_layout.setContentsMargins(16, 14, 16, 14)
            update_layout.setSpacing(8)

            update_title = QLabel("更新")
            update_title.setObjectName("aboutTitle")
            self.update_status_label = QLabel("可手动检查 GitHub Releases 中的最新版本。")
            self.update_status_label.setWordWrap(True)
            self.update_status_label.setObjectName("aboutText")

            update_button_layout = QHBoxLayout()
            update_button_layout.setContentsMargins(0, 4, 0, 0)
            update_button_layout.setSpacing(8)
            self.btn_check_updates = QPushButton("检查更新")
            self.btn_check_updates.setObjectName("dialogPrimaryBtn")
            self.btn_open_release = QPushButton("打开发布页")
            self.btn_open_release.setObjectName("dialogSecondaryBtn")
            self.btn_open_release.hide()
            update_button_layout.addWidget(self.btn_check_updates)
            update_button_layout.addWidget(self.btn_open_release)
            update_button_layout.addStretch()

            update_layout.addWidget(update_title)
            update_layout.addWidget(self.update_status_label)
            update_layout.addLayout(update_button_layout)
            self.content_layout.addWidget(update_card)

        self.content_layout.addStretch()

        btn_close = QPushButton("关 闭")
        btn_close.setObjectName("dialogSecondaryBtn")
        btn_close.setFixedHeight(36)
        btn_close.clicked.connect(self.accept)
        self.content_layout.addWidget(btn_close)

    def set_update_checking(self):
        # NoUpdate 变体（ENABLE_UPDATE_CHECK=False）不创建更新区控件
        if not hasattr(self, "btn_check_updates"):
            return
        self.btn_check_updates.setEnabled(False)
        self.update_status_label.setText("正在检查更新...")
        self.btn_open_release.hide()

    def set_update_result(self, message, release_url=""):
        if not hasattr(self, "btn_check_updates"):
            return
        self.btn_check_updates.setEnabled(True)
        self.update_status_label.setText(message)
        self.latest_release_url = release_url
        self.btn_open_release.setVisible(bool(release_url))


# ================== 自定义组件与主窗口 ==================
