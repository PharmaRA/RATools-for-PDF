
from PySide6.QtCore import Qt
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QAbstractItemView, QButtonGroup, QFileDialog, QFrame, QHBoxLayout, QHeaderView, QLabel, QLineEdit, QPushButton, QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget,
)

from ratools_pdf.ui.log_view_model import (
    build_log_summary_items,
    filter_log_summary_items,
)
from ratools_pdf.ui.theme import active_palette
from ratools_pdf.ui.dialogs.base import FramelessDraggableDialog


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
