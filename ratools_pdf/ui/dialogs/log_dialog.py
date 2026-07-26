import html
import re

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QAbstractItemView, QFrame, QHBoxLayout, QHeaderView, QLabel, QLineEdit, QPushButton, QSplitter, QTableWidget, QTableWidgetItem, QTextEdit, QVBoxLayout,
)

from ratools_pdf.ui.log_view_model import (
    build_log_summary_items,
    filter_log_summary_items,
)
from ratools_pdf.ui.theme import active_palette, log_status_colors
from ratools_pdf.ui.dialogs.base import FramelessDraggableDialog


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
