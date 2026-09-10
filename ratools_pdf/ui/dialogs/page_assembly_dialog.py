"""页面装配与拆分工作台：可视化多文档合并、页面抽取、排序、倒序与分卷拆分。"""

import os
from typing import List, Optional

from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtWidgets import (
    QAbstractItemView,
    QButtonGroup,
    QComboBox,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QRadioButton,
    QSpinBox,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from ratools_pdf.pdf import qpdf
from ratools_pdf.ui.dialogs.base import FramelessDraggableDialog


class PageAssemblyWorker(QThread):
    """后台执行页面装配或拆分任务的异步工作线程。"""

    progress = Signal(str)
    finished_signal = Signal(bool, str, str)  # success, message, target_path

    def __init__(
        self,
        mode: str,
        specs: List[qpdf.PageSpec],
        output_path: str,
        linearize: bool = True,
        object_streams: bool = True,
        split_pages_per_file: Optional[int] = None,
        parent=None,
    ):
        super().__init__(parent)
        self.mode = mode
        self.specs = specs
        self.output_path = output_path
        self.linearize = linearize
        self.object_streams = "generate" if object_streams else None
        self.split_pages_per_file = split_pages_per_file

    def run(self):
        try:
            if self.mode in ("merge", "extract"):
                self.progress.emit("正在装配页面...")
                res = qpdf.assemble_pages(
                    specs=self.specs,
                    output_pdf=self.output_path,
                    linearize=self.linearize,
                    object_streams=self.object_streams,
                )
                if res.is_success:
                    self.finished_signal.emit(
                        True,
                        f"✅ 成功生成装配文件：\n{self.output_path}",
                        self.output_path,
                    )
                else:
                    self.finished_signal.emit(
                        False, f"❌ 装配失败：{res.stderr or '未知错误'}", ""
                    )

            elif self.mode == "split":
                self.progress.emit("正在执行拆分...")
                for spec in self.specs:
                    base_name = os.path.splitext(os.path.basename(spec.file_path))[0]
                    pattern = os.path.join(self.output_path, f"{base_name}_part_%d.pdf")
                    qpdf.split_pages(
                        input_pdf=spec.file_path,
                        output_pattern=pattern,
                        pages_per_file=self.split_pages_per_file,
                        password=spec.password,
                    )
                self.finished_signal.emit(
                    True,
                    f"✅ 成功拆分文件至目录：\n{self.output_path}",
                    self.output_path,
                )
        except Exception as e:
            self.finished_signal.emit(False, f"❌ 执行异常：{str(e)}", "")


class PageAssemblyDialog(FramelessDraggableDialog):
    """页面装配与拆分工作台对话框。"""

    def __init__(self, initial_files: Optional[List[str]] = None, parent=None):
        super().__init__("📑 页面装配与拆分工作台", parent)
        self.resize(920, 680)
        self.setMinimumSize(820, 560)
        self.worker: Optional[PageAssemblyWorker] = None
        self.last_output_path = ""

        self.content_layout.setSpacing(12)

        self._build_mode_selector()
        self._build_table_section()
        self._build_quick_actions()
        self._build_split_settings()
        self._build_output_section()
        self._build_footer()

        self._on_mode_changed()
        if initial_files:
            self.add_files(initial_files)

    def _build_mode_selector(self):
        """顶部操作模式选择单选组。"""
        mode_card = QFrame()
        mode_card.setObjectName("wizardCard")
        layout = QHBoxLayout(mode_card)
        layout.setContentsMargins(14, 10, 14, 10)
        layout.setSpacing(20)

        lbl = QLabel("工作模式：")
        lbl.setStyleSheet("font-weight: 700;")
        layout.addWidget(lbl)

        self.mode_group = QButtonGroup(self)
        self.rb_merge = QRadioButton("多文件合并与装配")
        self.rb_extract = QRadioButton("单文件提取与重排")
        self.rb_split = QRadioButton("批量页面拆分")
        self.rb_merge.setChecked(True)

        self.mode_group.addButton(self.rb_merge, 0)
        self.mode_group.addButton(self.rb_extract, 1)
        self.mode_group.addButton(self.rb_split, 2)

        layout.addWidget(self.rb_merge)
        layout.addWidget(self.rb_extract)
        layout.addWidget(self.rb_split)
        layout.addStretch()

        self.mode_group.idClicked.connect(lambda _id: self._on_mode_changed())
        self.content_layout.addWidget(mode_card)

    def _build_table_section(self):
        """表格与文件操作按钮区。"""
        table_container = QWidget()
        layout = QHBoxLayout(table_container)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(10)

        self.table = QTableWidget()
        self.table.setObjectName("previewTable")
        self.table.setColumnCount(6)
        self.table.setHorizontalHeaderLabels(
            ["#", "文件名", "总页数", "提取页面范围 (如 1-5, z-1)", "页面旋转", "完整路径"]
        )
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Fixed)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeToContents)
        self.table.setColumnWidth(0, 40)
        self.table.setColumnHidden(5, True)  # 隐藏完整路径列
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.table.setAlternatingRowColors(True)
        self.table.itemChanged.connect(self._on_table_item_changed)

        # 启用拖拽添加文件
        self.table.setAcceptDrops(True)
        self.table.dragEnterEvent = self._drag_enter_event
        self.table.dragMoveEvent = self._drag_move_event
        self.table.dropEvent = self._drop_event

        layout.addWidget(self.table, stretch=1)

        # 侧边工具按钮
        btn_bar = QVBoxLayout()
        btn_bar.setSpacing(8)

        self.btn_add = QPushButton("+ 添加文件")
        self.btn_add.clicked.connect(self._on_add_files_clicked)
        self.btn_remove = QPushButton("- 移除选中")
        self.btn_remove.clicked.connect(self._on_remove_file_clicked)
        self.btn_move_up = QPushButton("↑ 上移")
        self.btn_move_up.clicked.connect(self._on_move_up_clicked)
        self.btn_move_down = QPushButton("↓ 下移")
        self.btn_move_down.clicked.connect(self._on_move_down_clicked)
        self.btn_clear = QPushButton("清空")
        self.btn_clear.clicked.connect(self._on_clear_clicked)

        for btn in (
            self.btn_add,
            self.btn_remove,
            self.btn_move_up,
            self.btn_move_down,
            self.btn_clear,
        ):
            btn.setFixedWidth(100)
            btn_bar.addWidget(btn)

        btn_bar.addStretch()
        layout.addLayout(btn_bar)

        self.content_layout.addWidget(table_container, stretch=1)

    def _build_quick_actions(self):
        """快捷页面范围设定栏。"""
        bar = QWidget()
        layout = QHBoxLayout(bar)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)

        lbl = QLabel("当前选中行快捷范围：")
        lbl.setStyleSheet("color: gray; font-size: 12px;")
        layout.addWidget(lbl)

        self.btn_range_all = QPushButton("全部页面 (1-z)")
        self.btn_range_all.clicked.connect(lambda: self._set_current_range("1-z"))
        self.btn_range_reverse = QPushButton("全部倒序 (z-1)")
        self.btn_range_reverse.clicked.connect(lambda: self._set_current_range("z-1"))
        self.btn_range_odd = QPushButton("仅奇数页")
        self.btn_range_odd.clicked.connect(lambda: self._set_current_range("1-z:odd"))
        self.btn_range_even = QPushButton("仅偶数页")
        self.btn_range_even.clicked.connect(lambda: self._set_current_range("1-z:even"))

        for btn in (
            self.btn_range_all,
            self.btn_range_reverse,
            self.btn_range_odd,
            self.btn_range_even,
        ):
            btn.setFixedHeight(26)
            layout.addWidget(btn)

        layout.addStretch()
        self.quick_actions_bar = bar
        self.content_layout.addWidget(bar)

    def _build_split_settings(self):
        """拆分模式下的参数配置面板。"""
        self.split_card = QFrame()
        self.split_card.setObjectName("wizardCard")
        layout = QHBoxLayout(self.split_card)
        layout.setContentsMargins(14, 10, 14, 10)
        layout.setSpacing(18)

        lbl = QLabel("拆分方式：")
        lbl.setStyleSheet("font-weight: 700;")
        layout.addWidget(lbl)

        self.split_group = QButtonGroup(self)
        self.rb_split_fixed = QRadioButton("按固定页数分卷：每")
        self.spin_split_pages = QSpinBox()
        self.spin_split_pages.setRange(1, 9999)
        self.spin_split_pages.setValue(10)
        self.spin_split_pages.setSuffix(" 页为一卷")
        self.spin_split_pages.setFixedWidth(110)

        self.rb_split_each = QRadioButton("每一页拆分为独立 PDF 文件")
        self.rb_split_fixed.setChecked(True)

        self.split_group.addButton(self.rb_split_fixed, 0)
        self.split_group.addButton(self.rb_split_each, 1)

        layout.addWidget(self.rb_split_fixed)
        layout.addWidget(self.spin_split_pages)
        layout.addWidget(self.rb_split_each)
        layout.addStretch()

        self.content_layout.addWidget(self.split_card)

    def _build_output_section(self):
        """输出路径与优化选项配置。"""
        output_container = QWidget()
        v_layout = QVBoxLayout(output_container)
        v_layout.setContentsMargins(0, 4, 0, 4)
        v_layout.setSpacing(8)

        h_layout = QHBoxLayout()
        h_layout.setSpacing(8)
        self.lbl_output = QLabel("目标文件：")
        self.lbl_output.setFixedWidth(70)
        self.txt_output = QLineEdit()
        self.btn_browse_output = QPushButton("浏览...")
        self.btn_browse_output.clicked.connect(self._on_browse_output_clicked)

        h_layout.addWidget(self.lbl_output)
        h_layout.addWidget(self.txt_output, stretch=1)
        h_layout.addWidget(self.btn_browse_output)
        v_layout.addLayout(h_layout)

        # 优化选项
        opts_layout = QHBoxLayout()
        opts_layout.setSpacing(16)
        self.cb_linearize = QPushButton("Web 快速视图 (线性化)")
        self.cb_linearize.setCheckable(True)
        self.cb_linearize.setChecked(True)
        self.cb_object_streams = QPushButton("压缩对象流 (PDF 1.5+)")
        self.cb_object_streams.setCheckable(True)
        self.cb_object_streams.setChecked(True)

        for btn in (self.cb_linearize, self.cb_object_streams):
            btn.setObjectName("choiceToggleBtn")
            btn.setFixedHeight(28)
            opts_layout.addWidget(btn)

        opts_layout.addStretch()
        v_layout.addLayout(opts_layout)

        self.content_layout.addWidget(output_container)

    def _build_footer(self):
        """底部状态栏与主操作按钮。"""
        footer = QWidget()
        layout = QHBoxLayout(footer)
        layout.setContentsMargins(0, 8, 0, 0)
        layout.setSpacing(12)

        self.lbl_status = QLabel("就绪")
        self.lbl_status.setStyleSheet("color: gray; font-size: 12px;")
        layout.addWidget(self.lbl_status)
        layout.addStretch()

        self.btn_open_target = QPushButton("打开目标")
        self.btn_open_target.setVisible(False)
        self.btn_open_target.clicked.connect(self._on_open_target_clicked)
        layout.addWidget(self.btn_open_target)

        self.btn_cancel = QPushButton("关闭")
        self.btn_cancel.clicked.connect(self.reject)
        layout.addWidget(self.btn_cancel)

        self.btn_execute = QPushButton("开始处理")
        self.btn_execute.setObjectName("choiceToggleBtn")
        self.btn_execute.setStyleSheet(
            "QPushButton { background-color: #2563EB; color: white; font-weight: bold; border-radius: 8px; padding: 8px 18px; }"
            "QPushButton:hover { background-color: #1D4ED8; }"
            "QPushButton:disabled { background-color: #93C5FD; }"
        )
        self.btn_execute.clicked.connect(self._on_execute_clicked)
        layout.addWidget(self.btn_execute)

        self.content_layout.addWidget(footer)

    def _on_mode_changed(self):
        """模式改变时切换 UI 显示与文本。"""
        mode_id = self.mode_group.checkedId()
        is_split = mode_id == 2
        is_extract = mode_id == 1

        self.split_card.setVisible(is_split)
        self.quick_actions_bar.setVisible(not is_split)

        if is_split:
            self.lbl_output.setText("输出目录：")
            self.btn_execute.setText("开始拆分")
            self.txt_output.setPlaceholderText("选择拆分后输出保存的文件夹目录...")
        elif is_extract:
            self.lbl_output.setText("输出文件：")
            self.btn_execute.setText("开始提取")
            self.txt_output.setPlaceholderText("选择输出的 PDF 完整路径...")
        else:
            self.lbl_output.setText("输出文件：")
            self.btn_execute.setText("开始合并")
            self.txt_output.setPlaceholderText("选择合并后生成的 PDF 完整路径...")

        self._update_status_summary()

    def add_files(self, file_paths: List[str]):
        """向表格添加 PDF 文件。"""
        pdf_paths = [p for p in file_paths if p.lower().endswith(".pdf") and os.path.exists(p)]
        if not pdf_paths:
            return

        self.table.blockSignals(True)
        for path in pdf_paths:
            row = self.table.rowCount()
            self.table.insertRow(row)

            # 0: 序号
            item_idx = QTableWidgetItem(str(row + 1))
            item_idx.setTextAlignment(Qt.AlignCenter)
            item_idx.setFlags(item_idx.flags() & ~Qt.ItemIsEditable)
            self.table.setItem(row, 0, item_idx)

            # 1: 文件名
            name = os.path.basename(path)
            item_name = QTableWidgetItem(name)
            item_name.setToolTip(path)
            item_name.setFlags(item_name.flags() & ~Qt.ItemIsEditable)
            self.table.setItem(row, 1, item_name)

            # 2: 总页数
            total_pages = qpdf.get_pdf_page_count(path)
            item_pages = QTableWidgetItem(str(total_pages))
            item_pages.setTextAlignment(Qt.AlignCenter)
            item_pages.setFlags(item_pages.flags() & ~Qt.ItemIsEditable)
            self.table.setItem(row, 2, item_pages)

            # 3: 提取范围
            item_range = QTableWidgetItem("1-z")
            item_range.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(row, 3, item_range)

            # 4: 旋转角度
            combo_rot = QComboBox()
            combo_rot.addItems(["原样保持", "顺时针 90°", "180°", "逆时针 90°"])
            self.table.setCellWidget(row, 4, combo_rot)

            # 5: 完整路径
            item_path = QTableWidgetItem(path)
            self.table.setItem(row, 5, item_path)

        self.table.blockSignals(False)
        self._refresh_row_indices()
        self._update_status_summary()

        # 默认推断输出路径
        if not self.txt_output.text() and self.table.rowCount() > 0:
            first_path = self.table.item(0, 5).text()
            first_dir = os.path.dirname(first_path)
            mode_id = self.mode_group.checkedId()
            if mode_id == 2:
                self.txt_output.setText(first_dir)
            else:
                self.txt_output.setText(os.path.join(first_dir, "Assembled_Output.pdf"))

    def _refresh_row_indices(self):
        """重新刷新每行序号。"""
        for r in range(self.table.rowCount()):
            item = self.table.item(r, 0)
            if item:
                item.setText(str(r + 1))

    def _set_current_range(self, range_str: str):
        """设置当前选中的行范围。"""
        row = self.table.currentRow()
        if row >= 0:
            item = self.table.item(row, 3)
            if item:
                item.setText(range_str)
        else:
            # 若未选中行但有行，设置全部行
            for r in range(self.table.rowCount()):
                item = self.table.item(r, 3)
                if item:
                    item.setText(range_str)

    def _update_status_summary(self):
        """更新底部汇总状态文本。"""
        count = self.table.rowCount()
        mode_id = self.mode_group.checkedId()
        if count == 0:
            self.lbl_status.setText("请添加需要处理的 PDF 文件（支持拖拽）")
            return

        if mode_id == 2:
            self.lbl_status.setText(f"已选 {count} 个待拆分文件")
        else:
            total_est_pages = 0
            for r in range(count):
                total_p = int(self.table.item(r, 2).text() or "0")
                range_str = self.table.item(r, 3).text() if self.table.item(r, 3) else "1-z"
                selected = qpdf.parse_page_range(range_str, total_p)
                total_est_pages += len(selected)
            self.lbl_status.setText(f"已选 {count} 个文件，预计输出 {total_est_pages} 页")

    def _on_table_item_changed(self, item):
        if item.column() == 3:
            self._update_status_summary()

    def _on_add_files_clicked(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "选择 PDF 文件", "", "PDF 文件 (*.pdf)"
        )
        if files:
            self.add_files(files)

    def _on_remove_file_clicked(self):
        row = self.table.currentRow()
        if row >= 0:
            self.table.removeRow(row)
            self._refresh_row_indices()
            self._update_status_summary()

    def _on_move_up_clicked(self):
        row = self.table.currentRow()
        if row > 0:
            self._swap_rows(row, row - 1)
            self.table.selectRow(row - 1)

    def _on_move_down_clicked(self):
        row = self.table.currentRow()
        if 0 <= row < self.table.rowCount() - 1:
            self._swap_rows(row, row + 1)
            self.table.selectRow(row + 1)

    def _swap_rows(self, r1: int, r2: int):
        self.table.blockSignals(True)
        # 获取 r1 旋转
        rot1_idx = self.table.cellWidget(r1, 4).currentIndex()
        rot2_idx = self.table.cellWidget(r2, 4).currentIndex()

        for col in (1, 2, 3, 5):
            it1 = self.table.takeItem(r1, col)
            it2 = self.table.takeItem(r2, col)
            self.table.setItem(r1, col, it2)
            self.table.setItem(r2, col, it1)

        self.table.cellWidget(r1, 4).setCurrentIndex(rot2_idx)
        self.table.cellWidget(r2, 4).setCurrentIndex(rot1_idx)

        self.table.blockSignals(False)
        self._refresh_row_indices()

    def _on_clear_clicked(self):
        self.table.setRowCount(0)
        self._update_status_summary()

    def _on_browse_output_clicked(self):
        mode_id = self.mode_group.checkedId()
        if mode_id == 2:
            folder = QFileDialog.getExistingDirectory(self, "选择拆分输出目录")
            if folder:
                self.txt_output.setText(folder)
        else:
            file_path, _ = QFileDialog.getSaveFileName(
                self, "保存输出 PDF", "Assembled.pdf", "PDF 文件 (*.pdf)"
            )
            if file_path:
                self.txt_output.setText(file_path)

    def _drag_enter_event(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def _drag_move_event(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def _drop_event(self, event):
        urls = event.mimeData().urls()
        files = [u.toLocalFile() for u in urls if u.toLocalFile().lower().endswith(".pdf")]
        if files:
            self.add_files(files)
            event.acceptProposedAction()

    def _on_execute_clicked(self):
        """收集参数并启动后台工作线程。"""
        if self.table.rowCount() == 0:
            QMessageBox.warning(self, "提示", "请先添加至少一个 PDF 文件！")
            return

        out_path = self.txt_output.text().strip()
        if not out_path:
            QMessageBox.warning(self, "提示", "请先指定输出文件或目录路径！")
            return

        mode_id = self.mode_group.checkedId()
        mode_str = "split" if mode_id == 2 else ("extract" if mode_id == 1 else "merge")

        rot_map = {0: 0, 1: 90, 2: 180, 3: 270}
        specs: List[qpdf.PageSpec] = []
        for r in range(self.table.rowCount()):
            file_path = self.table.item(r, 5).text()
            range_str = self.table.item(r, 3).text() if self.table.item(r, 3) else "1-z"
            rot_idx = self.table.cellWidget(r, 4).currentIndex()
            rot_deg = rot_map.get(rot_idx, 0)
            specs.append(
                qpdf.PageSpec(
                    file_path=file_path,
                    page_range=range_str,
                    rotation=rot_deg if rot_deg != 0 else None,
                )
            )

        split_pages_per_file = None
        if mode_str == "split":
            if not os.path.exists(out_path):
                try:
                    os.makedirs(out_path, exist_ok=True)
                except Exception as e:
                    QMessageBox.critical(self, "错误", f"无法创建输出目录：{str(e)}")
                    return
            split_pages_per_file = (
                1 if self.rb_split_each.isChecked() else self.spin_split_pages.value()
            )

        self.btn_execute.setEnabled(False)
        self.btn_cancel.setEnabled(False)
        self.lbl_status.setText("⏳ 处理中，请稍候...")
        self.btn_open_target.setVisible(False)

        self.worker = PageAssemblyWorker(
            mode=mode_str,
            specs=specs,
            output_path=out_path,
            linearize=self.cb_linearize.isChecked(),
            object_streams=self.cb_object_streams.isChecked(),
            split_pages_per_file=split_pages_per_file,
            parent=self,
        )
        self.worker.progress.connect(lambda msg: self.lbl_status.setText(msg))
        self.worker.finished_signal.connect(self._on_worker_finished)
        self.worker.start()

    def _on_worker_finished(self, success: bool, message: str, target_path: str):
        self.btn_execute.setEnabled(True)
        self.btn_cancel.setEnabled(True)
        self.lbl_status.setText("处理完成" if success else "处理失败")

        if success:
            self.last_output_path = target_path
            self.btn_open_target.setVisible(True)
            QMessageBox.information(self, "完成", message)
        else:
            QMessageBox.critical(self, "错误", message)

    def _on_open_target_clicked(self):
        """打开生成的输出文件或输出文件夹。"""
        if not self.last_output_path or not os.path.exists(self.last_output_path):
            return
        if os.path.isdir(self.last_output_path):
            os.startfile(self.last_output_path)
        else:
            os.startfile(os.path.dirname(self.last_output_path))
