"""Interactive, explicit selection of removable PDF objects."""

from pathlib import Path

from PySide6.QtCore import Qt, QThread, QTimer
from PySide6.QtGui import QPixmap
from PySide6.QtWidgets import (
    QCheckBox, QFileDialog, QHBoxLayout, QHeaderView, QLabel, QLineEdit,
    QPushButton, QScrollArea, QSpinBox, QSplitter, QTreeWidget,
    QTreeWidgetItem, QVBoxLayout, QWidget,
)

from ratools_pdf.pdf import watermarks
from ratools_pdf.ui.dialogs.base import FramelessDraggableDialog


class WatermarkTask(QThread):
    """Only immutable Python data crosses threads; all fitz objects stay local."""

    def __init__(self, action, args, parent=None):
        super().__init__(parent)
        self.action = action
        self.args = args
        self.result = None
        self.error = ''

    def run(self):
        try:
            self.result = self.action(*self.args, cancel=self.isInterruptionRequested)
        except watermarks.WatermarkError as error:
            self.error = str(error)
        except Exception:
            self.error = '操作失败，未能完成输出。请检查 PDF 和输出目录后重试。'


class WatermarkDialog(FramelessDraggableDialog):
    def __init__(self, initial_file=None, parent=None):
        super().__init__('去除水印', parent)
        self.resize(1120, 800)
        self.analysis = None
        self.worker = None
        self._preview_key = None
        self._viewed_pages = {}
        self._preview_requested = False
        self._generation = 0
        self._build()
        self.preview_timer = QTimer(self)
        self.preview_timer.setSingleShot(True)
        self.preview_timer.setInterval(180)
        self.preview_timer.timeout.connect(self._preview)
        self.txt_source.textChanged.connect(self._invalidate)
        self.txt_pages.textChanged.connect(self._invalidate)
        self.objects.itemChanged.connect(self._selection_changed)
        self.objects.currentItemChanged.connect(self._focus_item)
        self.page_number.valueChanged.connect(self._queue_preview)
        if initial_file:
            self.txt_source.setText(initial_file)

    def _build(self):
        hint = QLabel('选择可独立移除的批注、文字、图片或绘图对象。候选中可能包含正文；请查看前后效果后勾选。'
                      '扫描图像内的水印和复杂混排内容暂不支持。')
        hint.setWordWrap(True)
        self.content_layout.addWidget(hint)
        row = QHBoxLayout()
        self.txt_source = QLineEdit()
        self.txt_source.setPlaceholderText('选择 PDF 文件')
        self.btn_source = QPushButton('选择 PDF…')
        self.btn_source.clicked.connect(self._choose_source)
        row.addWidget(self.txt_source, 1)
        row.addWidget(self.btn_source)
        self.content_layout.addLayout(row)
        row = QHBoxLayout()
        row.addWidget(QLabel('处理页码'))
        self.txt_pages = QLineEdit()
        self.txt_pages.setPlaceholderText('例如 1-3,5；留空表示全部页面')
        self.btn_analyze = QPushButton('分析对象')
        self.btn_analyze.setObjectName('dialogPrimaryBtn')
        self.btn_analyze.clicked.connect(self._analyze)
        row.addWidget(self.txt_pages, 1)
        row.addWidget(self.btn_analyze)
        self.content_layout.addLayout(row)

        splitter = QSplitter(Qt.Horizontal)
        self.objects = QTreeWidget()
        self.objects.setHeaderLabels(['选择', '页码', '类型', '对象'])
        self.objects.setRootIsDecorated(False)
        self.objects.header().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.objects.header().setStretchLastSection(True)
        splitter.addWidget(self.objects)
        previews = QWidget()
        layout = QVBoxLayout(previews)
        controls = QHBoxLayout()
        controls.addWidget(QLabel('预览页码'))
        self.page_number = QSpinBox()
        self.page_number.setMinimum(1)
        controls.addWidget(self.page_number)
        self.btn_preview = QPushButton('刷新预览')
        self.btn_preview.clicked.connect(self._queue_preview)
        controls.addWidget(self.btn_preview)
        controls.addStretch()
        layout.addLayout(controls)
        images = QHBoxLayout()
        self.before = self._image_panel(images, '原始页面')
        self.after = self._image_panel(images, '去除勾选对象后的页面')
        layout.addLayout(images, 1)
        splitter.addWidget(previews)
        splitter.setSizes([360, 700])
        self.content_layout.addWidget(splitter, 1)
        self.lbl_notes = QLabel()
        self.lbl_notes.setWordWrap(True)
        self.content_layout.addWidget(self.lbl_notes)
        self.confirm = QCheckBox('已检查所选对象的预览，确认去除；修改后的副本不保留原数字签名的有效性。')
        self.confirm.toggled.connect(self._update_actions)
        self.content_layout.addWidget(self.confirm)
        row = QHBoxLayout()
        self.lbl_status = QLabel('选择文件与页码后，点击“分析对象”。')
        self.lbl_status.setWordWrap(True)
        row.addWidget(self.lbl_status, 1)
        self.btn_cancel = QPushButton('取消当前操作')
        self.btn_cancel.clicked.connect(self._cancel)
        self.btn_save = QPushButton('另存为 PDF…')
        self.btn_save.setObjectName('dialogPrimaryBtn')
        self.btn_save.clicked.connect(self._save)
        row.addWidget(self.btn_cancel)
        row.addWidget(self.btn_save)
        self.content_layout.addLayout(row)
        self._update_actions()

    @staticmethod
    def _image_panel(layout, title):
        panel = QVBoxLayout()
        panel.addWidget(QLabel(title))
        area = QScrollArea()
        area.setWidgetResizable(True)
        label = QLabel('分析后显示预览')
        label.setAlignment(Qt.AlignCenter)
        area.setWidget(label)
        panel.addWidget(area, 1)
        layout.addLayout(panel, 1)
        return label

    def _choose_source(self):
        path, _ = QFileDialog.getOpenFileName(self, '选择 PDF', '', 'PDF 文件 (*.pdf)')
        if path:
            self.txt_source.setText(path)

    def _invalidate(self, *_args):
        self._generation += 1
        self.preview_timer.stop()
        self.analysis = None
        self._preview_key = None
        self._viewed_pages.clear()
        self._preview_requested = False
        self.objects.clear()
        self.confirm.setChecked(False)
        self.before.setText('请重新分析')
        self.after.setText('请重新分析')
        self.lbl_notes.clear()
        self.lbl_status.setText('文件或页码已改变，请重新分析。')
        self._update_actions()

    def selected_ids(self):
        return tuple(self.objects.topLevelItem(i).data(0, Qt.UserRole)
                     for i in range(self.objects.topLevelItemCount())
                     if self.objects.topLevelItem(i).checkState(0) == Qt.Checked)

    def _busy(self):
        return self.worker is not None

    def _page_selections(self):
        selected = set(self.selected_ids())
        result = {}
        if self.analysis is not None:
            for candidate in self.analysis.candidates:
                if candidate.id in selected:
                    result.setdefault(candidate.page, []).append(candidate.id)
        return {page: tuple(ids) for page, ids in result.items()}

    def _pending_pages(self):
        return [page for page, ids in self._page_selections().items()
                if self._viewed_pages.get(page) != ids]

    def _update_actions(self, *_args):
        busy = self._busy()
        for widget in (self.txt_source, self.txt_pages, self.btn_source, self.btn_analyze,
                       self.objects, self.page_number, self.confirm):
            widget.setEnabled(not busy)
        self.btn_cancel.setEnabled(busy)
        self.btn_preview.setEnabled(not busy and self.analysis is not None)
        self.btn_save.setEnabled(not busy and self.analysis is not None and
                                 bool(self.selected_ids()) and self.confirm.isChecked() and
                                 not self._pending_pages())

    def _start(self, action, args, callback, message):
        if self._busy():
            return
        self.preview_timer.stop()
        self.lbl_status.setText(message)
        task = WatermarkTask(action, args, self)
        self.worker = task
        generation = self._generation

        def completed():
            self.worker = None
            if generation == self._generation:
                if task.error:
                    self.lbl_status.setText(task.error)
                else:
                    callback(task.result)
            task.deleteLater()
            self._update_actions()
            if self._preview_requested:
                self._preview_requested = False
                self.preview_timer.start()

        task.finished.connect(completed)
        self._update_actions()
        task.start()

    def _analyze(self):
        self._invalidate()
        self._start(watermarks.analyze, (self.txt_source.text().strip(), self.txt_pages.text()),
                    self._analyzed, '正在分析可独立移除的对象…')

    def _analyzed(self, analysis):
        self.analysis = analysis
        self.objects.blockSignals(True)
        for candidate in analysis.candidates:
            item = QTreeWidgetItem(['', str(candidate.page + 1), candidate.kind, candidate.label])
            item.setData(0, Qt.UserRole, candidate.id)
            item.setData(1, Qt.UserRole, candidate.page)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setCheckState(0, Qt.Unchecked)
            self.objects.addTopLevelItem(item)
        self.objects.blockSignals(False)
        self.page_number.blockSignals(True)
        self.page_number.setMaximum(analysis.page_count)
        self.page_number.setValue(analysis.pages[0] + 1)
        self.page_number.blockSignals(False)
        note_text = '\n'.join(analysis.notes[:4])
        if len(analysis.notes) > 4:
            note_text += f'\n另有 {len(analysis.notes) - 4} 条提示（悬停查看全部）。'
        self.lbl_notes.setText(note_text)
        self.lbl_notes.setToolTip('\n'.join(analysis.notes))
        self.lbl_status.setText(f'找到 {len(analysis.candidates)} 个可选对象，均未勾选。勾选后查看右侧变化。')
        self._queue_preview()

    def _selection_changed(self, item, _column):
        self.confirm.setChecked(False)
        self._preview_key = None
        self.page_number.setValue(item.data(1, Qt.UserRole) + 1)
        self._queue_preview()
        self._update_actions()

    def _focus_item(self, current, _previous):
        if current is not None:
            self.page_number.setValue(current.data(1, Qt.UserRole) + 1)

    def _queue_preview(self, *_args):
        if self.analysis is not None:
            self.preview_timer.start()

    def _preview(self):
        if self.analysis is None:
            return
        if self._busy():
            self._preview_requested = True
            return
        selected = self.selected_ids()
        number = self.page_number.value() - 1
        if number not in self.analysis.pages:
            self.before.setText('该页未包含在处理页码中')
            self.after.setText('请切换至处理范围内的页面')
            return

        def rendered(images):
            for label, data in zip((self.before, self.after), images):
                image = QPixmap()
                image.loadFromData(data, 'PNG')
                label.setPixmap(image.scaledToWidth(430, Qt.SmoothTransformation))
            self._preview_key = selected
            self._viewed_pages[number] = self._page_selections().get(number, ())
            pending = self._pending_pages()
            remaining = (' 请继续预览第 ' + ', '.join(str(page + 1) for page in pending) + ' 页。') if pending else ''
            self.lbl_status.setText(f'第 {number + 1} 页预览已更新；共选择 {len(selected)} 个对象。' + remaining)

        self._start(watermarks.preview, (self.analysis, selected, number), rendered, '正在生成前后预览…')

    def _save(self):
        if not self.btn_save.isEnabled():
            return
        source = Path(self.analysis.path)
        path, _ = QFileDialog.getSaveFileName(self, '另存去水印 PDF（使用新文件名）',
                                             str(source.with_name(source.stem + '_去水印.pdf')),
                                             'PDF 文件 (*.pdf)')
        if not path:
            return
        count = len(self.selected_ids())
        self._start(watermarks.export, (self.analysis, self.selected_ids(), path),
                    lambda target: self.lbl_status.setText(f'已去除 {count} 个所选对象并另存：' + target),
                    '正在处理、复检并另存…')

    def _cancel(self):
        if self.worker is not None:
            self.worker.requestInterruption()
            self.lbl_status.setText('正在取消，等待当前 PDF 操作结束…')

    def done(self, result):
        if self._can_close():
            self.preview_timer.stop()
            super().done(result)
