"""Interactive, explicit selection of removable PDF objects."""

from pathlib import Path

from PySide6.QtCore import QSize, Qt, QThread, QTimer
from PySide6.QtGui import QColor, QPixmap
from PySide6.QtWidgets import (
    QCheckBox, QFileDialog, QFrame, QGridLayout, QHBoxLayout, QHeaderView,
    QLabel, QLineEdit, QPushButton, QScrollArea, QSizePolicy, QSpinBox,
    QSplitter, QTreeWidget, QTreeWidgetItem, QVBoxLayout, QWidget,
)

from ratools_pdf.pdf import watermarks
from ratools_pdf.ui import icons
from ratools_pdf.ui.dialogs.base import FramelessDraggableDialog
from ratools_pdf.ui.theme import active_palette
from ratools_pdf.ui.widgets import NoFocusItemDelegate


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
        self.resize(1040, 680)
        self.setMinimumSize(820, 560)
        self.content_layout.setSpacing(12)
        self.analysis = None
        self.worker = None
        self._preview_key = None
        self._viewed_pages = {}
        self._preview_requested = False
        self._preview_pixmaps = {}
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
        pal = active_palette()

        hint = QLabel(
            '选择可独立移除的批注、文字、图片或绘图对象。候选中可能包含正文；请查看前后效果后勾选。'
            '扫描图像内的水印和复杂混排内容暂不支持。'
        )
        hint.setObjectName('dialogInfoPanel')
        hint.setWordWrap(True)
        hint.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Minimum)
        self.content_layout.addWidget(hint)

        input_card = QFrame()
        input_card.setObjectName('wizardCard')
        grid = QGridLayout(input_card)
        grid.setContentsMargins(14, 12, 14, 12)
        grid.setSpacing(10)

        lbl_source = QLabel('目标文件')
        lbl_source.setObjectName('dialogCaption')
        self.txt_source = QLineEdit()
        self.txt_source.setObjectName('settingsPathEdit')
        self.txt_source.setFixedHeight(34)
        self.txt_source.setPlaceholderText('选择需要去除水印的 PDF 文件…')
        self.btn_source = QPushButton('选择 PDF…')
        self.btn_source.setObjectName('dialogSecondaryBtn')
        self.btn_source.setFixedHeight(34)
        self.btn_source.setMinimumWidth(112)
        self.btn_source.setCursor(Qt.PointingHandCursor)
        self.btn_source.setIconSize(QSize(14, 14))
        self.btn_source.setIcon(
            icons.svg_icon('folder_open', pal.text_body, size=14, disabled_color=pal.text_faint)
        )
        self.btn_source.clicked.connect(self._choose_source)
        grid.addWidget(lbl_source, 0, 0)
        grid.addWidget(self.txt_source, 0, 1)
        grid.addWidget(self.btn_source, 0, 2)

        lbl_pages = QLabel('处理页码')
        lbl_pages.setObjectName('dialogCaption')
        self.txt_pages = QLineEdit()
        self.txt_pages.setObjectName('settingsPathEdit')
        self.txt_pages.setFixedHeight(34)
        self.txt_pages.setPlaceholderText('例如 1-3,5；留空表示全部页面')
        self.btn_analyze = QPushButton('分析对象')
        self.btn_analyze.setObjectName('dialogPrimaryBtn')
        self.btn_analyze.setFixedHeight(34)
        self.btn_analyze.setMinimumWidth(112)
        self.btn_analyze.setCursor(Qt.PointingHandCursor)
        self.btn_analyze.setIconSize(QSize(14, 14))
        self.btn_analyze.setIcon(
            icons.svg_icon('search', pal.text_on_primary, size=14, disabled_color=pal.surface)
        )
        self.btn_analyze.clicked.connect(self._analyze)
        grid.addWidget(lbl_pages, 1, 0)
        grid.addWidget(self.txt_pages, 1, 1)
        grid.addWidget(self.btn_analyze, 1, 2)
        grid.setColumnStretch(1, 1)
        self.content_layout.addWidget(input_card)

        self.splitter = QSplitter(Qt.Horizontal)
        self.splitter.setObjectName('watermarkSplitter')

        left_card = QFrame()
        left_card.setObjectName('watermarkPanelCard')
        left_layout = QVBoxLayout(left_card)
        left_layout.setContentsMargins(1, 1, 1, 1)
        left_layout.setSpacing(0)

        left_header = QFrame()
        left_header.setObjectName('watermarkPanelHeader')
        left_header.setFixedHeight(40)
        left_header_layout = QHBoxLayout(left_header)
        left_header_layout.setContentsMargins(12, 0, 10, 0)
        left_header_layout.setSpacing(8)
        lbl_objects_title = QLabel('可选对象')
        lbl_objects_title.setObjectName('dialogSubTitle')
        self.lbl_objects_summary = QLabel('未分析')
        self.lbl_objects_summary.setObjectName('selectionSummary')
        left_header_layout.addWidget(lbl_objects_title)
        left_header_layout.addStretch()
        left_header_layout.addWidget(self.lbl_objects_summary)
        left_layout.addWidget(left_header)

        self.objects = QTreeWidget()
        self.objects.setHeaderLabels(['选择', '页码', '类型', '对象'])
        self.objects.setFrameShape(QFrame.NoFrame)
        self.objects.setAlternatingRowColors(True)
        self.objects.setUniformRowHeights(True)
        self.objects.setRootIsDecorated(False)
        self.objects.setIndentation(0)
        self.objects.setItemDelegate(NoFocusItemDelegate(self.objects))
        header = self.objects.header()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.Stretch)
        left_layout.addWidget(self.objects, 1)
        self.splitter.addWidget(left_card)

        right_card = QFrame()
        right_card.setObjectName('watermarkPanelCard')
        right_layout = QVBoxLayout(right_card)
        right_layout.setContentsMargins(1, 1, 1, 1)
        right_layout.setSpacing(0)

        right_header = QFrame()
        right_header.setObjectName('watermarkPanelHeader')
        right_header.setFixedHeight(40)
        controls = QHBoxLayout(right_header)
        controls.setContentsMargins(12, 0, 10, 0)
        controls.setSpacing(8)
        lbl_preview_title = QLabel('效果对照预览')
        lbl_preview_title.setObjectName('dialogSubTitle')
        controls.addWidget(lbl_preview_title)
        controls.addStretch()
        lbl_page_num = QLabel('预览页码')
        lbl_page_num.setObjectName('dialogMuted')
        controls.addWidget(lbl_page_num)
        self.page_number = QSpinBox()
        self.page_number.setMinimum(1)
        self.page_number.setFixedHeight(28)
        self.page_number.setMinimumWidth(72)
        controls.addWidget(self.page_number)
        self.btn_preview = QPushButton('刷新预览')
        self.btn_preview.setObjectName('dialogSecondaryBtn')
        self.btn_preview.setFixedHeight(28)
        self.btn_preview.setCursor(Qt.PointingHandCursor)
        self.btn_preview.setIconSize(QSize(14, 14))
        self.btn_preview.setIcon(
            icons.svg_icon('refresh', pal.text_body, size=14, disabled_color=pal.text_faint)
        )
        self.btn_preview.clicked.connect(self._queue_preview)
        controls.addWidget(self.btn_preview)
        right_layout.addWidget(right_header)

        previews_body = QWidget()
        previews_layout = QHBoxLayout(previews_body)
        previews_layout.setContentsMargins(12, 12, 12, 12)
        previews_layout.setSpacing(12)
        self.before_scroll, self.before = self._image_panel(previews_layout, '原始页面')
        self.after_scroll, self.after = self._image_panel(
            previews_layout, '去除勾选对象后的页面', accent='after'
        )
        right_layout.addWidget(previews_body, 1)
        self.splitter.addWidget(right_card)

        self.splitter.setChildrenCollapsible(False)
        self.splitter.setSizes([340, 660])
        self.splitter.setStretchFactor(0, 1)
        self.splitter.setStretchFactor(1, 2)
        self.splitter.splitterMoved.connect(self._rescale_previews)
        self.content_layout.addWidget(self.splitter, 1)

        self.lbl_notes = QLabel()
        self.lbl_notes.setObjectName('watermarkNotesPanel')
        self.lbl_notes.setWordWrap(True)
        self.lbl_notes.setVisible(False)
        self.content_layout.addWidget(self.lbl_notes)

        self.confirm_card = QFrame()
        self.confirm_card.setObjectName('watermarkConfirmCard')
        self.confirm_card.setProperty('checked', 'false')
        confirm_layout = QHBoxLayout(self.confirm_card)
        confirm_layout.setContentsMargins(12, 10, 12, 10)
        confirm_layout.setSpacing(8)
        self.confirm = QCheckBox('已检查所选对象的预览，确认去除；修改后的副本不保留原数字签名的有效性。')
        self.confirm.setCursor(Qt.PointingHandCursor)
        self.confirm.toggled.connect(self._sync_confirm_card)
        self.confirm.toggled.connect(self._update_actions)
        confirm_layout.addWidget(self.confirm)
        self.content_layout.addWidget(self.confirm_card)

        footer = QHBoxLayout()
        footer.setContentsMargins(0, 2, 0, 0)
        footer.setSpacing(10)
        self.lbl_status = QLabel()
        self.lbl_status.setObjectName('dialogStatus')
        self.lbl_status.setWordWrap(True)
        self._set_status('选择文件与页码后，点击“分析对象”。', 'muted')
        footer.addWidget(self.lbl_status, 1)

        self.btn_cancel = QPushButton('取消当前操作')
        self.btn_cancel.setObjectName('dialogSecondaryBtn')
        self.btn_cancel.setFixedHeight(34)
        self.btn_cancel.setCursor(Qt.PointingHandCursor)
        self.btn_cancel.clicked.connect(self._cancel)

        self.btn_close_dialog = QPushButton('关闭')
        self.btn_close_dialog.setObjectName('dialogSecondaryBtn')
        self.btn_close_dialog.setFixedHeight(34)
        self.btn_close_dialog.setCursor(Qt.PointingHandCursor)
        self.btn_close_dialog.clicked.connect(self.reject)

        self.btn_save = QPushButton('另存为 PDF…')
        self.btn_save.setObjectName('dialogPrimaryBtn')
        self.btn_save.setFixedHeight(34)
        self.btn_save.setCursor(Qt.PointingHandCursor)
        self.btn_save.setIconSize(QSize(14, 14))
        self.btn_save.setIcon(
            icons.svg_icon('download', pal.text_on_primary, size=14, disabled_color=pal.surface)
        )
        self.btn_save.clicked.connect(self._save)

        footer.addWidget(self.btn_cancel)
        footer.addWidget(self.btn_close_dialog)
        footer.addWidget(self.btn_save)
        self.content_layout.addLayout(footer)
        self._update_actions()

    @staticmethod
    def _image_panel(layout, title, accent=None):
        pane = QWidget()
        pane.setObjectName('watermarkPreviewPane')
        panel = QVBoxLayout(pane)
        panel.setContentsMargins(0, 0, 0, 0)
        panel.setSpacing(6)

        title_lbl = QLabel(title)
        title_lbl.setObjectName('watermarkPreviewTitle')
        if accent:
            title_lbl.setProperty('accent', accent)
        panel.addWidget(title_lbl)

        area = QScrollArea()
        area.setObjectName('watermarkPreviewScroll')
        area.setFrameShape(QFrame.NoFrame)
        area.setWidgetResizable(True)
        label = QLabel('分析后显示预览')
        label.setObjectName('watermarkPreviewCanvas')
        label.setAlignment(Qt.AlignCenter)
        label.setWordWrap(True)
        area.setWidget(label)
        panel.addWidget(area, 1)
        layout.addWidget(pane, 1)
        return area, label

    def _set_status(self, text, state='muted'):
        self.lbl_status.setText(text)
        self.lbl_status.setProperty('state', state)
        self.lbl_status.style().unpolish(self.lbl_status)
        self.lbl_status.style().polish(self.lbl_status)

    def _sync_confirm_card(self, checked):
        self.confirm_card.setProperty('checked', 'true' if checked else 'false')
        self.confirm_card.style().unpolish(self.confirm_card)
        self.confirm_card.style().polish(self.confirm_card)

    def _update_objects_summary(self):
        if self.analysis is None:
            self.lbl_objects_summary.setText('未分析')
            return
        total = len(self.analysis.candidates)
        selected = len(self.selected_ids())
        self.lbl_objects_summary.setText(f'共 {total} 项 · 已选 {selected} 项')

    @staticmethod
    def _apply_preview_pixmap(label, scroll_area, pixmap):
        viewport_width = scroll_area.viewport().width()
        target_width = max(220, viewport_width - 24) if viewport_width >= 100 else 320
        label.setPixmap(pixmap.scaledToWidth(target_width, Qt.SmoothTransformation))

    def _rescale_previews(self, *_args):
        for key, label, scroll_area in (
            ('before', self.before, self.before_scroll),
            ('after', self.after, self.after_scroll),
        ):
            pixmap = self._preview_pixmaps.get(key)
            if pixmap is not None and not pixmap.isNull():
                self._apply_preview_pixmap(label, scroll_area, pixmap)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._rescale_previews()

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
        self._preview_pixmaps.clear()
        self.objects.clear()
        self.confirm.setChecked(False)
        self.before.setPixmap(QPixmap())
        self.after.setPixmap(QPixmap())
        self.before.setText('请重新分析')
        self.after.setText('请重新分析')
        self.lbl_notes.clear()
        self.lbl_notes.setVisible(False)
        self._update_objects_summary()
        self._set_status('文件或页码已改变，请重新分析。', 'muted')
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
                       self.objects, self.page_number, self.confirm, self.btn_close_dialog):
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
        self._set_status(message, 'info')
        task = WatermarkTask(action, args, self)
        self.worker = task
        generation = self._generation

        def completed():
            self.worker = None
            if generation == self._generation:
                if task.error:
                    self._set_status(task.error, 'error')
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
        pal = active_palette()
        self.objects.blockSignals(True)
        for candidate in analysis.candidates:
            kind_text = f'{candidate.kind} · 疑似' if candidate.watermark_likely else candidate.kind
            item = QTreeWidgetItem(['', str(candidate.page + 1), kind_text, candidate.label])
            item.setData(0, Qt.UserRole, candidate.id)
            item.setData(1, Qt.UserRole, candidate.page)
            item.setData(2, Qt.UserRole, candidate)
            item.setTextAlignment(1, Qt.AlignCenter)
            if candidate.watermark_likely:
                item.setForeground(2, QColor(pal.primary_text))
            tip = f'{candidate.label}\n{candidate.detail}' if candidate.detail else candidate.label
            item.setToolTip(3, tip)
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
        self.lbl_notes.setVisible(bool(note_text))
        self._update_objects_summary()
        self._set_status(f'找到 {len(analysis.candidates)} 个可选对象，均未勾选。勾选后查看右侧变化。', 'muted')
        self._queue_preview()

    def _selection_changed(self, item, _column):
        self.confirm.setChecked(False)
        self._preview_key = None
        self._update_objects_summary()
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
            self._preview_pixmaps.clear()
            self.before.setPixmap(QPixmap())
            self.after.setPixmap(QPixmap())
            self.before.setText('该页未包含在处理页码中')
            self.after.setText('请切换至处理范围内的页面')
            return

        def rendered(images):
            for key, label, scroll_area, data in zip(
                ('before', 'after'),
                (self.before, self.after),
                (self.before_scroll, self.after_scroll),
                images,
            ):
                image = QPixmap()
                image.loadFromData(data, 'PNG')
                self._preview_pixmaps[key] = image
                self._apply_preview_pixmap(label, scroll_area, image)
            self._preview_key = selected
            self._viewed_pages[number] = self._page_selections().get(number, ())
            pending = self._pending_pages()
            remaining = (' 请继续预览第 ' + ', '.join(str(page + 1) for page in pending) + ' 页。') if pending else ''
            self._set_status(
                f'第 {number + 1} 页预览已更新；共选择 {len(selected)} 个对象。' + remaining,
                'info' if selected else 'muted',
            )

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
        self._start(
            watermarks.export,
            (self.analysis, self.selected_ids(), path),
            lambda target: self._set_status(f'已去除 {count} 个所选对象并另存：' + target, 'success'),
            '正在处理、复检并另存…',
        )

    def _cancel(self):
        if self.worker is not None:
            self.worker.requestInterruption()
            self._set_status('正在取消，等待当前 PDF 操作结束…', 'info')

    def done(self, result):
        if self._can_close():
            self.preview_timer.stop()
            super().done(result)
