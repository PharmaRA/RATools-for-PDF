"""Read-only, cancellable batch preview; the dialog owns its reader thread."""

from copy import deepcopy

from PySide6.QtCore import QThread, QTimer, Qt, Signal
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QFrame, QHBoxLayout, QHeaderView, QLabel, QPushButton, QSizePolicy, QTreeWidget, QTreeWidgetItem,
)

from ratools_pdf.config.rules_catalog import option_title
from ratools_pdf.services.processing_plan import build_processing_plan
from ratools_pdf.ui.dialogs.base import FramelessDraggableDialog
from ratools_pdf.ui.theme import active_palette


class PlanReader(QThread):
    ready = Signal(object)
    progress = Signal(int, int)

    def __init__(self, files, options, mode, settings, base, parent=None, profile=None):
        super().__init__(parent)
        self.arguments = (list(files), set(options), mode, dict(settings) if settings else None, base)
        self.profile = deepcopy(profile)

    def run(self):
        self.ready.emit(build_processing_plan(
            *self.arguments, cancelled=self.isInterruptionRequested,
            progress=lambda done, total, _item: self.progress.emit(done, total), profile=self.profile))


class ProcessingPlanDialog(FramelessDraggableDialog):
    def __init__(self, files, options, mode='smart', settings=None, common_base='', parent=None, profile=None):
        super().__init__('处理计划预览', parent)
        self.resize(900, 580)
        self.content_layout.setSpacing(12)
        self.plan = build_processing_plan([], options, mode, settings, common_base, profile=profile)
        self._cancel_requested = False

        self.lbl_status = QLabel('正在只读检查输入与规则，请稍候；可随时取消。')
        self.lbl_status.setWordWrap(True)
        self.lbl_status.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Minimum)
        self.lbl_status.setObjectName('dialogInfoPanel')
        self.content_layout.addWidget(self.lbl_status)

        tree_frame = QFrame()
        tree_frame.setObjectName('logSummaryFrame')
        tree_box = QHBoxLayout(tree_frame)
        tree_box.setContentsMargins(1, 1, 1, 1)
        tree_box.setSpacing(0)

        self.preview_tree = QTreeWidget()
        self.preview_tree.setColumnCount(4)
        self.preview_tree.setHeaderLabels(['文件 / 已选规则', '预检', '计划', '原因 / 风险'])
        self.preview_tree.setAlternatingRowColors(True)
        self.preview_tree.setUniformRowHeights(True)
        self.preview_tree.setIndentation(16)
        header = self.preview_tree.header()
        header.setSectionResizeMode(0, QHeaderView.Stretch)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.Stretch)
        tree_box.addWidget(self.preview_tree)
        self.content_layout.addWidget(tree_frame, 1)

        notice = QLabel('签名有效性、规则实际 changed 与规则级验证：not_verified。结构对照与保存重开检查分别记录。')
        notice.setWordWrap(True)
        notice.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Minimum)
        notice.setObjectName('dialogMuted')
        self.content_layout.addWidget(notice)

        buttons = QHBoxLayout()
        buttons.setSpacing(10)
        buttons.addStretch()
        self.btn_cancel = QPushButton('取消')
        self.btn_cancel.setObjectName('dialogSecondaryBtn')
        self.btn_cancel.setFixedHeight(34)
        self.btn_cancel.setCursor(Qt.PointingHandCursor)
        self.btn_confirm = QPushButton('确认计划并开始')
        self.btn_confirm.setObjectName('dialogPrimaryBtn')
        self.btn_confirm.setFixedHeight(34)
        self.btn_confirm.setCursor(Qt.PointingHandCursor)
        self.btn_confirm.setEnabled(False)
        self.btn_cancel.clicked.connect(self.reject)
        self.btn_confirm.clicked.connect(self.accept)
        buttons.addWidget(self.btn_cancel)
        buttons.addWidget(self.btn_confirm)
        self.content_layout.addLayout(buttons)

        self.worker = PlanReader(files, options, mode, settings, common_base, self, profile=profile)
        self.worker.ready.connect(self._receive_plan)
        self.worker.progress.connect(lambda done, total: self.lbl_status.setText(f'正在生成只读计划：{done} / {total}'))
        self.worker.finished.connect(self._reader_finished)
        QTimer.singleShot(0, self._start_reader)

    def _start_reader(self):
        if not self._cancel_requested:
            self.worker.start()

    def _receive_plan(self, plan):
        plan['batch_id'] = self.plan['batch_id']
        self.plan = plan

    def _reader_finished(self):
        self.worker.wait()
        if self._cancel_requested:
            super().reject()
            return
        pal = active_palette()
        detection = {'hit': '命中', 'miss': '未命中', 'unsupported': '不支持预检', 'unknown': '未知'}
        reasons = {'selected_by_user': '用户选择后执行', 'force_mode': '强制执行用户选择',
                   'precheck_hit': '预检命中', 'precheck_miss': '未发现需处理的问题',
                   'inspection_failed': '输入检查失败，跳过', 'external_workflow_required': '需要外部字体工作流'}
        risks = {'signature_present': '含签名，重写可能破坏有效性', 'forms_present': '含表单，请复核',
                 'attachments_present': '含附件，请复核', 'geometry_may_change': '页面尺寸或旋转可能变化'}
        ready = 0
        for item in self.plan['files']:
            good = item['status'] == 'ready'
            ready += int(good)
            warning = '；'.join(risks.get(value, value) for value in item['risks'])
            if item['errors']:
                warning = '检查失败：' + '、'.join(item['errors'])
            node = QTreeWidgetItem([item['relative_path'], '已检查' if good else '失败',
                                   '可执行' if good else '跳过', warning])
            node.setForeground(1, QColor(pal.success_text if good else pal.danger_text))
            node.setForeground(2, QColor(pal.primary_text if good else pal.text_muted))
            if warning:
                node.setForeground(3, QColor(pal.danger_text if item['errors'] else pal.warning_text))
            self.preview_tree.addTopLevelItem(node)
            for rule in item['rules']:
                execute = rule['decision'] == 'execute'
                child = QTreeWidgetItem([
                    option_title(rule['option_id']), detection[rule['precheck']],
                    '执行' if execute else '跳过', reasons[rule['reason']]])
                child.setForeground(2, QColor(pal.primary_text if execute else pal.text_muted))
                child.setForeground(3, QColor(pal.text_muted))
                node.addChild(child)
            node.setExpanded(True)
        self.preview_tree.resizeColumnToContents(0)
        self.lbl_status.setText(f'计划已就绪：{ready} / {len(self.plan["files"])} 个文件可执行。检查失败的文件保留失败事实。')
        self.btn_confirm.setEnabled(ready > 0 and not self.plan['cancelled'])

    def reject(self):
        self._cancel_requested = True
        self.btn_confirm.setEnabled(False)
        if self.worker.isRunning():
            self.worker.requestInterruption()
            self.lbl_status.setText('正在取消只读检查并等待线程安全退出…')
            return
        super().reject()

    def closeEvent(self, event):
        if self.worker.isRunning():
            self.reject()
            event.ignore()
            return
        super().closeEvent(event)
