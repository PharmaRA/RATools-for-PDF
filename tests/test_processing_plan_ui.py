"""PD02 user entry and structured batch evidence, without real settings storage."""

import tempfile
import time
import unittest
from pathlib import Path
from types import MethodType, SimpleNamespace
from unittest.mock import Mock, patch

from PySide6.QtCore import QObject, QThread
from PySide6.QtWidgets import QApplication, QDialog, QPushButton

from ratools_pdf.controllers.main_controller import MainController
from ratools_pdf.services.processing_plan import build_processing_plan
from test_processing_verification import write_pdf

try:
    from ratools_pdf.ui.dialogs.processing_plan import ProcessingPlanDialog
except ImportError:
    ProcessingPlanDialog = None


class ProcessingPlanUITests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.root = Path(self.directory.name)
        self.source = self.root/'source.pdf'
        write_pdf(self.source)
        self.plan = build_processing_plan([str(self.source)], {'compress_images'}, common_base=str(self.root))
        self.view = SimpleNamespace(
            all_checkboxes={}, get_selected_options=lambda: {'compress_images'},
            get_processing_mode=lambda: 'smart', get_compression_settings=lambda: None,
            settings_dialog=SimpleNamespace(default_output_edit=SimpleNamespace(text=lambda: '')),
            set_task_busy=Mock(), show_warning_message=Mock(), style=self.app.style)
        self.view.btn_start = QPushButton()
        self.host = QObject()
        self.host.__dict__.update(
            worker=None, loaded_files=[str(self.source)], view=self.view,
            _is_precheck_running=lambda: False, _is_detection_running=lambda: False,
            io=SimpleNamespace(is_running=lambda: False),
            _prompt_skip_signed_files=lambda files: files, _get_processing_worker_count=lambda: 1)
        self.host.ensure_idle = MethodType(MainController.ensure_idle, self.host)
        if hasattr(MainController, '_preview_processing_plan'):
            self.host._preview_processing_plan = MethodType(MainController._preview_processing_plan, self.host)

    def spin_until(self, condition):
        end = time.monotonic() + 4
        while not condition() and time.monotonic() < end:
            self.app.processEvents()
            QThread.msleep(5)
        self.assertTrue(condition())

    def require_dialog(self):
        self.assertIsNotNone(ProcessingPlanDialog, 'PD02 preview dialog is missing')

    def test_preview_planning_blocks_other_main_controller_actions(self):
        self.host.planning_in_progress = True
        self.assertFalse(self.host.ensure_idle('导出'))

    def test_cancel_preview_keeps_pdf_unchanged_and_does_not_create_output(self):
        self.require_dialog()
        original = self.source.read_bytes()
        dialog = Mock()
        dialog.plan = self.plan
        dialog.exec.return_value = QDialog.Rejected
        with patch('ratools_pdf.controllers.main_controller.ProcessingPlanDialog', return_value=dialog), \
                patch('ratools_pdf.controllers.main_controller.QFileDialog.getExistingDirectory', return_value=str(self.root)), \
                patch('ratools_pdf.controllers.main_controller.ProcessWorker'):
            MainController.start_processing(self.host)
        self.assertIsNone(self.host.worker)
        self.assertEqual(self.source.read_bytes(), original)
        self.assertFalse((self.root/'RATools_Output').exists())
        self.assertEqual(self.host.processing_evidence['outcome'], 'cancelled')
        self.assertTrue(self.host.processing_evidence['ended_at'])
        self.assertFalse(self.host.planning_in_progress)

    def test_modal_reentry_is_rechecked_before_batch_ownership(self):
        self.require_dialog()
        dialog = Mock()
        dialog.plan = self.plan
        def accept_after_other_task():
            self.host.io.is_running = lambda: True
            return QDialog.Accepted
        dialog.exec.side_effect = accept_after_other_task
        with patch('ratools_pdf.controllers.main_controller.ProcessingPlanDialog', return_value=dialog), \
                patch('ratools_pdf.controllers.main_controller.QFileDialog.getExistingDirectory', return_value=str(self.root)), \
                patch('ratools_pdf.controllers.main_controller.ProcessWorker'):
            MainController.start_processing(self.host)
        self.assertIsNone(self.host.worker)
        self.assertFalse((self.root/'RATools_Output').exists())
        self.assertEqual(self.host.processing_evidence['outcome'], 'not_started')

    def test_configuration_changed_during_preview_refuses_execution(self):
        self.require_dialog()
        dialog = Mock()
        dialog.plan = self.plan
        def change_options():
            self.view.get_selected_options = lambda: {'cleanup_remove_metadata'}
            return QDialog.Accepted
        dialog.exec.side_effect = change_options
        with patch('ratools_pdf.controllers.main_controller.ProcessingPlanDialog', return_value=dialog), \
                patch('ratools_pdf.controllers.main_controller.QFileDialog.getExistingDirectory', return_value=str(self.root)), \
                patch('ratools_pdf.controllers.main_controller.ProcessWorker'):
            MainController.start_processing(self.host)
        self.assertIsNone(self.host.worker)
        self.assertEqual(self.host.processing_evidence['outcome'], 'stale_plan')

    def test_real_preview_loads_rules_and_cancel_finishes_read_thread(self):
        self.require_dialog()
        dialog = ProcessingPlanDialog([str(self.source)], {'compress_images'}, 'smart', None, str(self.root))
        self.addCleanup(dialog.close)
        original = self.source.read_bytes()
        dialog.show()
        self.spin_until(lambda: dialog.btn_confirm.isEnabled())
        self.assertEqual(dialog.preview_tree.topLevelItemCount(), 1)
        self.assertEqual(dialog.preview_tree.topLevelItem(0).childCount(), 1)
        self.assertEqual(dialog.plan['files'][0]['rules'][0]['reason'], 'selected_by_user')
        dialog.reject()
        self.assertTrue(dialog.worker.wait(0))
        self.assertEqual(dialog.result(), QDialog.Rejected)
        self.assertEqual(self.source.read_bytes(), original)

    def test_cancel_while_planning_does_not_destroy_a_running_qthread(self):
        self.require_dialog()
        dialog = ProcessingPlanDialog([str(self.source)] * 30, {'compress_images'}, 'smart', None, str(self.root))
        self.addCleanup(dialog.close)
        dialog.show()
        self.app.processEvents()
        dialog.reject()
        self.spin_until(lambda: dialog.worker.wait(0) and not dialog.isVisible())
        self.assertEqual(dialog.result(), QDialog.Rejected)
        self.assertFalse(dialog.btn_confirm.isEnabled())

    def test_events_remain_in_own_batch_and_do_not_parse_localized_messages(self):
        self.assertTrue(hasattr(MainController, '_retain_processing_plan'), 'PD02 batch evidence retention is missing')
        MainController._retain_processing_plan(self.host, self.plan, 'running')
        first = self.host.processing_evidence
        second = build_processing_plan([str(self.source)], {'compress_images'}, common_base=str(self.root))
        MainController._retain_processing_plan(self.host, second, 'running')
        self.host.process_log_starts = {}
        self.host.process_log_rows = []
        file_id = self.plan['files'][0]['file_id']
        event = {'type': 'terminal', 'batch_id': self.plan['batch_id'], 'file_id': file_id,
                 'file_path': str(self.source), 'status': '处理完成', 'time': '12:00:00',
                 'message': 'private arbitrary localized text', 'outcome': 'verification_failed',
                 'engine_success': True, 'published': False,
                 'before': self.plan['files'][0]['before'], 'after': None,
                 'verification': {'save_reopen': 'failed'}, 'ended_at': '2026-11-09T04:00:00+00:00'}
        MainController.record_process_log_event(self.host, event)
        self.assertEqual(first['files'][file_id]['outcome'], 'verification_failed')
        self.assertFalse(first['files'][file_id]['published'])
        self.assertEqual(len(first['events']), 1)
        self.assertNotIn('message', first['events'][0])
        self.assertEqual(self.host.processing_evidence['events'], [])
        self.assertEqual(len(self.host.processing_evidence_history), 2)

    def test_finished_batch_retains_snapshot_after_ui_state_reset(self):
        self.assertTrue(hasattr(MainController, '_retain_processing_plan'), 'PD02 batch evidence retention is missing')
        MainController._retain_processing_plan(self.host, self.plan, 'running')
        MainController._finish_processing_evidence(self.host, 'finished')
        self.assertTrue(self.host.processing_evidence['ended_at'])
        self.assertEqual(self.host.processing_evidence['plan']['files'][0]['before']['page_count'], 1)
        self.assertEqual(self.host.processing_evidence['files'][self.plan['files'][0]['file_id']]['outcome'], 'not_started')


if __name__ == '__main__':
    unittest.main()
