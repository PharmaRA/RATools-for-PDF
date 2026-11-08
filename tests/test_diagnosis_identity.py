"""Only the latest diagnosis for the displayed file may publish an AST."""

import tempfile
import threading
import time
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

from PySide6.QtCore import QCoreApplication, QEvent, QThread, Signal
from PySide6.QtWidgets import QApplication

from ratools_pdf.ui.dialogs.pdf_inspector_dialog import PdfInspectorDialog


class ManualDiagnostic(QThread):
    finished_signal = Signal(dict)

    def __init__(self, pdf_path, parent=None):
        super().__init__(parent)
        self.pdf_path = pdf_path

    def start(self):
        pass

    def publish(self, label, error=None, path=None):
        data = dict(path=path or self.pdf_path, check={'output':label, 'is_healthy':True},
                    encryption='file is not encrypted', linearization='not linearized',
                    json_ast={'document':label})
        if error:
            data['error'] = error
        self.finished_signal.emit(data)


class HeldDiagnostic(ManualDiagnostic):
    def __init__(self, pdf_path, parent=None):
        super().__init__(pdf_path, parent)
        self.release = threading.Event()

    def start(self):
        QThread.start(self)

    def run(self):
        self.release.wait(5)
        self.publish(Path(self.pdf_path).stem)


class DiagnosisIdentityTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.a, self.b = (Path(self.directory.name)/name for name in ('a.pdf','b.pdf'))
        self.a.write_bytes(b'synthetic A')
        self.b.write_bytes(b'synthetic B')
        self.dialog = PdfInspectorDialog()
        self.addCleanup(self.cleanup_dialog)
        self.workers = []
        def create(path, parent=None):
            worker = ManualDiagnostic(path, parent)
            self.workers.append(worker)
            return worker
        self.factory = patch('ratools_pdf.ui.dialogs.pdf_inspector_dialog.DiagnosticWorker', side_effect=create)
        self.factory.start()
        self.addCleanup(self.factory.stop)

    def cleanup_dialog(self):
        self.dialog.close()
        self.dialog.deleteLater()
        QCoreApplication.sendPostedEvents(self.dialog, QEvent.DeferredDelete)

    def start(self, path):
        self.dialog.txt_path.setText(str(path))
        self.dialog.run_diagnosis(str(path))
        return self.workers[-1]

    def test_late_previous_file_cannot_replace_latest_display_or_export_ast(self):
        first, latest = self.start(self.a), self.start(self.b)
        latest.publish('B')
        first.publish('A')
        self.assertEqual(self.dialog.txt_check_log.toPlainText(), 'B')
        self.assertEqual(self.dialog.current_json_ast, {'document':'B'})
        self.assertEqual(self.dialog.txt_path.text(), str(self.b))

    def test_same_path_revisited_uses_generation_not_just_path(self):
        first, second, latest = self.start(self.a), self.start(self.b), self.start(self.a)
        latest.publish('latest A')
        first.publish('old A')
        second.publish('old B')
        self.assertEqual(self.dialog.current_json_ast, {'document':'latest A'})
        self.assertEqual(self.dialog.txt_check_log.toPlainText(), 'latest A')

    def test_path_edit_clears_old_ast_and_rejects_inflight_result(self):
        first = self.start(self.a)
        first.publish('A')
        self.dialog.txt_path.setText(str(self.b))
        self.assertIsNone(self.dialog.current_json_ast)
        self.assertFalse(self.dialog.btn_export_json.isEnabled())
        self.assertEqual(self.dialog.txt_check_log.toPlainText(), '')
        first.publish('late A')
        self.assertIsNone(self.dialog.current_json_ast)

    def test_wrong_payload_path_cannot_publish(self):
        worker = self.start(self.a)
        worker.publish('wrong document', path=str(self.b))
        self.assertIsNone(self.dialog.current_json_ast)
        self.assertFalse(self.dialog.btn_export_json.isEnabled())

    def test_failure_clears_partial_ast_and_retry_can_succeed(self):
        failed = self.start(self.a)
        failed.publish('partial', error='synthetic failure')
        self.assertIsNone(self.dialog.current_json_ast)
        self.assertFalse(self.dialog.btn_export_json.isEnabled())
        self.assertIn('失败', self.dialog.lbl_status.text())
        failed.finished.emit()
        self.assertTrue(self.dialog.btn_diagnose.isEnabled())
        retry = self.start(self.a)
        retry.publish('retried A')
        self.assertEqual(self.dialog.current_json_ast, {'document':'retried A'})

    def test_each_finished_worker_is_released_without_clearing_newer_task(self):
        first, latest = self.start(self.a), self.start(self.b)
        first_deleted, latest_deleted = Mock(), Mock()
        first.destroyed.connect(first_deleted)
        latest.destroyed.connect(latest_deleted)
        first.finished.emit()
        self.assertIs(self.dialog.worker, latest)
        self.assertFalse(self.dialog.btn_diagnose.isEnabled())
        latest.publish('B')
        latest.finished.emit()
        self.assertIsNone(self.dialog.worker)
        QCoreApplication.sendPostedEvents(first, QEvent.DeferredDelete)
        QCoreApplication.sendPostedEvents(latest, QEvent.DeferredDelete)
        first_deleted.assert_called_once()
        latest_deleted.assert_called_once()

    def test_start_failure_restores_retry_and_keeps_export_disabled(self):
        with patch.object(ManualDiagnostic, 'start', side_effect=RuntimeError('cannot start')):
            self.start(self.a)
        self.assertIsNone(self.dialog.worker)
        self.assertTrue(self.dialog.btn_diagnose.isEnabled())
        self.assertFalse(self.dialog.btn_export_json.isEnabled())
        self.assertIn('失败', self.dialog.lbl_status.text())

    def test_path_change_during_export_picker_cannot_write_stale_ast(self):
        self.start(self.a).publish('A')
        output = Path(self.directory.name)/'export.json'
        def choose(*_):
            self.dialog.txt_path.setText(str(self.b))
            return str(output), ''
        with patch('ratools_pdf.ui.dialogs.pdf_inspector_dialog.QFileDialog.getSaveFileName', side_effect=choose), \
                patch('ratools_pdf.ui.dialogs.pdf_inspector_dialog.QMessageBox.information'):
            self.dialog._on_export_json_clicked()
        self.assertFalse(output.exists())

    def test_real_threads_complete_out_of_order_and_both_release(self):
        workers, destroyed = [], []
        def create(path, parent=None):
            worker = HeldDiagnostic(path, parent)
            worker.destroyed.connect(lambda: destroyed.append(path))
            workers.append(worker)
            return worker
        def spin_until(condition):
            deadline = time.monotonic() + 3
            while not condition() and time.monotonic() < deadline:
                self.app.processEvents()
                for worker in workers:
                    try:
                        QCoreApplication.sendPostedEvents(worker, QEvent.DeferredDelete)
                    except RuntimeError:
                        pass  # This worker was already destroyed on an earlier iteration.
                QThread.msleep(5)
            self.assertTrue(condition())
        try:
            with patch('ratools_pdf.ui.dialogs.pdf_inspector_dialog.DiagnosticWorker', side_effect=create):
                self.dialog.run_diagnosis(str(self.a))
                self.dialog.run_diagnosis(str(self.b))
            workers[1].release.set()
            spin_until(lambda: str(self.b) in destroyed)
            self.assertEqual(self.dialog.current_json_ast, {'document':'b'})
            self.assertTrue(workers[0].isRunning())
            workers[0].release.set()
            spin_until(lambda: str(self.a) in destroyed)
            self.assertIsNone(self.dialog.worker)
            self.assertEqual(self.dialog.current_json_ast, {'document':'b'})
            self.assertTrue(self.dialog._can_close())
        finally:
            for worker in workers:
                worker.release.set()
                try:
                    worker.wait(2000)
                except RuntimeError:
                    pass


if __name__ == '__main__':
    unittest.main()
