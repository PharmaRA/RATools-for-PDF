"""Preparation failures must terminate the batch and restore the security UI."""

import tempfile
import time
import unittest
from pathlib import Path
from unittest.mock import patch

import fitz
from PySide6.QtCore import QCoreApplication, QEvent, QThread
from PySide6.QtWidgets import QApplication, QTableWidgetItem

from ratools_pdf.ui.dialogs.security_center_dialog import DecryptWorker, SecurityCenterDialog


class SilentDecrypt(DecryptWorker):
    def run(self):
        pass


class DecryptFailureRecoveryTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.root = Path(self.directory.name)
        self.source = self.root/'input.pdf'
        self.source.write_bytes(b'synthetic source')

    def check_preparation_failure(self, error=None):
        output = self.root/'output'
        if error is None:
            output.write_bytes(b'existing ordinary file')
        worker = DecryptWorker([(0,str(self.source)),(1,str(self.root/'second.pdf'))], str(output))
        rows, summaries = [], []
        worker.file_finished.connect(lambda *args: rows.append(args))
        worker.finished_all.connect(lambda *args: summaries.append(args))
        with patch('ratools_pdf.ui.dialogs.security_center_dialog.qpdf.decrypt_pdf') as engine:
            if error:
                with patch('ratools_pdf.ui.dialogs.security_center_dialog.os.makedirs', side_effect=error):
                    worker.run()
            else:
                worker.run()
            engine.assert_not_called()
        self.assertEqual([row[:2] for row in rows], [(0,False),(1,False)])
        self.assertTrue(all(row[2] for row in rows))
        self.assertEqual(summaries, [(0,2,str(output))])
        self.assertEqual(self.source.read_bytes(), b'synthetic source')
        if error is None:
            self.assertEqual(output.read_bytes(), b'existing ordinary file')

    def test_permission_denied_reports_every_file_and_batch_completion(self):
        self.check_preparation_failure(PermissionError('synthetic denied directory'))

    def test_missing_device_reports_every_file_and_batch_completion(self):
        self.check_preparation_failure(FileNotFoundError('synthetic disconnected directory'))

    def test_existing_regular_file_is_preserved_and_reports_completion(self):
        self.check_preparation_failure()

    def dialog(self):
        dialog = SecurityCenterDialog()
        dialog.decrypt_table.setRowCount(1)
        for column in (3,4,5):
            dialog.decrypt_table.setItem(0,column,QTableWidgetItem(str(self.source) if column==5 else ''))
        dialog.txt_dec_out_dir.setText(str(self.root/'out'))
        self.addCleanup(self.dispose_dialog, dialog)
        return dialog

    def dispose_dialog(self, dialog):
        for worker in dialog.findChildren(QThread):
            worker.wait()
        dialog.close()
        dialog.deleteLater()
        QCoreApplication.sendPostedEvents(dialog,QEvent.DeferredDelete)

    def await_recovery(self, dialog):
        deadline = time.monotonic()+3
        while dialog.decrypt_worker is not None and time.monotonic()<deadline:
            self.app.processEvents()
            QThread.msleep(5)
        self.assertIsNone(dialog.decrypt_worker)
        self.assertTrue(dialog.btn_execute.isEnabled())
        self.assertTrue(dialog.btn_cancel.isEnabled())
        self.assertIn('失败', dialog.decrypt_table.item(0,3).text())
        self.assertTrue(dialog.btn_open_target.isHidden())

    def test_real_thread_preparation_failure_restores_ui_and_allows_retry(self):
        dialog = self.dialog()
        with patch('ratools_pdf.ui.dialogs.security_center_dialog.os.makedirs', side_effect=PermissionError('denied')), \
                patch('ratools_pdf.ui.dialogs.security_center_dialog.QMessageBox.warning'):
            dialog._execute_decryption()
            self.await_recovery(dialog)
            dialog._execute_decryption()
            self.await_recovery(dialog)

    def test_thread_return_without_summary_still_restores_ui(self):
        dialog = self.dialog()
        with patch('ratools_pdf.ui.dialogs.security_center_dialog.DecryptWorker', SilentDecrypt), \
                patch('ratools_pdf.ui.dialogs.security_center_dialog.QMessageBox.warning'):
            dialog._execute_decryption()
            self.await_recovery(dialog)

    def test_start_failure_reports_error_and_restores_ui(self):
        dialog = self.dialog()
        with patch.object(DecryptWorker,'start',side_effect=RuntimeError('cannot start')), \
                patch('ratools_pdf.ui.dialogs.security_center_dialog.QMessageBox.warning'):
            dialog._execute_decryption()
            self.await_recovery(dialog)

    def test_real_success_publishes_pdf_and_restores_output_controls(self):
        with fitz.open() as doc:
            doc.new_page().insert_text((72,72),'Preserved decrypted content')
            doc.save(self.source)
        dialog = self.dialog()
        with patch('ratools_pdf.ui.dialogs.security_center_dialog.QMessageBox.information') as message:
            dialog._execute_decryption()
            deadline = time.monotonic()+10
            while dialog.decrypt_worker is not None and time.monotonic()<deadline:
                self.app.processEvents()
                QThread.msleep(5)
            self.assertIsNone(dialog.decrypt_worker)
            self.assertTrue(dialog.btn_execute.isEnabled())
            self.assertFalse(dialog.btn_open_target.isHidden())
            self.assertIn('已解密',dialog.decrypt_table.item(0,3).text())
            message.assert_called_once()
        with fitz.open(self.root/'out'/'input_decrypted.pdf') as doc:
            self.assertIn('Preserved decrypted content',doc[0].get_text())


if __name__ == '__main__':
    unittest.main()
