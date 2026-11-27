import tempfile
import time
from pathlib import Path
import unittest
from unittest.mock import patch

import fitz
from PySide6.QtCore import QCoreApplication, QEvent, Qt
from PySide6.QtWidgets import QApplication
from shiboken6 import isValid

from ratools_pdf.ui.dialogs.watermark_dialog import WatermarkDialog


class WatermarkDialogTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.source = Path(self.temp.name) / 'source.pdf'
        self.output = Path(self.temp.name) / 'output.pdf'
        with fitz.open() as doc:
            doc.new_page().insert_text((30, 80), 'DRAFT')
            doc.save(self.source)
        self.dialog = WatermarkDialog(str(self.source))
        self.dialog.show()

    def tearDown(self):
        self.wait_idle()
        self.dialog.close()
        self.dialog.deleteLater()
        # Do not drain deferred destruction belonging to earlier test modules.
        QCoreApplication.sendPostedEvents(self.dialog, QEvent.DeferredDelete)
        self.assertFalse(isValid(self.dialog))
        self.temp.cleanup()

    def wait_idle(self):
        deadline = time.monotonic() + 10
        while self.dialog.worker is not None or self.dialog.preview_timer.isActive():
            self.app.processEvents()
            if time.monotonic() > deadline:
                self.fail('watermark task did not finish')
            time.sleep(.005)
        self.app.processEvents()

    def test_analyze_select_preview_confirm_and_save(self):
        self.dialog.btn_analyze.click()
        self.wait_idle()
        self.assertEqual(self.dialog.objects.topLevelItemCount(), 1)
        item = self.dialog.objects.topLevelItem(0)
        self.assertEqual(item.checkState(0), Qt.Unchecked)
        self.assertFalse(self.dialog.btn_save.isEnabled())
        item.setCheckState(0, Qt.Checked)
        self.wait_idle()
        self.assertNotEqual(self.dialog.before.pixmap().toImage(), self.dialog.after.pixmap().toImage())
        self.dialog.confirm.setChecked(True)
        self.assertTrue(self.dialog.btn_save.isEnabled())
        with patch('ratools_pdf.ui.dialogs.watermark_dialog.QFileDialog.getSaveFileName',
                   return_value=(str(self.output), 'PDF')):
            self.dialog.btn_save.click()
        self.wait_idle()
        self.assertTrue(self.output.exists(), self.dialog.lbl_status.text())
        with fitz.open(self.output) as result:
            self.assertNotIn('DRAFT', result[0].get_text())

    def test_range_change_invalidates_selection_preview_and_confirmation(self):
        self.dialog.btn_analyze.click()
        self.wait_idle()
        self.dialog.objects.topLevelItem(0).setCheckState(0, Qt.Checked)
        self.wait_idle()
        self.dialog.confirm.setChecked(True)
        self.dialog.txt_pages.setText('1')
        self.assertIsNone(self.dialog.analysis)
        self.assertFalse(self.dialog.confirm.isChecked())
        self.assertFalse(self.dialog.btn_save.isEnabled())
        self.assertEqual(self.dialog.objects.topLevelItemCount(), 0)

    def test_close_waits_for_running_task(self):
        import threading
        started, release = threading.Event(), threading.Event()

        def action(cancel=None):
            started.set()
            release.wait(5)

        self.dialog._start(action, (), lambda _: None, 'busy')
        try:
            self.assertTrue(started.wait(2))
            self.dialog.reject()
            self.assertTrue(self.dialog.isVisible())
            self.assertFalse(self.dialog.btn_save.isEnabled())
        finally:
            release.set()
        self.wait_idle()
        self.dialog.reject()
        self.assertFalse(self.dialog.isVisible())

    def test_each_selected_page_requires_preview_of_current_selection(self):
        with fitz.open() as doc:
            doc.new_page().insert_text((30, 80), 'DRAFT ONE')
            doc.new_page().insert_text((30, 80), 'DRAFT TWO')
            doc.save(self.source)
        self.dialog.btn_analyze.click()
        self.wait_idle()
        # Two selections before the debounce fires: only the last page renders.
        self.dialog.objects.topLevelItem(0).setCheckState(0, Qt.Checked)
        self.dialog.objects.topLevelItem(1).setCheckState(0, Qt.Checked)
        self.wait_idle()
        self.dialog.confirm.setChecked(True)
        self.assertFalse(self.dialog.btn_save.isEnabled())
        self.dialog.page_number.setValue(1)
        self.wait_idle()
        self.assertTrue(self.dialog.btn_save.isEnabled())
        self.dialog.objects.topLevelItem(0).setCheckState(0, Qt.Unchecked)
        self.assertFalse(self.dialog.confirm.isChecked())

    def test_main_window_entry_passes_selected_pdf(self):
        from types import SimpleNamespace
        from unittest.mock import Mock
        from ratools_pdf.ui.main_window import MainWindow

        selected = Mock()
        selected.text.return_value = str(self.source)
        view = SimpleNamespace(tree=Mock())
        view.tree.selectedItems.return_value = [selected]
        with patch('ratools_pdf.ui.dialogs.watermark_dialog.WatermarkDialog') as dialog:
            MainWindow.show_watermark_dialog(view)
        dialog.assert_called_once_with(initial_file=str(self.source), parent=view)
        dialog.return_value.exec.assert_called_once()
        dialog.return_value.deleteLater.assert_called_once()

    def test_real_main_window_exposes_watermark_button(self):
        from ratools_pdf.ui.main_window import MainWindow

        window = MainWindow(settings_path=str(Path(self.temp.name) / 'settings.ini'))
        try:
            window.show()
            index = next(i for i, module in enumerate(window.MODULES_DATA)
                         if module['title'] == '内容合规与安全性')
            window.switch_settings_page(index)
            self.app.processEvents()
            self.assertTrue(window.btn_remove_watermark.isVisible())
            self.assertTrue(window.btn_remove_watermark.isEnabled())
            with patch('ratools_pdf.ui.dialogs.watermark_dialog.WatermarkDialog') as dialog:
                window.btn_remove_watermark.click()
            dialog.return_value.exec.assert_called_once()
        finally:
            window.close()
            window.deleteLater()
            QCoreApplication.sendPostedEvents(window, QEvent.DeferredDelete)
            self.assertFalse(isValid(window))
