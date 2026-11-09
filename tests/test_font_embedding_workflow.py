"""PD05 UI behavior; real widgets/workflow with only OS launch substituted."""

from pathlib import Path
import tempfile
import threading
import time
import unittest

import fitz
from PySide6.QtWidgets import QApplication

from ratools_pdf.controllers.font_embedding_controller import FontEmbeddingController
from ratools_pdf.ui.dialogs.font_embedding import ManualFontEmbeddingDialog


class FontEmbeddingDialogTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def make_dialog(self, *, callback=None):
        try:
            dialog = ManualFontEmbeddingDialog(
                ["sample.pdf"], lambda: (True, "已打开，等待手动保存与复检"),
                capability_text="自动修复不可用；仅手动交接", recheck_callback=callback,
            )
        except TypeError:
            self.fail("Manual font dialog does not expose capability and after-save recheck")
        self.addCleanup(dialog.close)
        return dialog

    def finish_recheck(self, dialog):
        deadline = time.monotonic() + 5
        while not dialog.btn_recheck.isEnabled() and time.monotonic() < deadline:
            self.app.processEvents()
            time.sleep(0.005)
        self.assertTrue(dialog.btn_recheck.isEnabled(), "background recheck did not finish")

    def test_opened_document_stays_pending_instead_of_repair_success(self):
        dialog = self.make_dialog()
        dialog.open_acrobat()
        self.assertEqual(dialog.status_label.property("state"), "pending")
        self.assertIn("自动修复不可用", dialog.capability_label.text())

    def test_after_save_recheck_can_confirm_risks_and_reenable_controls(self):
        dialog = self.make_dialog(callback=lambda: (False, "未嵌入 1；替代风险 0；状态未知 0"))
        dialog.start_recheck()
        self.finish_recheck(dialog)
        self.assertEqual(dialog.status_label.property("state"), "error")
        self.assertIn("未嵌入 1", dialog.status_label.text())
        self.assertTrue(dialog.btn_open_acrobat.isEnabled())
        self.assertTrue(dialog.btn_close.isEnabled())

    def test_recheck_exception_is_visible_and_does_not_strand_busy_state(self):
        def fail():
            raise RuntimeError("private-input-path")

        dialog = self.make_dialog(callback=fail)
        dialog.start_recheck()
        self.finish_recheck(dialog)
        self.assertEqual(dialog.status_label.property("state"), "error")
        self.assertNotIn("private-input-path", dialog.status_label.text())
        self.assertTrue(dialog.btn_close.isEnabled())

    def test_successful_recheck_is_a_separate_verified_state(self):
        dialog = self.make_dialog(callback=lambda: (True, "字体复检通过：三类风险均为 0"))
        dialog.start_recheck()
        self.finish_recheck(dialog)
        self.assertEqual(dialog.status_label.property("state"), "success")

    def test_background_recheck_keeps_dialog_alive_until_worker_finishes(self):
        entered = threading.Event()
        release = threading.Event()

        def recheck():
            entered.set()
            release.wait(5)
            return False, "字体风险仍存在"

        dialog = self.make_dialog(callback=recheck)
        dialog.show()
        dialog.start_recheck()
        try:
            self.assertTrue(entered.wait(2), "recheck did not start in background")
            self.app.processEvents()
            dialog.close()
            dialog.reject()
            dialog.accept()
            self.assertTrue(dialog.isVisible(), "dialog closed while its worker was active")
            self.assertFalse(dialog.btn_close.isEnabled())
        finally:
            release.set()
            self.finish_recheck(dialog)
        dialog.close()
        self.assertFalse(dialog.isVisible())

    def test_selected_files_receive_capability_and_executable_recheck_callback(self):
        class Item:
            def __init__(self, path):
                self.path = path

            def text(self, column):
                return str(self.path) if column == 1 else "sample"

        class View:
            def __init__(self, path):
                self.tree = self
                self.path = path
                self.kwargs = None

            def selectedItems(self):
                return [Item(self.path)]

            def show_manual_font_embedding_dialog(self, paths, open_callback, **kwargs):
                self.kwargs = kwargs

        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "sample.pdf"
            with fitz.open() as doc:
                doc.new_page()
                doc.save(path)
            view = View(path)
            FontEmbeddingController(view).open_selected_files_in_acrobat()
            self.assertIn("自动修复不可用", view.kwargs["capability_text"])
            verified, message = view.kwargs["recheck_callback"]()
            self.assertTrue(verified)
            self.assertIn("三类风险均为 0", message)

    def test_controller_rechecks_real_pdf_using_all_three_existing_risk_counts(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "sample.pdf"
            with fitz.open() as doc:
                doc.new_page().insert_text((72, 72), "Font sample")
                doc.save(path)
            before = path.read_bytes()
            controller = FontEmbeddingController(object())
            self.assertTrue(hasattr(controller, "_recheck_paths"),
                            "controller has no after-save recheck")
            verified, message = controller._recheck_paths([path])
            self.assertFalse(verified)
            self.assertIn("未嵌入 1", message)
            self.assertIn("替代风险 0", message)
            self.assertIn("状态未知 0", message)
            self.assertEqual(path.read_bytes(), before)


if __name__ == "__main__":
    unittest.main()
