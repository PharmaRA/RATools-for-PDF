import os
import tempfile
import unittest
from unittest.mock import patch

import fitz

from ratools_pdf.controllers.main_controller import MainController
from ratools_pdf.pdf import precheck
from ratools_pdf.pdf.processor import PDFProcessor


def _make_unsigned_pdf(path):
    doc = fitz.open()
    doc.new_page()
    doc.save(path)
    doc.close()


def _make_signed_pdf(path):
    """构造带签名域标志 (SigFlags=3) 的 PDF，等价于监管 PDF 的已签名状态。"""
    doc = fitz.open()
    doc.new_page()
    catalog_xref = doc.pdf_catalog()
    acroform_xref = doc.get_new_xref()
    doc.update_object(acroform_xref, "<< /Fields [] /SigFlags 3 >>")
    doc.xref_set_key(catalog_xref, "AcroForm", f"{acroform_xref} 0 R")
    doc.save(path)
    doc.close()


class SignatureDetectionTests(unittest.TestCase):
    def test_unsigned_pdf_reports_no_signature(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "unsigned.pdf")
            _make_unsigned_pdf(path)
            self.assertFalse(PDFProcessor._pdf_has_signature(path))

    def test_signed_pdf_is_detected(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "signed.pdf")
            _make_signed_pdf(path)
            self.assertTrue(PDFProcessor._pdf_has_signature(path))

    def test_missing_file_reports_no_signature(self):
        self.assertFalse(PDFProcessor._pdf_has_signature(r"C:\does\not\exist.pdf"))

    def test_empty_path_reports_no_signature(self):
        self.assertFalse(precheck._pdf_has_signature(""))


class _FakeView:
    def __init__(self, prompt_result="skip"):
        self.prompt_result = prompt_result
        self.prompt_calls = []
        self.warnings = []

    def show_signed_files_prompt(self, signed_files):
        self.prompt_calls.append(list(signed_files))
        return self.prompt_result

    def show_warning_message(self, title, message):
        self.warnings.append((title, message))


class _FakeController:
    """仅承载 _prompt_skip_signed_files 所需的最小状态。"""

    def __init__(self, view):
        self.view = view
        self.process_logs = ""

    _prompt_skip_signed_files = MainController._prompt_skip_signed_files


class PromptSkipSignedFilesTests(unittest.TestCase):
    ALL_FILES = [r"C:\docs\a.pdf", r"C:\docs\signed.pdf", r"C:\docs\b.pdf"]
    SIGNED = {r"C:\docs\signed.pdf"}

    def _run(self, prompt_result):
        controller = _FakeController(_FakeView(prompt_result))
        with patch.object(PDFProcessor, "_pdf_has_signature", side_effect=lambda p: p in self.SIGNED):
            result = controller._prompt_skip_signed_files(list(self.ALL_FILES))
        return controller, result

    def test_no_signed_files_skips_prompt(self):
        controller = _FakeController(_FakeView("cancel"))
        with patch.object(PDFProcessor, "_pdf_has_signature", return_value=False):
            result = controller._prompt_skip_signed_files(list(self.ALL_FILES))
        self.assertEqual(result, self.ALL_FILES)
        self.assertEqual(controller.view.prompt_calls, [])

    def test_cancel_returns_none(self):
        _controller, result = self._run("cancel")
        self.assertIsNone(result)

    def test_skip_removes_signed_files(self):
        controller, result = self._run("skip")
        self.assertEqual(result, [r"C:\docs\a.pdf", r"C:\docs\b.pdf"])
        self.assertEqual(controller.view.prompt_calls, [list(self.SIGNED)])
        self.assertIn("已按用户选择跳过", controller.process_logs)

    def test_process_all_keeps_signed_files(self):
        controller, result = self._run("process_all")
        self.assertEqual(result, self.ALL_FILES)
        self.assertIn("仍然处理全部", controller.process_logs)


if __name__ == "__main__":
    unittest.main()
