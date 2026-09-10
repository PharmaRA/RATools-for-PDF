"""Phase 4：PDF 深度体检、语法探针与受损修复单元测试。"""

import os
import tempfile
import unittest

import fitz
from PySide6.QtWidgets import QApplication

from ratools_pdf.pdf import qpdf
from ratools_pdf.ui.dialogs.pdf_inspector_dialog import PdfInspectorDialog


app = QApplication.instance() or QApplication([])


def _create_sample_pdf(path: str, page_count: int = 2, text: str = "Test"):
    doc = fitz.open()
    for i in range(page_count):
        p = doc.new_page()
        p.insert_text((50, 50), f"{text} {i+1}")
    doc.save(path)
    doc.close()


class PdfInspectorAndRepairEngineTests(unittest.TestCase):
    def test_syntax_check_healthy_and_corrupt(self):
        with tempfile.TemporaryDirectory() as tmp:
            healthy = os.path.join(tmp, "healthy.pdf")
            _create_sample_pdf(healthy, 2)

            res = qpdf.check_pdf_syntax(healthy)
            self.assertTrue(res["is_healthy"])
            self.assertEqual(res["returncode"], 0)
            self.assertIn("healthy", res["output"].lower())

    def test_get_linearization_report(self):
        with tempfile.TemporaryDirectory() as tmp:
            raw_pdf = os.path.join(tmp, "raw.pdf")
            lin_pdf = os.path.join(tmp, "lin.pdf")
            _create_sample_pdf(raw_pdf, 3)

            # 未线性化报告
            rep_raw = qpdf.get_linearization_report(raw_pdf)
            self.assertIn("not linearized", rep_raw.lower())

            # 线性化后报告
            qpdf.rewrite_with_qpdf(raw_pdf, lin_pdf, linearize=True)
            rep_lin = qpdf.get_linearization_report(lin_pdf)
            self.assertIn("linearization data", rep_lin.lower())
            self.assertNotIn("not linearized", rep_lin.lower())

    def test_inspect_pdf_json_ast(self):
        with tempfile.TemporaryDirectory() as tmp:
            pdf_path = os.path.join(tmp, "doc.pdf")
            _create_sample_pdf(pdf_path, 2)

            ast = qpdf.inspect_pdf_json(pdf_path)
            self.assertIsInstance(ast, dict)
            self.assertIn("pages", ast)
            self.assertEqual(len(ast["pages"]), 2)
            self.assertIn("qpdf", ast)
            self.assertIn("encrypt", ast)

    def test_repair_pdf_engine(self):
        with tempfile.TemporaryDirectory() as tmp:
            src = os.path.join(tmp, "src.pdf")
            repaired = os.path.join(tmp, "repaired.pdf")
            _create_sample_pdf(src, 3)

            res = qpdf.repair_pdf(src, repaired)
            self.assertTrue(res.is_success, res.stderr)
            self.assertTrue(os.path.exists(repaired))

            doc = fitz.open(repaired)
            self.assertEqual(doc.page_count, 3)
            doc.close()


class PdfInspectorDialogUiTests(unittest.TestCase):
    def test_inspector_dialog_init(self):
        with tempfile.TemporaryDirectory() as tmp:
            pdf_path = os.path.join(tmp, "sample.pdf")
            _create_sample_pdf(pdf_path, 1)

            dlg = PdfInspectorDialog(initial_file=pdf_path)
            self.assertEqual(dlg.txt_path.text(), pdf_path)
            self.assertEqual(dlg.tabs.count(), 5)
            self.assertEqual(dlg.tabs.tabText(0), "语法健康体检")
            self.assertEqual(dlg.tabs.tabText(1), "加密与权限透视")
            self.assertEqual(dlg.tabs.tabText(2), "线性化(Web快速视图)")
            self.assertEqual(dlg.tabs.tabText(3), "底层 JSON 探针")
            self.assertEqual(dlg.tabs.tabText(4), "受损抢救修复")
            dlg.close()


if __name__ == "__main__":
    unittest.main()
