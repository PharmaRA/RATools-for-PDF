import os
import tempfile
import unittest

import fitz

from ratools_pdf.pdf import precheck
from ratools_pdf.pdf.processor import PDFProcessor


# 使用 PyMuPDF 内置的简体中文字体，确保测试文本能被正确写入并抽取；
# 默认 helv 字体无法绘制中文，会导致 get_text 抽不到内容。
def _make_plain_pdf(path, text="正常正文内容"):
    doc = fitz.open()
    page = doc.new_page()
    page.insert_text((72, 72), text, fontname="china-s")
    doc.save(path)
    doc.close()


def _make_pdf_with_annotations(path):
    doc = fitz.open()
    page = doc.new_page()
    page.insert_text((72, 72), "示例文字用于高亮", fontname="china-s")
    # 便签(文本注释)
    page.add_text_annot((100, 100), "这是一个便签")
    # 高亮注释
    page.add_highlight_annot(fitz.Rect(72, 66, 200, 80))
    doc.save(path)
    doc.close()


def _make_pdf_with_text(path, text):
    doc = fitz.open()
    page = doc.new_page()
    page.insert_text((72, 72), text, fontname="china-s")
    doc.save(path)
    doc.close()


class AnnotationPrecheckTests(unittest.TestCase):
    def test_plain_pdf_has_no_annotation_finding(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "plain.pdf")
            _make_plain_pdf(path)
            report = PDFProcessor.build_precheck_report(path)
            self.assertNotIn("annotation_precheck_review", report.get("suggestions", {}))
            self.assertEqual(report.get("annotation_summary", ""), "")

    def test_annotations_are_reported_as_review_only(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "annots.pdf")
            _make_pdf_with_annotations(path)
            report = PDFProcessor.build_precheck_report(path)
            suggestions = report.get("suggestions", {})
            self.assertIn("annotation_precheck_review", suggestions)
            # 复核项不应作为可自动处理的建议规则参与勾选
            self.assertTrue(suggestions["annotation_precheck_review"].get("report_only"))
            self.assertTrue(report.get("annotation_summary"))

    def test_collect_annotation_findings_counts_types(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "annots.pdf")
            _make_pdf_with_annotations(path)
            doc = fitz.open(path)
            try:
                findings = precheck._collect_annotation_findings(doc)
            finally:
                doc.close()
            self.assertTrue(findings["has_annotations"])
            self.assertEqual(findings["count"], 2)


class BrokenReferencePrecheckTests(unittest.TestCase):
    def test_word_broken_reference_zh_is_detected(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "broken_zh.pdf")
            _make_pdf_with_text(path, "错误！未找到引用源。")
            report = PDFProcessor.build_precheck_report(path)
            self.assertIn("broken_reference_precheck_review", report.get("suggestions", {}))
            self.assertTrue(report.get("broken_reference_summary"))

    def test_word_broken_reference_en_is_detected(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "broken_en.pdf")
            _make_pdf_with_text(path, "Error! Reference source not found.")
            report = PDFProcessor.build_precheck_report(path)
            self.assertIn("broken_reference_precheck_review", report.get("suggestions", {}))

    def test_bookmark_not_defined_is_detected(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "bookmark.pdf")
            _make_pdf_with_text(path, "Error! Bookmark not defined.")
            report = PDFProcessor.build_precheck_report(path)
            self.assertIn("broken_reference_precheck_review", report.get("suggestions", {}))

    def test_plain_error_word_does_not_false_positive(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "plain_error.pdf")
            # 正文中正常出现“错误”和“error”，不应触发失效引用误报
            _make_pdf_with_text(path, "本节介绍常见错误与 error handling 的处理方法。")
            report = PDFProcessor.build_precheck_report(path)
            self.assertNotIn("broken_reference_precheck_review", report.get("suggestions", {}))
            self.assertEqual(report.get("broken_reference_summary", ""), "")

    def test_collect_broken_reference_findings_reports_page(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "broken.pdf")
            _make_pdf_with_text(path, "错误！未定义书签。")
            doc = fitz.open(path)
            try:
                findings = precheck._collect_broken_reference_findings(doc)
            finally:
                doc.close()
            self.assertTrue(findings["has_broken_reference"])
            self.assertEqual(findings["count"], 1)


if __name__ == "__main__":
    unittest.main()
