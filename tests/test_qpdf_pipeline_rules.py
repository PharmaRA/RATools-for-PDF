"""qpdf 规则在 process_document 管线与预检中的端到端集成测试。"""

import os
import tempfile
import unittest

import fitz

from ratools_pdf.pdf import precheck, qpdf
from ratools_pdf.pdf.processor import PDFProcessor


class QpdfPipelineRulesTests(unittest.TestCase):
    def test_flatten_rotation_rule(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = os.path.join(tmp, "rotated.pdf")
            output = os.path.join(tmp, "out.pdf")

            doc = fitz.open()
            page = doc.new_page()
            page.set_rotation(90)
            page.insert_text((100, 100), "Rotated text")
            doc.save(source)
            doc.close()

            ok, msg = PDFProcessor.process_document(
                source, output, {"flatten_rotation"}, processing_mode="force"
            )
            self.assertTrue(ok, msg)
            self.assertIn("已展平页面物理旋转角", msg)

            doc_out = fitz.open(output)
            self.assertEqual(doc_out[0].rotation, 0, "旋转角未能归零展平")
            doc_out.close()

    def test_flatten_annotations_rule(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = os.path.join(tmp, "annotated.pdf")
            output = os.path.join(tmp, "out.pdf")

            doc = fitz.open()
            page = doc.new_page()
            page.insert_text((100, 100), "Annotated text")
            page.add_rect_annot(fitz.Rect(50, 50, 150, 150))
            doc.save(source)
            doc.close()

            # 验证原始文件确实含有非链接注释
            doc_check = fitz.open(source)
            self.assertEqual(len(list(doc_check[0].annots())), 1)
            doc_check.close()

            ok, msg = PDFProcessor.process_document(
                source, output, {"flatten_annotations"}, processing_mode="force"
            )
            self.assertTrue(ok, msg)
            self.assertIn("已展平所有注释与批注", msg)

            doc_out = fitz.open(output)
            self.assertEqual(len(list(doc_out[0].annots())), 0, "注释未能展平入内容流")
            doc_out.close()

    def test_generate_object_streams_rule(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = os.path.join(tmp, "objstms.pdf")
            output = os.path.join(tmp, "out.pdf")

            doc = fitz.open()
            for i in range(3):
                p = doc.new_page()
                p.insert_text((50, 50), f"Page {i}")
            doc.save(source)
            doc.close()

            ok, msg = PDFProcessor.process_document(
                source, output, {"generate_object_streams"}, processing_mode="force"
            )
            self.assertTrue(ok, msg)
            self.assertIn("已生成对象流压缩", msg)
            self.assertTrue(os.path.exists(output))

    def test_realloc_flate_rule(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = os.path.join(tmp, "realloc.pdf")
            output = os.path.join(tmp, "out.pdf")

            doc = fitz.open()
            p = doc.new_page()
            p.insert_text((50, 50), "Realloc flate test")
            doc.save(source)
            doc.close()

            ok, msg = PDFProcessor.process_document(
                source, output, {"realloc_flate"}, processing_mode="force"
            )
            self.assertTrue(ok, msg)
            self.assertIn("已优化重压缩内容流", msg)
            self.assertTrue(os.path.exists(output))

    def test_combined_qpdf_pipeline(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = os.path.join(tmp, "combo.pdf")
            output = os.path.join(tmp, "out.pdf")

            doc = fitz.open()
            page = doc.new_page()
            page.set_rotation(270)
            page.insert_text((50, 50), "Combo test")
            page.add_rect_annot(fitz.Rect(20, 20, 80, 80))
            doc.save(source)
            doc.close()

            options = {
                "flatten_rotation",
                "flatten_annotations",
                "generate_object_streams",
                "realloc_flate",
                "fast_web_view",
            }
            ok, msg = PDFProcessor.process_document(
                source, output, options, processing_mode="force"
            )
            self.assertTrue(ok, msg)
            self.assertIn("已展平页面物理旋转角", msg)
            self.assertIn("已展平所有注释与批注", msg)
            self.assertIn("已生成对象流压缩", msg)
            self.assertIn("已优化重压缩内容流", msg)
            self.assertIn("已启用快速网页浏览", msg)

            doc_out = fitz.open(output)
            self.assertEqual(doc_out[0].rotation, 0)
            self.assertEqual(len(list(doc_out[0].annots())), 0)
            doc_out.close()

            self.assertTrue(qpdf.is_pdf_linearized(output), "未能完成线性化")

    def test_precheck_detects_flatten_rotation(self):
        with tempfile.TemporaryDirectory() as tmp:
            pdf_path = os.path.join(tmp, "rot.pdf")
            doc = fitz.open()
            p = doc.new_page()
            p.set_rotation(90)
            doc.save(pdf_path)
            doc.close()

            report = precheck.build_precheck_report(pdf_path, selected_options={"flatten_rotation"})
            suggestions = report.get("suggestions", {})
            self.assertIn("flatten_rotation", suggestions)

    def test_precheck_detects_flatten_annotations(self):
        with tempfile.TemporaryDirectory() as tmp:
            pdf_path = os.path.join(tmp, "annot.pdf")
            doc = fitz.open()
            p = doc.new_page()
            p.add_rect_annot(fitz.Rect(10, 10, 50, 50))
            doc.save(pdf_path)
            doc.close()

            report = precheck.build_precheck_report(pdf_path, selected_options={"flatten_annotations"})
            suggestions = report.get("suggestions", {})
            self.assertIn("flatten_annotations", suggestions)


if __name__ == "__main__":
    unittest.main()
