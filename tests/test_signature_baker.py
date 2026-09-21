"""数字签名外观靶向烘焙模块单元测试。"""

import os
import tempfile
import unittest

import fitz

from ratools_pdf.pdf import qpdf
from ratools_pdf.pdf.signature_baker import bake_digital_signatures, has_digital_signatures


class SignatureBakerTests(unittest.TestCase):
    def test_has_digital_signatures_unsigned(self):
        with tempfile.TemporaryDirectory() as tmp:
            pdf_path = os.path.join(tmp, "unsigned.pdf")
            doc = fitz.open()
            doc.new_page()
            doc.save(pdf_path)
            doc.close()

            self.assertFalse(has_digital_signatures(pdf_path))

    def test_has_digital_signatures_with_widget(self):
        with tempfile.TemporaryDirectory() as tmp:
            pdf_path = os.path.join(tmp, "signed.pdf")
            doc = fitz.open()
            page = doc.new_page()
            widget = fitz.Widget()
            widget.field_type = fitz.PDF_WIDGET_TYPE_SIGNATURE
            widget.field_name = "Signature1"
            widget.rect = fitz.Rect(100, 100, 300, 200)
            page.add_widget(widget)
            doc.save(pdf_path)
            doc.close()

            self.assertTrue(has_digital_signatures(pdf_path))

    def test_bake_signatures_preserves_appearance_and_links(self):
        with tempfile.TemporaryDirectory() as tmp:
            in_pdf = os.path.join(tmp, "signed_doc.pdf")
            baked_pdf = os.path.join(tmp, "baked_doc.pdf")

            # 构造包含正文文本、超链接、注释以及数字签名控件的文档
            doc = fitz.open()
            page = doc.new_page(width=600, height=800)
            page.insert_text(fitz.Point(50, 50), "Study Report Clinical Summary")
            page.insert_link(
                {"kind": fitz.LINK_URI, "from": fitz.Rect(50, 80, 250, 100), "uri": "https://ectd.example.com"}
            )
            page.add_text_annot(fitz.Point(50, 120), "Auditor Note")

            widget = fitz.Widget()
            widget.field_type = fitz.PDF_WIDGET_TYPE_SIGNATURE
            widget.field_name = "Signature_QA"
            widget.rect = fitz.Rect(100, 300, 400, 450)
            page.add_widget(widget)
            doc.save(in_pdf)
            doc.close()

            # 执行靶向烘焙
            count = bake_digital_signatures(in_pdf, baked_pdf, dpi=150)
            self.assertEqual(count, 1)

            # 验证烘焙后的 PDF
            doc_baked = fitz.open(baked_pdf)
            p = doc_baked[0]

            # 1. 签名控件已被安全清除
            self.assertEqual(len(list(p.widgets())), 0)

            # 2. 签名外观已被固化为正文图像
            self.assertEqual(len(p.get_images()), 1)

            # 3. 超链接 100% 保持完好且可点击
            links = list(p.get_links())
            self.assertEqual(len(links), 1)
            self.assertEqual(links[0].get("uri"), "https://ectd.example.com")

            # 4. 其他注释（便签/批注）100% 保持完好
            annots = list(p.annots())
            self.assertEqual(len(annots), 1)

            # 5. 正文文本完好
            self.assertIn("Study Report Clinical Summary", p.get_text())
            doc_baked.close()

    def test_bake_signatures_on_rotated_page(self):
        with tempfile.TemporaryDirectory() as tmp:
            in_pdf = os.path.join(tmp, "rotated.pdf")
            baked_pdf = os.path.join(tmp, "baked_rotated.pdf")

            doc = fitz.open()
            page = doc.new_page(width=600, height=800)
            page.set_rotation(90)

            widget = fitz.Widget()
            widget.field_type = fitz.PDF_WIDGET_TYPE_SIGNATURE
            widget.field_name = "SigRot"
            widget.rect = fitz.Rect(100, 100, 300, 250)
            page.add_widget(widget)
            doc.save(in_pdf)
            doc.close()

            count = bake_digital_signatures(in_pdf, baked_pdf, dpi=150)
            self.assertEqual(count, 1)

            doc_chk = fitz.open(baked_pdf)
            p = doc_chk[0]
            self.assertEqual(len(list(p.widgets())), 0)
            self.assertEqual(len(p.get_images()), 1)
            img_rect = p.get_image_rects(p.get_images()[0][0])[0]
            self.assertAlmostEqual(img_rect.x0, 100, delta=1.0)
            self.assertAlmostEqual(img_rect.y0, 100, delta=1.0)
            self.assertAlmostEqual(img_rect.x1, 300, delta=1.0)
            self.assertAlmostEqual(img_rect.y1, 250, delta=1.0)
            doc_chk.close()

    def test_decrypt_pdf_end_to_end_with_signature(self):
        with tempfile.TemporaryDirectory() as tmp:
            in_pdf = os.path.join(tmp, "signed.pdf")
            out_pdf = os.path.join(tmp, "decrypted.pdf")

            # 构造包含签名和超链接的文档
            doc = fitz.open()
            page = doc.new_page(width=600, height=800)
            page.insert_text(fitz.Point(50, 50), "Test Content")
            page.insert_link({"kind": fitz.LINK_URI, "from": fitz.Rect(50, 80, 200, 100), "uri": "https://test.com"})
            widget = fitz.Widget()
            widget.field_type = fitz.PDF_WIDGET_TYPE_SIGNATURE
            widget.field_name = "Sign1"
            widget.rect = fitz.Rect(100, 200, 300, 350)
            page.add_widget(widget)
            doc.save(in_pdf)
            doc.close()

            # 调用 decrypt_pdf
            res = qpdf.decrypt_pdf(
                in_pdf,
                out_pdf,
                remove_restrictions=True,
                remove_acroform=False,
                linearize=False,
                bake_signatures=True,
            )
            self.assertTrue(res.is_success, res.stderr)

            # 校验最终输出：签名控件已除、外观图像固化、超链接完好
            doc_out = fitz.open(out_pdf)
            p_out = doc_out[0]
            self.assertEqual(len(list(p_out.widgets())), 0)
            self.assertEqual(len(p_out.get_images()), 1)
            links = list(p_out.get_links())
            self.assertEqual(len(links), 1)
            self.assertEqual(links[0].get("uri"), "https://test.com")
            doc_out.close()


if __name__ == "__main__":
    unittest.main()
