"""Phase 3：安全加密、图层套印(Overlay/Underlay)与逻辑页码单元测试。"""

import os
import tempfile
import unittest

import fitz
from PySide6.QtWidgets import QApplication

from ratools_pdf.pdf import qpdf
from ratools_pdf.ui.dialogs.overlay_dialog import OverlayUnderlayDialog
from ratools_pdf.ui.dialogs.security_center_dialog import SecurityCenterDialog


app = QApplication.instance() or QApplication([])


def _create_sample_pdf(path: str, page_count: int = 3, text: str = "Sample"):
    doc = fitz.open()
    for i in range(page_count):
        p = doc.new_page()
        p.insert_text((50, 50), f"{text} Page {i+1}")
    doc.save(path)
    doc.close()


class SecurityAndOverlayEngineTests(unittest.TestCase):
    def test_encrypt_pdf_aes256_dual_password(self):
        with tempfile.TemporaryDirectory() as tmp:
            src = os.path.join(tmp, "src.pdf")
            out = os.path.join(tmp, "enc.pdf")
            _create_sample_pdf(src, 2, "Secret Content")

            res = qpdf.encrypt_pdf(
                input_pdf=src,
                output_pdf=out,
                user_password="userpass123",
                owner_password="ownerpass123",
                bits=256,
                print_perm="full",
                modify_perm="none",
                extract=False,
                accessibility=True,
                cleartext_metadata=True,
                linearize=True,
            )
            self.assertTrue(res.is_success, res.stderr)
            self.assertTrue(os.path.exists(out))

            doc = fitz.open(out)
            self.assertEqual(doc.needs_pass, 1)
            # 用户密码认证返回 2
            auth = doc.authenticate("userpass123")
            self.assertEqual(auth, 2)
            self.assertEqual(doc.page_count, 2)
            doc.close()

    def test_encrypt_pdf_owner_only_permissions(self):
        with tempfile.TemporaryDirectory() as tmp:
            src = os.path.join(tmp, "src.pdf")
            out = os.path.join(tmp, "owner_only.pdf")
            _create_sample_pdf(src, 2, "Public Read")

            res = qpdf.encrypt_pdf(
                input_pdf=src,
                output_pdf=out,
                owner_password="adminpassword",
                bits=256,
                print_perm="none",
                modify_perm="none",
            )
            self.assertTrue(res.is_success, res.stderr)

            # 无需密码直接可打开
            doc = fitz.open(out)
            self.assertEqual(doc.needs_pass, 0)
            # 校验 qpdf 的加密限制报告
            self.assertTrue(qpdf.qpdf_reports_restrictions(out))
            doc.close()

    def test_apply_overlay_watermark(self):
        with tempfile.TemporaryDirectory() as tmp:
            target = os.path.join(tmp, "report.pdf")
            stamp = os.path.join(tmp, "confidential.pdf")
            out = os.path.join(tmp, "stamped.pdf")

            _create_sample_pdf(target, 3, "Report")
            _create_sample_pdf(stamp, 1, "CONFIDENTIAL")

            res = qpdf.apply_overlay_underlay(
                input_pdf=target,
                output_pdf=out,
                overlay_file=stamp,
                repeat="1-z",
                linearize=True,
            )
            self.assertTrue(res.is_success, res.stderr)

            doc = fitz.open(out)
            self.assertEqual(doc.page_count, 3)
            for i in range(3):
                text = doc[i].get_text()
                self.assertIn("Report", text)
                self.assertIn("CONFIDENTIAL", text)
            doc.close()

    def test_apply_underlay_letterhead(self):
        with tempfile.TemporaryDirectory() as tmp:
            target = os.path.join(tmp, "letter.pdf")
            letterhead = os.path.join(tmp, "header.pdf")
            out = os.path.join(tmp, "letter_stamped.pdf")

            _create_sample_pdf(target, 2, "Body")
            _create_sample_pdf(letterhead, 1, "HEADER_LOGO")

            res = qpdf.apply_overlay_underlay(
                input_pdf=target,
                output_pdf=out,
                underlay_file=letterhead,
                repeat="1-z",
            )
            self.assertTrue(res.is_success, res.stderr)

            doc = fitz.open(out)
            self.assertEqual(doc.page_count, 2)
            self.assertIn("HEADER_LOGO", doc[0].get_text())
            doc.close()

    def test_set_and_remove_page_labels(self):
        with tempfile.TemporaryDirectory() as tmp:
            src = os.path.join(tmp, "book.pdf")
            out_labeled = os.path.join(tmp, "book_labeled.pdf")
            out_removed = os.path.join(tmp, "book_unlabeled.pdf")

            _create_sample_pdf(src, 6, "Page")

            # 前言 1:r (i, ii), 正文 3:D (1, 2, 3, 4)
            res1 = qpdf.set_page_labels(src, out_labeled, ["1:r", "3:D"])
            self.assertTrue(res1.is_success, res1.stderr)

            doc_labeled = fitz.open(out_labeled)
            # PyMuPDF 支持 get_page_labels()
            labels = [p.get_label() for p in doc_labeled]
            doc_labeled.close()
            self.assertEqual(labels[:4], ["i", "ii", "1", "2"])

            # 移除页面标签
            res2 = qpdf.remove_page_labels(out_labeled, out_removed)
            self.assertTrue(res2.is_success, res2.stderr)

            doc_removed = fitz.open(out_removed)
            labels_removed = [p.get_label() for p in doc_removed]
            doc_removed.close()
            # 移除后恢复默认无自定义标签（返回空字符串）
            self.assertEqual(labels_removed[:4], ["", "", "", ""])


class SecurityAndOverlayDialogUiTests(unittest.TestCase):
    def test_security_dialog_init(self):
        with tempfile.TemporaryDirectory() as tmp:
            pdf_path = os.path.join(tmp, "test.pdf")
            _create_sample_pdf(pdf_path, 1)

            dlg = SecurityCenterDialog(initial_file=pdf_path)
            self.assertEqual(dlg.txt_input.text(), pdf_path)
            self.assertIn("_encrypted.pdf", dlg.txt_output.text())
            dlg.close()

    def test_overlay_dialog_init(self):
        with tempfile.TemporaryDirectory() as tmp:
            pdf_path = os.path.join(tmp, "test.pdf")
            _create_sample_pdf(pdf_path, 1)

            dlg = OverlayUnderlayDialog(initial_file=pdf_path)
            self.assertEqual(dlg.txt_target.text(), pdf_path)
            self.assertIn("_stamped.pdf", dlg.txt_output.text())
            self.assertTrue(dlg.rb_overlay.isChecked())
            dlg.close()


if __name__ == "__main__":
    unittest.main()
