"""页面装配与拆分后端引擎及对话框单元测试。"""

import glob
import os
import tempfile
import unittest

import fitz
from PySide6.QtWidgets import QApplication

from ratools_pdf.pdf import qpdf
from ratools_pdf.ui.dialogs.page_assembly_dialog import PageAssemblyDialog


app = QApplication.instance() or QApplication([])


def _create_sample_pdf(path: str, page_count: int = 3, prefix: str = "P"):
    doc = fitz.open()
    for i in range(page_count):
        p = doc.new_page()
        p.insert_text((50, 50), f"{prefix}_{i+1}")
    doc.save(path)
    doc.close()


class PageAssemblyEngineTests(unittest.TestCase):
    def test_parse_page_range_patterns(self):
        self.assertEqual(qpdf.parse_page_range("1-3", 10), [1, 2, 3])
        self.assertEqual(qpdf.parse_page_range("z-1", 4), [4, 3, 2, 1])
        self.assertEqual(qpdf.parse_page_range("1-z:odd", 5), [1, 3, 5])
        self.assertEqual(qpdf.parse_page_range("1-z:even", 5), [2, 4])
        self.assertEqual(qpdf.parse_page_range("1, 3, 5-6", 10), [1, 3, 5, 6])
        self.assertEqual(qpdf.parse_page_range("all", 3), [1, 2, 3])
        self.assertEqual(qpdf.parse_page_range("1-z", 3), [1, 2, 3])

    def test_assemble_pages_multi_document(self):
        with tempfile.TemporaryDirectory() as tmp:
            doc1 = os.path.join(tmp, "doc1.pdf")
            doc2 = os.path.join(tmp, "doc2.pdf")
            out = os.path.join(tmp, "assembled.pdf")

            _create_sample_pdf(doc1, 3, "Doc1")
            _create_sample_pdf(doc2, 2, "Doc2")

            specs = [
                qpdf.PageSpec(file_path=doc1, page_range="1-2"),
                qpdf.PageSpec(file_path=doc2, page_range="z-1"),
            ]
            res = qpdf.assemble_pages(specs, out, linearize=True)
            self.assertTrue(res.is_success, res.stderr)

            doc_out = fitz.open(out)
            self.assertEqual(doc_out.page_count, 4)
            text0 = doc_out[0].get_text()
            text2 = doc_out[2].get_text()
            doc_out.close()

            self.assertIn("Doc1_1", text0)
            self.assertIn("Doc2_2", text2)

    def test_assemble_pages_with_rotation(self):
        with tempfile.TemporaryDirectory() as tmp:
            doc1 = os.path.join(tmp, "doc1.pdf")
            out = os.path.join(tmp, "rotated.pdf")

            _create_sample_pdf(doc1, 3, "Doc")

            specs = [
                qpdf.PageSpec(file_path=doc1, page_range="1-2", rotation=90),
                qpdf.PageSpec(file_path=doc1, page_range="3", rotation=180),
            ]
            res = qpdf.assemble_pages(specs, out)
            self.assertTrue(res.is_success, res.stderr)

            doc_out = fitz.open(out)
            self.assertEqual(doc_out.page_count, 3)
            self.assertEqual(doc_out[0].rotation, 90)
            self.assertEqual(doc_out[1].rotation, 90)
            self.assertEqual(doc_out[2].rotation, 180)
            doc_out.close()

    def test_split_pages_into_single_pages(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = os.path.join(tmp, "src.pdf")
            pattern = os.path.join(tmp, "split_%d.pdf")

            _create_sample_pdf(source, 4, "Page")

            res = qpdf.split_pages(source, pattern, pages_per_file=1)
            self.assertTrue(res.is_success, res.stderr)

            splits = sorted(glob.glob(os.path.join(tmp, "split_*.pdf")))
            self.assertEqual(len(splits), 4)

            doc_split1 = fitz.open(splits[0])
            self.assertEqual(doc_split1.page_count, 1)
            self.assertIn("Page_1", doc_split1[0].get_text())
            doc_split1.close()

    def test_split_pages_into_chunks(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = os.path.join(tmp, "src.pdf")
            pattern = os.path.join(tmp, "chunk_%d.pdf")

            _create_sample_pdf(source, 5, "Page")

            res = qpdf.split_pages(source, pattern, pages_per_file=2)
            self.assertTrue(res.is_success, res.stderr)

            splits = sorted(glob.glob(os.path.join(tmp, "chunk_*.pdf")))
            self.assertEqual(len(splits), 3)

            doc0 = fitz.open(splits[0])
            self.assertEqual(doc0.page_count, 2)
            doc0.close()

            doc2 = fitz.open(splits[2])
            self.assertEqual(doc2.page_count, 1)
            doc2.close()


class PageAssemblyDialogUiTests(unittest.TestCase):
    def test_dialog_add_files_and_modes(self):
        with tempfile.TemporaryDirectory() as tmp:
            doc1 = os.path.join(tmp, "file1.pdf")
            doc2 = os.path.join(tmp, "file2.pdf")
            _create_sample_pdf(doc1, 2)
            _create_sample_pdf(doc2, 3)

            dlg = PageAssemblyDialog(initial_files=[doc1, doc2])
            self.assertEqual(dlg.table.rowCount(), 2)

            # 验证各列初始值
            self.assertEqual(dlg.table.item(0, 1).text(), "file1.pdf")
            self.assertEqual(dlg.table.item(0, 2).text(), "2")
            self.assertEqual(dlg.table.item(0, 3).text(), "1-z")

            # 切换到拆分模式
            dlg.rb_split.setChecked(True)
            dlg._on_mode_changed()
            self.assertFalse(dlg.split_card.isHidden())

            # 切换回合并模式
            dlg.rb_merge.setChecked(True)
            dlg._on_mode_changed()
            self.assertTrue(dlg.split_card.isHidden())

            dlg.close()


if __name__ == "__main__":
    unittest.main()
