import os
import tempfile
import unittest

import fitz

from pdf_processor import PDFProcessor


class PDFProcessorRoundTripTests(unittest.TestCase):
    def test_bookmark_export_import_preserves_external_and_internal_targets(self):
        with tempfile.TemporaryDirectory() as tmp:
            source_pdf = os.path.join(tmp, "source.pdf")
            csv_path = os.path.join(tmp, "bookmarks.csv")
            output_pdf = os.path.join(tmp, "output.pdf")

            doc = fitz.open()
            doc.new_page()
            doc.new_page()
            doc.set_toc([
                [1, "External", -1, {"kind": fitz.LINK_URI, "uri": "https://example.com"}],
                [1, "Internal", 2, {"kind": fitz.LINK_GOTO, "page": 1, "to": fitz.Point(144, 288), "zoom": 2.0}],
            ])
            doc.save(source_pdf)
            doc.close()

            PDFProcessor.export_bookmarks(source_pdf, csv_path)
            PDFProcessor.import_bookmarks(source_pdf, csv_path, output_pdf)

            restored = fitz.open(output_pdf)
            toc = restored.get_toc(simple=False)
            restored.close()

            self.assertEqual(toc[0][3].get("kind"), fitz.LINK_URI)
            self.assertEqual(toc[0][3].get("uri"), "https://example.com")
            self.assertEqual(toc[1][3].get("kind"), fitz.LINK_GOTO)
            self.assertEqual(tuple(toc[1][3].get("to")), (144.0, 288.0))
            self.assertEqual(toc[1][3].get("zoom"), 2.0)

    def test_link_export_import_preserves_internal_target_coordinates(self):
        with tempfile.TemporaryDirectory() as tmp:
            source_pdf = os.path.join(tmp, "source.pdf")
            json_path = os.path.join(tmp, "links.json")
            output_pdf = os.path.join(tmp, "output.pdf")

            doc = fitz.open()
            doc.new_page()
            doc.new_page()
            doc[0].insert_link({
                "kind": fitz.LINK_GOTO,
                "from": fitz.Rect(70, 60, 180, 80),
                "page": 1,
                "to": fitz.Point(144, 288),
            })
            doc.save(source_pdf)
            doc.close()

            PDFProcessor.export_links(source_pdf, json_path)
            PDFProcessor.import_links(source_pdf, json_path, output_pdf)

            restored = fitz.open(output_pdf)
            link = restored[0].get_links()[0]
            restored.close()

            self.assertEqual(tuple(link.get("to")), (144.0, 288.0))

    def test_smart_mode_reports_unsupported_options_as_forced(self):
        with tempfile.TemporaryDirectory() as tmp:
            source_pdf = os.path.join(tmp, "source.pdf")
            doc = fitz.open()
            doc.new_page(width=595, height=842)
            doc.save(source_pdf)
            doc.close()

            result = PDFProcessor.resolve_processing_options(source_pdf, {"page_size_a4"}, "smart")

            self.assertIn("page_size_a4", result["forced_unsupported"])


if __name__ == "__main__":
    unittest.main()
