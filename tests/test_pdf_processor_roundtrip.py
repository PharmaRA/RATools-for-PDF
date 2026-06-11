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


if __name__ == "__main__":
    unittest.main()
