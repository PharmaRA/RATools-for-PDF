"""Do not publish a PDF with lost signature appearances after a rendering failure."""

import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import fitz

from ratools_pdf.pdf import qpdf
from ratools_pdf.pdf.signature_baker import bake_digital_signatures


class SignatureFailureSafetyTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.root = Path(self.directory.name)
        self.source = self.root/'source.pdf'
        self.output = self.root/'output.pdf'
        with fitz.open() as doc:
            page = doc.new_page(width=600,height=800)
            page.insert_text((50,50),'Keep original content')
            for index in range(2):
                widget = fitz.Widget()
                widget.field_type = fitz.PDF_WIDGET_TYPE_SIGNATURE
                widget.field_name = f'Signature{index}'
                widget.rect = fitz.Rect(100,100+index*200,300,200+index*200)
                page.add_widget(widget)
            doc.save(self.source)
        self.original = self.source.read_bytes()
        self.output.write_bytes(b'previous successful output')

    def assert_originals_preserved(self):
        self.assertEqual(self.source.read_bytes(),self.original)
        self.assertEqual(self.output.read_bytes(),b'previous successful output')
        with fitz.open(self.source) as doc:
            self.assertEqual(len(list(doc[0].widgets())),2)

    def test_both_render_paths_failing_do_not_publish_or_remove_signatures(self):
        with patch.object(fitz.Annot,'get_pixmap',side_effect=RuntimeError('render failed')), \
                patch.object(fitz.Page,'get_pixmap',side_effect=RuntimeError('fallback failed')):
            with self.assertRaises(RuntimeError):
                bake_digital_signatures(str(self.source),str(self.output))
        self.assert_originals_preserved()

    def test_image_insertion_failure_does_not_publish_partial_bake(self):
        original_insert = fitz.Page.insert_image
        calls = 0
        def fail_second(page,*args,**kwargs):
            nonlocal calls
            calls += 1
            if calls==2:
                raise RuntimeError('image insertion failed')
            return original_insert(page,*args,**kwargs)
        with patch.object(fitz.Page,'insert_image',fail_second):
            with self.assertRaises(RuntimeError):
                bake_digital_signatures(str(self.source),str(self.output))
        self.assert_originals_preserved()

    def test_widget_removal_failure_stops_before_saving(self):
        with patch.object(fitz.Page,'delete_widget',side_effect=RuntimeError('remove failed')):
            with self.assertRaises(RuntimeError):
                bake_digital_signatures(str(self.source),str(self.output))
        self.assert_originals_preserved()

    def test_in_place_failure_preserves_original_bytes(self):
        with patch.object(fitz.Page,'insert_image',side_effect=RuntimeError('insert failed')):
            with self.assertRaises(RuntimeError):
                bake_digital_signatures(str(self.source),str(self.source))
        self.assert_originals_preserved()

    def test_decrypt_stops_on_bake_error_without_calling_qpdf(self):
        with patch('ratools_pdf.pdf.signature_baker.bake_digital_signatures',side_effect=RuntimeError('render failed')), \
                patch('ratools_pdf.pdf.qpdf.run_qpdf_command') as engine:
            with self.assertRaises(RuntimeError):
                qpdf.decrypt_pdf(str(self.source),str(self.output))
        engine.assert_not_called()
        self.assert_originals_preserved()

    def test_rewrite_stops_on_bake_error_without_calling_qpdf(self):
        with patch('ratools_pdf.pdf.signature_baker.bake_digital_signatures',side_effect=RuntimeError('render failed')), \
                patch('ratools_pdf.pdf.qpdf.run_qpdf_command') as engine:
            with self.assertRaises(RuntimeError):
                qpdf.rewrite_with_qpdf(str(self.source),str(self.output),remove_restrictions=True)
        engine.assert_not_called()
        self.assert_originals_preserved()

    def test_invisible_zero_rect_signatures_can_be_removed_without_rendering(self):
        with fitz.open(self.source) as doc:
            for widget in doc[0].widgets():
                doc.xref_set_key(widget.xref,'Rect','[0 0 0 0]')
            data = doc.tobytes()
        self.source.write_bytes(data)
        with patch.object(fitz.Annot,'get_pixmap',side_effect=AssertionError('invisible signature must not render')):
            self.assertEqual(bake_digital_signatures(str(self.source),str(self.output)),2)
        with fitz.open(self.output) as doc:
            self.assertEqual(len(list(doc[0].widgets())),0)
            self.assertEqual(len(doc[0].get_images()),0)
            self.assertIn('Keep original content',doc[0].get_text())

    def test_hidden_signatures_are_not_made_visible_by_baking(self):
        with fitz.open(self.source) as doc:
            for widget in doc[0].widgets():
                doc.xref_set_key(widget.xref,'F','2')
            data = doc.tobytes()
        self.source.write_bytes(data)
        self.assertEqual(bake_digital_signatures(str(self.source),str(self.output)),2)
        with fitz.open(self.output) as doc:
            self.assertEqual(len(doc[0].get_images()),0)

    def test_partial_save_failure_preserves_previous_output(self):
        def fail_save(doc,path,*args,**kwargs):
            Path(path).write_bytes(b'partial output')
            raise OSError('save failed')
        with patch.object(fitz.Document,'save',fail_save):
            with self.assertRaises(OSError):
                bake_digital_signatures(str(self.source),str(self.output))
        self.assert_originals_preserved()
        self.assertFalse(any(path.is_dir() for path in self.root.iterdir()))

    def test_successful_in_place_bake_has_no_source_handle_lock(self):
        self.assertEqual(bake_digital_signatures(str(self.source),str(self.source)),2)
        with fitz.open(self.source) as doc:
            self.assertEqual(len(list(doc[0].widgets())),0)
            self.assertIn('Keep original content',doc[0].get_text())
        self.assertFalse(any(path.is_dir() for path in self.root.iterdir()))

    def test_fallback_render_on_rotated_page_preserves_visible_appearance(self):
        with fitz.open(self.source) as doc:
            doc[0].set_rotation(90)
            data = doc.tobytes()
        self.source.write_bytes(data)
        with fitz.open(self.source) as doc:
            before = doc[0].get_pixmap(dpi=144).samples
        with patch.object(fitz.Annot,'get_pixmap',side_effect=RuntimeError('primary render unavailable')):
            self.assertEqual(bake_digital_signatures(str(self.source),str(self.output),dpi=144),2)
        with fitz.open(self.output) as doc:
            self.assertEqual(doc[0].rotation,90)
            after = doc[0].get_pixmap(dpi=144).samples
            self.assertEqual(len(before),len(after))
            self.assertLess(sum(abs(a-b) for a,b in zip(before,after))/len(before),1)


if __name__ == '__main__':
    unittest.main()
