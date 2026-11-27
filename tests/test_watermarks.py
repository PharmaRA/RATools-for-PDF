"""Real PDF regression for page-local, explicitly selected watermark removal."""

from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

import fitz

from ratools_pdf.pdf import watermarks as wm


class WatermarkTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.source = self.root / 'source.pdf'
        self.output = self.root / 'output.pdf'

    def document(self, content=None):
        doc = fitz.open()
        self.addCleanup(doc.close)
        page = doc.new_page(width=300, height=220)
        if content is not None:
            page.insert_font(fontname='helv')
            xref = doc.get_new_xref()
            doc.update_object(xref, '<<>>')
            doc.update_stream(xref, content)
            page.set_contents(xref)
        return doc

    def analyze(self, doc, pages=''):
        doc.save(self.source)
        return wm.analyze(self.source, pages)

    def export(self, analysis, candidates):
        wm.export(analysis, [item.id for item in candidates], self.output)
        output = fitz.open(self.output)
        self.addCleanup(output.close)
        return output

    def test_overlapping_text_removal_preserves_body_and_preview(self):
        doc = self.document()
        doc[0].insert_text((30, 90), 'BODY TEXT', fontsize=18)
        doc[0].insert_text((30, 90), 'DRAFT', fontsize=28, color=(.7, .7, .7), fill_opacity=.4)
        original = doc[0].get_pixmap().samples
        analysis = self.analyze(doc)
        candidates = [item for item in analysis.candidates if item.kind == '文字']
        self.assertEqual(len(candidates), 2)
        before, after = wm.preview(analysis, [candidates[-1].id], 0)
        self.assertNotEqual(before, after)
        result = self.export(analysis, candidates[-1:])
        self.assertIn('BODY TEXT', result[0].get_text())
        self.assertNotIn('DRAFT', result[0].get_text())
        self.assertEqual(wm._pixmap(result[0]).tobytes('png'), after)
        self.assertEqual(doc[0].get_pixmap().samples, original)

    def test_shared_content_stream_is_copied_for_selected_page(self):
        doc = self.document(b'BT /helv 14 Tf 1 0 0 1 30 140 Tm (DRAFT) Tj ET')
        content = doc[0].get_contents()[0]
        doc.new_page(width=300, height=220)
        doc[1].insert_font(fontname='helv')
        doc[1].set_contents(content)
        untouched = doc[1].get_pixmap().samples
        analysis = self.analyze(doc, '1')
        result = self.export(analysis, analysis.candidates)
        self.assertNotIn('DRAFT', result[0].get_text())
        self.assertEqual(result[1].get_pixmap().samples, untouched)

    def test_shared_image_only_selected_occurrence_removed(self):
        doc = self.document()
        pix = fitz.Pixmap(fitz.csRGB, fitz.IRect(0, 0, 30, 30), False)
        pix.clear_with(100)
        xref = doc[0].insert_image(fitz.Rect(20, 20, 80, 80), pixmap=pix)
        doc[0].insert_image(fitz.Rect(140, 20, 200, 80), xref=xref)
        doc.new_page(width=300, height=220)
        doc[1].insert_image(fitz.Rect(20, 20, 80, 80), xref=xref)
        original_second = doc[1].get_pixmap().samples
        analysis = self.analyze(doc, '1')
        images = [item for item in analysis.candidates if item.kind == '图片']
        self.assertEqual(len(images), 2)
        result = self.export(analysis, images[:1])
        self.assertEqual(len(result[0].get_image_info()), 1)
        self.assertEqual(result[1].get_pixmap().samples, original_second)

    def test_shared_form_is_removed_by_invocation(self):
        layer = self.document()
        layer[0].draw_rect(fitz.Rect(20, 20, 100, 100), fill=(1, 0, 0))
        doc = self.document()
        doc[0].insert_text((30, 180), 'BODY')
        doc[0].show_pdf_page(fitz.Rect(0, 0, 130, 100), layer, 0)
        doc.new_page(width=300, height=220)
        doc[1].show_pdf_page(fitz.Rect(0, 0, 130, 100), layer, 0)
        untouched = doc[1].get_pixmap().samples
        analysis = self.analyze(doc, '1')
        forms = [item for item in analysis.candidates if item.kind == '绘图／组合对象']
        self.assertEqual(len(forms), 1)
        result = self.export(analysis, forms)
        self.assertIn('BODY', result[0].get_text())
        self.assertEqual(result[1].get_pixmap().samples, untouched)
        self.assertNotEqual(result[0].get_pixmap().samples, doc[0].get_pixmap().samples)

    def test_vector_path_removal_retains_other_paths_and_text(self):
        doc = self.document()
        doc[0].draw_rect(fitz.Rect(10, 10, 80, 80), fill=(1, 0, 0))
        doc[0].draw_rect(fitz.Rect(120, 10, 190, 80), fill=(0, 0, 1))
        doc[0].insert_text((30, 160), 'BODY')
        analysis = self.analyze(doc)
        paths = [item for item in analysis.candidates if item.kind == '绘图']
        result = self.export(analysis, paths[:1])
        self.assertEqual(len(result[0].get_drawings()), 1)
        self.assertEqual(result[0].get_drawings()[0]['fill'], (0, 0, 1))
        self.assertIn('BODY', result[0].get_text())

    def test_annotation_selection_preserves_other_annotations_and_links(self):
        doc = self.document()
        page = doc[0]
        page.insert_text((20, 80), 'BODY')
        stamp = page.add_stamp_annot(fitz.Rect(10, 20, 130, 60))
        page.add_text_annot((150, 100), 'Keep this note')
        page.insert_link({'kind': fitz.LINK_URI, 'from': fitz.Rect(10, 70, 100, 90), 'uri': 'https://example.invalid'})
        analysis = self.analyze(doc)
        chosen = [item for item in analysis.candidates if stamp.xref in item.annot_xrefs]
        self.assertEqual(len(chosen), 1)
        result = self.export(analysis, chosen)
        result_page = result[0]
        self.assertEqual([item.type[1] for item in result_page.annots()], ['Text'])
        self.assertEqual(len(result_page.get_links()), 1)
        self.assertIn('BODY', result_page.get_text())

    def test_multi_show_text_and_clipping_text_are_skipped(self):
        for content in (
            b'BT /helv 12 Tf 30 130 Td (DRAFT) Tj (BODY) Tj ET',
            b'BT /helv 12 Tf 7 Tr 30 130 Td (CLIP) Tj ET',
            b'BT /helv 12 Tf 3 Tr 30 130 Td (OCR) Tj ET',
        ):
            with self.subTest(content=content):
                doc = self.document(content)
                candidates, notes = wm._painting_candidates(doc, doc[0])
                self.assertEqual(candidates, [])
                self.assertTrue(notes)

    def test_inherited_text_state_retained_after_removal(self):
        doc = self.document(b'0 0 1 rg BT /helv 14 Tf 30 160 Td (DRAFT) Tj ET\n'
                            b'BT 30 100 Td (BODY) Tj ET')
        body = doc[0].get_texttrace()[1]
        analysis = self.analyze(doc)
        result = self.export(analysis, analysis.candidates[:1])
        trace = result[0].get_texttrace()[0]
        for key in ('bbox', 'color', 'font', 'size', 'chars'):
            self.assertEqual(trace[key], body[key])

    def test_inline_image_and_unknown_operator_reject_whole_stream(self):
        for content in (b'q BI /W 1 /H 1 /BPC 8 /CS /RGB ID xxx EI Q',
                        b'BT /helv 12 Tf (DRAFT) Tj ET unknown',
                        b'BT /helv 12 Tf (DRAFT) Tj ET 123',
                        b'q BT /helv 12 Tf (DRAFT) Tj ET'):
            with self.subTest(content=content):
                doc = self.document(content)
                with self.assertRaises(ValueError):
                    wm._painting_candidates(doc, doc[0])

    def test_full_page_scan_is_not_offered_as_watermark(self):
        doc = self.document()
        pix = fitz.Pixmap(fitz.csRGB, fitz.IRect(0, 0, 300, 220), False)
        pix.clear_with(100)
        doc[0].insert_image(doc[0].rect, pixmap=pix)
        analysis = self.analyze(doc)
        self.assertFalse(analysis.candidates)
        self.assertTrue(analysis.notes)

    def test_page_ranges_are_strict_and_not_broadened(self):
        self.assertEqual(wm.parse_pages('1-3,5,3', 5), (0, 1, 2, 4))
        self.assertEqual(wm.parse_pages('', 2), (0, 1))
        for expression in ('0', '3-1', '6', '1,,2', 'one', '-1', '1-'):
            with self.assertRaises(wm.WatermarkError):
                wm.parse_pages(expression, 5)

    def test_existing_outputs_and_source_are_preserved(self):
        doc = self.document()
        doc[0].insert_text((30, 80), 'DRAFT')
        analysis = self.analyze(doc)
        original = self.source.read_bytes()
        self.output.write_bytes(b'existing success')
        for target in (self.source, self.output):
            with self.assertRaises(wm.WatermarkError):
                wm.export(analysis, [analysis.candidates[0].id], target)
        self.assertEqual(self.source.read_bytes(), original)
        self.assertEqual(self.output.read_bytes(), b'existing success')

    def test_changed_input_and_unknown_selection_are_rejected(self):
        doc = self.document()
        doc[0].insert_text((30, 80), 'DRAFT')
        analysis = self.analyze(doc)
        with self.assertRaises(wm.WatermarkError):
            wm.export(analysis, ['invalid-id'], self.output)
        self.source.write_bytes(analysis.source + b'\n% changed')
        with self.assertRaises(wm.WatermarkError):
            wm.export(analysis, [analysis.candidates[0].id], self.output)
        self.assertFalse(self.output.exists())

    def test_cancellation_does_not_publish(self):
        doc = self.document()
        doc[0].insert_text((30, 80), 'DRAFT')
        analysis = self.analyze(doc)
        with self.assertRaises(wm.WatermarkError):
            wm.export(analysis, [analysis.candidates[0].id], self.output, cancel=lambda: True)
        self.assertFalse(self.output.exists())

    def test_failed_publication_recycles_staging_and_preserves_input(self):
        doc = self.document()
        doc[0].insert_text((30, 80), 'DRAFT')
        analysis = self.analyze(doc)
        with patch.object(wm.os, 'rename', side_effect=OSError('locked')):
            with self.assertRaises(OSError):
                wm.export(analysis, [analysis.candidates[0].id], self.output)
        self.assertEqual(list(self.root.glob('.ratools-watermark-*')), [])
        self.assertEqual(self.source.read_bytes(), analysis.source)

    def test_staging_collision_preserves_file_not_owned_by_task(self):
        doc = self.document()
        doc[0].insert_text((30, 80), 'DRAFT')
        analysis = self.analyze(doc)
        collision = self.root / '.ratools-watermark-collision.pdf'
        collision.write_bytes(b'not ours')
        with patch.object(wm.uuid, 'uuid4') as uuid:
            uuid.return_value.hex = 'collision'
            with self.assertRaises(FileExistsError):
                wm.export(analysis, [analysis.candidates[0].id], self.output)
        self.assertEqual(collision.read_bytes(), b'not ours')

    def test_destination_created_during_processing_is_not_overwritten(self):
        doc = self.document()
        doc[0].insert_text((30, 80), 'DRAFT')
        analysis = self.analyze(doc)
        rename = wm.os.rename

        def race(source, target):
            self.output.write_bytes(b'another writer')
            return rename(source, target)

        with patch.object(wm.os, 'rename', side_effect=race):
            with self.assertRaises(FileExistsError):
                wm.export(analysis, [analysis.candidates[0].id], self.output)
        self.assertEqual(self.output.read_bytes(), b'another writer')
        self.assertEqual(list(self.root.glob('.ratools-watermark-*')), [])

    def test_shared_annotation_array_is_only_changed_on_selected_page(self):
        doc = self.document()
        page = doc[0]
        page.add_stamp_annot(fitz.Rect(10, 20, 130, 60))
        _kind, refs = doc.xref_get_key(page.xref, 'Annots')
        array = doc.get_new_xref()
        doc.update_object(array, refs)
        doc.xref_set_key(page.xref, 'Annots', f'{array} 0 R')
        second = doc.new_page(width=300, height=220)
        doc.xref_set_key(second.xref, 'Annots', f'{array} 0 R')
        analysis = self.analyze(doc, '1')
        with fitz.open(stream=analysis.source, filetype='pdf') as baseline:
            original = baseline[1].get_pixmap().samples
        result = self.export(analysis, analysis.candidates)
        self.assertEqual(result[1].get_pixmap().samples, original)
        first = result[0]
        self.assertEqual(list(first.annots() or ()), [])

    def test_rotated_page_retains_geometry_and_body(self):
        doc = self.document()
        doc[0].insert_text((30, 80), 'BODY')
        doc[0].insert_text((30, 120), 'DRAFT')
        doc[0].set_rotation(90)
        doc[0].set_cropbox(fitz.Rect(10, 10, 280, 210))
        analysis = self.analyze(doc)
        result = self.export(analysis, analysis.candidates[-1:])
        self.assertEqual(result[0].rect, doc[0].rect)
        self.assertEqual(result[0].rotation, 90)
        self.assertEqual(result[0].cropbox, doc[0].cropbox)
        self.assertEqual(result[0].get_text().strip(), 'BODY')

    def test_cancel_during_final_input_check_does_not_publish(self):
        doc = self.document()
        doc[0].insert_text((30, 80), 'DRAFT')
        analysis = self.analyze(doc)
        original_read = Path.read_bytes
        reads = []

        def read(path):
            result = original_read(path)
            if path == self.source:
                reads.append(path)
            return result

        with patch.object(Path, 'read_bytes', read):
            with self.assertRaises(wm.WatermarkError):
                wm.export(analysis, [analysis.candidates[0].id], self.output,
                          cancel=lambda: len(reads) >= 2)
        self.assertFalse(self.output.exists())
        self.assertEqual(list(self.root.glob('.ratools-watermark-*')), [])

    def test_recycle_failure_reports_retained_staging_file(self):
        doc = self.document()
        doc[0].insert_text((30, 80), 'DRAFT')
        analysis = self.analyze(doc)
        with patch.object(wm.os, 'rename', side_effect=OSError('locked')), \
                patch.object(wm, 'recycle_owned_path', side_effect=OSError('recycle unavailable')):
            with self.assertRaisesRegex(wm.WatermarkError, '临时文件无法移入回收站'):
                wm.export(analysis, [analysis.candidates[0].id], self.output)
        self.assertFalse(self.output.exists())
        self.assertEqual(len(list(self.root.glob('.ratools-watermark-*'))), 1)


if __name__ == '__main__':
    unittest.main()
