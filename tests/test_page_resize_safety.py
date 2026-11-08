"""Paper normalization preserves rotated content and page-local objects."""

import unittest
import gc
import os
import tempfile
from pathlib import Path

import fitz

from ratools_pdf.pdf.page_layout import resize_pages_with_padding
from ratools_pdf.pdf.processor import process_document


class PageResizeSafetyTests(unittest.TestCase):
    def make_document(self, rotation=0, objects=False):
        doc = fitz.open()
        page = doc.new_page(width=200, height=400)
        page.insert_text((20, 80), 'MARKER', fontsize=14)
        page.draw_rect(fitz.Rect(30, 180, 100, 260), color=(0, 0, 1), fill=(1, .5, 0))
        if objects:
            rect = page.add_rect_annot(fitz.Rect(20, 100, 80, 140))
            rect.set_colors(stroke=(1, 0, 0))
            rect.update()
            highlight = page.add_highlight_annot(fitz.Rect(20, 60, 90, 82))
            highlight.update()
            widget = fitz.Widget()
            widget.field_name = 'test-field'
            widget.field_type = fitz.PDF_WIDGET_TYPE_TEXT
            widget.field_value = 'VALUE'
            widget.rect = fitz.Rect(20, 290, 150, 320)
            page.add_widget(widget)
            page.insert_link({'kind': fitz.LINK_URI, 'from': fitz.Rect(20, 60, 90, 82),
                              'uri': 'https://example.invalid/'})
        page.set_rotation(rotation)
        data = doc.tobytes()
        doc.close()
        doc = fitz.open(stream=data, filetype='pdf')
        self.addCleanup(doc.close)
        return doc

    def assert_rect_close(self, actual, expected):
        for value, wanted in zip(actual, expected):
            self.assertAlmostEqual(value, wanted, delta=.05)

    def test_rotations_preserve_direction_text_and_rendering(self):
        for rotation in (0, 90, 180, 270):
            with self.subTest(rotation=rotation):
                doc = self.make_document(rotation)
                page = doc[0]
                source = page.get_pixmap(matrix=fitz.Matrix(2, 2), annots=False)
                visible_word = fitz.Rect(page.get_text('words')[0][:4]) * page.rotation_matrix
                if rotation in (90, 270):
                    expected_rect = fitz.Rect(0, 0, 800, 600)
                    padding = fitz.Point(0, 100)
                    content_rect = fitz.Rect(0, 100, 800, 500)
                else:
                    expected_rect = fitz.Rect(0, 0, 600, 800)
                    padding = fitz.Point(100, 0)
                    content_rect = fitz.Rect(100, 0, 500, 800)
                expected_word = fitz.Rect(*(value * 2 + (padding.x if i % 2 == 0 else padding.y)
                                            for i, value in enumerate(visible_word)))
                self.assertEqual(resize_pages_with_padding(doc, fitz.Rect(0, 0, 600, 800)), 1)
                with fitz.open(stream=doc.tobytes(), filetype='pdf') as result:
                    page = result[0]
                    self.assert_rect_close(page.rect, expected_rect)
                    self.assertEqual(page.get_text().strip(), 'MARKER')
                    actual_word = fitz.Rect(page.get_text('words')[0][:4]) * page.rotation_matrix
                    self.assert_rect_close(actual_word, expected_word)
                    rendered = page.get_pixmap(clip=content_rect, annots=False)
                    self.assertEqual((rendered.width, rendered.height), (source.width, source.height))
                    mean_difference = sum(abs(a-b) for a,b in zip(rendered.samples, source.samples)) / len(source.samples)
                    self.assertLess(mean_difference, .5)

    def test_annotations_widgets_and_links_follow_content_for_every_rotation(self):
        for rotation in (0, 90, 180, 270):
            with self.subTest(rotation=rotation):
                doc = self.make_document(rotation, objects=True)
                page = doc[0]
                old_annots = [(tuple(a.rect), a.vertices) for a in page.annots()]
                old_widget = tuple(next(page.widgets()).rect)
                old_link = fitz.Rect(page.get_links()[0]['from'])
                original_pixels = page.get_pixmap(matrix=fitz.Matrix(2,2)).samples
                resize_pages_with_padding(doc, fitz.Rect(0, 0, 600, 800))
                with fitz.open(stream=doc.tobytes(), filetype='pdf') as result:
                    page = result[0]
                    for annot, (old_rect, vertices) in zip(page.annots(), old_annots):
                        self.assert_rect_close(annot.rect, tuple(v*2+(100 if i%2==0 else 0) for i,v in enumerate(old_rect)))
                        if vertices:
                            for actual, original in zip(annot.vertices, vertices):
                                self.assertAlmostEqual(actual[0], original[0]*2+100, delta=.05)
                                self.assertAlmostEqual(actual[1], original[1]*2, delta=.05)
                    self.assert_rect_close(next(page.widgets()).rect,
                                           tuple(v*2+(100 if i%2==0 else 0) for i,v in enumerate(old_widget)))
                    dx, dy = (0,100) if rotation in (90,270) else (100,0)
                    self.assert_rect_close(page.get_links()[0]['from'],
                                           tuple(v*2+(dx if i%2==0 else dy) for i,v in enumerate(old_link)))
                    self.assertEqual(page.get_links()[0]['uri'], 'https://example.invalid/')
                    clip = fitz.Rect(0,100,800,500) if rotation in (90,270) else fitz.Rect(100,0,500,800)
                    new_pixels = page.get_pixmap(clip=clip).samples
                    self.assertEqual(len(new_pixels), len(original_pixels))
                    difference = sum(abs(a-b) for a,b in zip(new_pixels,original_pixels))/len(original_pixels)
                    self.assertLess(difference,.8)

    def test_inline_link_is_scaled_without_rebuilding_its_action(self):
        doc = self.make_document(90)
        page = doc[0]
        doc.xref_set_key(page.xref, 'Annots',
                        '[<< /Type /Annot /Subtype /Link /Rect [20 250 80 280] '
                        '/A << /S /URI /URI (https://example.invalid/) >> >>]')
        page = doc.reload_page(page)
        before = tuple(page.get_links()[0]['from'])
        resize_pages_with_padding(doc, fitz.Rect(0,0,600,800))
        with fitz.open(stream=doc.tobytes(),filetype='pdf') as result:
            link = result[0].get_links()[0]
            self.assert_rect_close(link['from'], tuple(v*2+(0 if i%2==0 else 100) for i,v in enumerate(before)))
            self.assertEqual(link['uri'], 'https://example.invalid/')

    def test_font_state_shared_between_content_streams_is_preserved(self):
        doc = fitz.open()
        self.addCleanup(doc.close)
        page = doc.new_page(width=200,height=400)
        page.insert_font(fontname='helv')
        streams = []
        for content in (b'BT /helv 12 Tf 1 0 0 1 20 300 Tm (FIRST) Tj ET',
                        b'BT 1 0 0 1 20 200 Tm (SECOND) Tj ET'):
            xref = doc.get_new_xref()
            doc.update_object(xref,'<<>>')
            doc.update_stream(xref,content)
            streams.append(xref)
        doc.xref_set_key(page.xref,'Contents','['+' '.join(f'{xref} 0 R' for xref in streams)+']')
        page.set_rotation(90)
        before = page.get_text()
        resize_pages_with_padding(doc,fitz.Rect(0,0,600,800))
        self.assertEqual(doc[0].get_text(),before)

    def test_crop_offset_and_padding_preserve_visible_rotated_content(self):
        doc = fitz.open()
        self.addCleanup(doc.close)
        page = doc.new_page(width=300,height=500)
        page.insert_text((80,120),'MARKER')
        page.draw_rect(fitz.Rect(0,0,45,500),color=None,fill=(1,0,0))
        page.set_cropbox(fitz.Rect(50,50,250,450))
        page.set_rotation(90)
        original = page.get_pixmap(matrix=fitz.Matrix(2,2))
        resize_pages_with_padding(doc,fitz.Rect(0,0,600,800))
        with fitz.open(stream=doc.tobytes(),filetype='pdf') as result:
            page = result[0]
            self.assertEqual(tuple(page.rect),(0,0,800,600))
            pixels = page.get_pixmap(clip=fitz.Rect(0,100,800,500)).samples
            self.assertEqual(len(pixels),len(original.samples))
            self.assertLess(sum(abs(a-b) for a,b in zip(pixels,original.samples))/len(pixels),.5)
            self.assertTrue(all(value==255 for value in page.get_pixmap(clip=fitz.Rect(0,0,800,99)).samples))

    def test_resizing_does_not_modify_content_shared_with_an_unchanged_page(self):
        doc = self.make_document()
        source_page = doc[0]
        contents = doc.xref_get_key(source_page.xref,'Contents')[1]
        resources = doc.xref_get_key(source_page.xref,'Resources')[1]
        other = doc.new_page(width=600,height=800)
        doc.xref_set_key(other.xref,'Contents',contents)
        doc.xref_set_key(other.xref,'Resources',resources)
        before = other.get_pixmap().samples
        self.assertEqual(resize_pages_with_padding(doc,fitz.Rect(0,0,600,800)),1)
        self.assertEqual(doc[1].get_pixmap().samples,before)

    def test_a4_pipeline_saves_rotated_content_and_objects(self):
        doc = self.make_document(90,objects=True)
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory)/'source.pdf'
            output = Path(directory)/'output.pdf'
            doc.save(source)
            ok,message = process_document(str(source),str(output),{'page_size_a4'},processing_mode='force')
            self.assertTrue(ok,message)
            with fitz.open(output) as result:
                page = result[0]
                self.assertAlmostEqual(page.rect.width,297/25.4*72,delta=.01)
                self.assertAlmostEqual(page.rect.height,210/25.4*72,delta=.01)
                self.assertIn('MARKER',page.get_text())
                self.assertEqual(len(list(page.annots())),2)
                self.assertEqual(len(page.get_links()),1)
                self.assertEqual(next(page.widgets()).field_value,'VALUE')

    def test_shared_border_style_is_copied_per_resized_annotation(self):
        doc = self.make_document()
        page = doc[0]
        resized_xrefs = []
        for y in (100,200):
            annotation = page.add_rect_annot(fitz.Rect(20,y,80,y+30))
            annotation.update()
            resized_xrefs.append(annotation.xref)
        other = doc.new_page(width=600,height=800)
        annotation = other.add_rect_annot(fitz.Rect(20,100,80,130))
        annotation.update()
        unchanged_xref = annotation.xref
        border = doc.get_new_xref()
        doc.update_object(border,'<< /W 2 /S /D /D [3 2] >>')
        for xref in [*resized_xrefs,unchanged_xref]:
            doc.xref_set_key(xref,'BS',f'{border} 0 R')
        resize_pages_with_padding(doc,fitz.Rect(0,0,600,800))
        for xref in resized_xrefs:
            self.assertEqual(float(doc.xref_get_key(xref,'BS/W')[1]),4)
            self.assertEqual(doc.xref_get_key(xref,'BS/D')[1],'[6 4]')
        self.assertEqual(float(doc.xref_get_key(unchanged_xref,'BS/W')[1]),2)
        self.assertEqual(doc.xref_get_key(border,'D')[1],'[3 2]')

    @unittest.skipUnless(os.name == 'nt','Windows exclusive file handles')
    def test_pipeline_releases_source_handle_without_waiting_for_cycle_collection(self):
        import ctypes
        from ctypes import wintypes
        kernel = ctypes.WinDLL('kernel32',use_last_error=True)
        kernel.CreateFileW.argtypes = (wintypes.LPCWSTR,wintypes.DWORD,wintypes.DWORD,
                                       wintypes.LPVOID,wintypes.DWORD,wintypes.DWORD,wintypes.HANDLE)
        kernel.CreateFileW.restype = wintypes.HANDLE
        kernel.CloseHandle.argtypes = (wintypes.HANDLE,)
        doc = self.make_document()
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory)/'source.pdf'
            output = Path(directory)/'output.pdf'
            doc.save(source)
            was_enabled = gc.isenabled()
            gc.disable()
            try:
                ok,message = process_document(str(source),str(output),{'page_size_a4'},processing_mode='force')
                self.assertTrue(ok,message)
                handle = kernel.CreateFileW(str(source),0x80000000,0,None,3,0,None)
                acquired = handle != ctypes.c_void_p(-1).value
                if acquired:
                    kernel.CloseHandle(handle)
                self.assertTrue(acquired,'closed PDF still holds the source file open')
            finally:
                if was_enabled:
                    gc.enable()
                gc.collect()


if __name__ == '__main__':
    unittest.main()
