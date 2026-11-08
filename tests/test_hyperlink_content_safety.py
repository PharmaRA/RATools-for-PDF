"""Changing hyperlink color must preserve text and unrelated graphics state."""

import unittest
import tempfile
from pathlib import Path

import fitz

from ratools_pdf.pdf.hyperlink_styles import _apply_text_color_via_content_stream
from ratools_pdf.pdf.processor import process_document


class HyperlinkContentSafetyTests(unittest.TestCase):
    def document(self, content):
        doc = fitz.open()
        self.addCleanup(doc.close)
        page = doc.new_page(width=300, height=200)
        page.insert_font(fontname='helv')
        xref = doc.get_new_xref()
        doc.update_object(xref, '<<>>')
        doc.update_stream(xref, content)
        page.set_contents(xref)
        return doc, page

    def assert_recolor_preserves_text(self, content):
        doc, page = self.document(content)
        before = page.get_text('words')
        self.assertTrue(before, 'fixture must contain visible text')
        self.assertTrue(_apply_text_color_via_content_stream(doc, page, [page.rect], (0, 0, 1)))
        self.assertEqual(page.get_text('words'), before)
        self.assertTrue(all(trace['color'] == (0.0, 0.0, 1.0) for trace in page.get_texttrace()))
        with fitz.open(stream=doc.tobytes(), filetype='pdf') as reopened:
            self.assertEqual(reopened[0].get_text('words'), before)
            pixels = reopened[0].get_pixmap().samples
            self.assertTrue(any(pixels[i + 2] > pixels[i] for i in range(0, len(pixels), 3)),
                            'render must contain visible blue text')

    def test_default_color_tj_preserves_text(self):
        self.assert_recolor_preserves_text(b'BT /helv 12 Tf 1 0 0 1 20 150 Tm (LINK) Tj ET')

    def test_default_color_tj_array_preserves_text_and_spacing(self):
        self.assert_recolor_preserves_text(b'BT /helv 12 Tf 1 0 0 1 20 150 Tm [(LINK) -100 (TEXT)] TJ ET')

    def test_inherited_color_preserves_text(self):
        self.assert_recolor_preserves_text(b'0 1 0 rg BT /helv 12 Tf 1 0 0 1 20 150 Tm (LINK) Tj ET')

    def test_escaped_parentheses_preserve_literal_text(self):
        self.assert_recolor_preserves_text(rb'BT /helv 12 Tf 1 0 0 1 20 150 Tm (LINK \(A\)) Tj ET')

    def test_operator_words_inside_literal_strings_are_not_commands(self):
        self.assert_recolor_preserves_text(b'BT /helv 12 Tf 1 0 0 1 20 150 Tm (TARGET BT ET 0 0 0 rg) Tj ET')

    def test_operator_words_in_comments_do_not_terminate_text_block(self):
        self.assert_recolor_preserves_text(b'BT /helv 12 Tf % ET 0 0 0 rg\n1 0 0 1 20 150 Tm (LINK) Tj ET')

    def test_recolor_does_not_change_following_text_or_inherited_font(self):
        doc, page = self.document(
            b'BT /helv 12 Tf 1 0 0 1 20 150 Tm (LINK) Tj ET\n'
            b'BT 1 0 0 1 20 100 Tm (BODY) Tj ET'
        )
        before = page.get_text('words')
        target = fitz.Rect(page.get_texttrace()[0]['bbox'])
        self.assertTrue(_apply_text_color_via_content_stream(doc, page, [target], (0, 0, 1)))
        self.assertEqual(page.get_text('words'), before)
        colors = [trace['color'] for trace in page.get_texttrace()]
        self.assertEqual(colors, [(0.0, 0.0, 1.0), (0.0, 0.0, 0.0)])

    def test_explicit_gray_color_preserves_text(self):
        self.assert_recolor_preserves_text(b'BT /helv 12 Tf 0.4 g 1 0 0 1 20 150 Tm (LINK) Tj ET')

    def test_hex_text_preserves_characters(self):
        self.assert_recolor_preserves_text(b'BT /helv 12 Tf 1 0 0 1 20 150 Tm <4c494e4b> Tj ET')

    def test_color_inherited_across_content_streams_is_restored(self):
        doc, page = self.document(b'0 1 0 rg')
        first = page.get_contents()[0]
        second = doc.get_new_xref()
        doc.update_object(second, '<<>>')
        doc.update_stream(second,
                          b'BT /helv 12 Tf 1 0 0 1 20 150 Tm (LINK) Tj ET\n'
                          b'BT 1 0 0 1 20 100 Tm (BODY) Tj ET')
        doc.xref_set_key(page.xref, 'Contents', f'[{first} 0 R {second} 0 R]')
        page = doc.reload_page(page)
        before = page.get_text('words')
        target = fitz.Rect(page.get_texttrace()[0]['bbox'])
        self.assertTrue(_apply_text_color_via_content_stream(doc, page, [target], (0, 0, 1)))
        self.assertEqual(page.get_text('words'), before)
        self.assertEqual([trace['color'] for trace in page.get_texttrace()],
                         [(0.0, 0.0, 1.0), (0.0, 1.0, 0.0)])

    def test_processing_rule_preserves_text_and_links_in_saved_pdf(self):
        doc, page = self.document(b'BT /helv 12 Tf 1 0 0 1 20 150 Tm (LINK) Tj ET')
        before = page.get_text('words')
        page.insert_link({'kind': fitz.LINK_URI, 'from': fitz.Rect(15, 30, 80, 60),
                          'uri': 'https://example.invalid/'})
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / 'source.pdf'
            output = Path(directory) / 'output.pdf'
            doc.save(source)
            ok, message = process_document(str(source), str(output), {'link_text_blue'}, processing_mode='force')
            self.assertTrue(ok, message)
            with fitz.open(output) as result:
                self.assertEqual(result[0].get_text('words'), before)
                self.assertEqual(result[0].get_texttrace()[0]['color'], (0.0, 0.0, 1.0))
                self.assertEqual(result[0].get_links()[0]['uri'], 'https://example.invalid/')


if __name__ == '__main__':
    unittest.main()
