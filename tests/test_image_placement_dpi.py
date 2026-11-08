"""Downsampling uses the largest drawn size of each shared image object."""

from io import BytesIO
import tempfile
import unittest
from pathlib import Path

import fitz
from PIL import Image, ImageDraw

from ratools_pdf.pdf.processor import process_document


class ImagePlacementDpiTests(unittest.TestCase):
    def setUp(self):
        self.directory=tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.source=Path(self.directory.name)/'source.pdf'
        self.output=Path(self.directory.name)/'out.pdf'

    def image(self, width=1200, height=1200):
        image=Image.new('RGB',(width,height))
        draw=ImageDraw.Draw(image)
        for y in range(height):
            draw.line((0,y,width,y), fill=(y%256,(y*2)%256,100))
        stream=BytesIO()
        image.save(stream,format='JPEG',quality=95)
        return stream.getvalue()

    def process(self, doc, dpi, expected):
        doc.save(self.source)
        doc.close()
        ok,message=process_document(str(self.source),str(self.output),{'compress_images'},
                                    processing_mode='force',compression_settings={'dpi':dpi,'quality':85})
        self.assertTrue(ok,message)
        with fitz.open(self.output) as restored:
            sizes={(info[2],info[3]) for page in restored for info in page.get_images(full=True)}
            self.assertEqual(len(sizes),1)
            width,height=sizes.pop()
            self.assertAlmostEqual(width,expected[0],delta=1)
            self.assertAlmostEqual(height,expected[1],delta=1)
            for page in restored:
                self.assertGreater(len(page.get_pixmap(matrix=fitz.Matrix(.5,.5)).samples),0)

    def test_one_inch_picture_is_reduced_to_requested_300_dpi(self):
        doc=fitz.open()
        page=doc.new_page()
        page.insert_image(fitz.Rect(72,72,144,144),stream=self.image())
        self.process(doc,300,(300,300))

    def test_landscape_12_by_6_inches_retains_150_dpi(self):
        doc=fitz.open()
        page=doc.new_page(width=900,height=500)
        page.insert_image(fitz.Rect(0,0,864,432),stream=self.image(2400,1200))
        self.process(doc,150,(1800,900))

    def test_large_poster_uses_drawn_size_not_a4_limit(self):
        doc=fitz.open()
        page=doc.new_page(width=1440,height=1440)
        page.insert_image(page.rect,stream=self.image(2400,2400))
        self.process(doc,72,(1440,1440))

    def test_shared_image_uses_largest_placement_on_later_page(self):
        for large_first in (False,True):
            with self.subTest(large_first=large_first):
                doc=fitz.open()
                xref=None
                for inches in ((3,1) if large_first else (1,3)):
                    page=doc.new_page()
                    rect=fitz.Rect(0,0,72*inches,72*inches)
                    xref=page.insert_image(rect,stream=self.image() if xref is None else None,xref=xref or 0)
                self.process(doc,300,(900,900))

    def test_shared_image_uses_largest_occurrence_on_same_page(self):
        doc=fitz.open()
        page=doc.new_page()
        xref=page.insert_image(fitz.Rect(0,0,72,72),stream=self.image())
        page.insert_image(fitz.Rect(100,100,316,316),xref=xref)
        self.process(doc,300,(900,900))

    def test_nonuniform_stretch_keeps_both_axes_at_least_target_dpi(self):
        doc=fitz.open()
        page=doc.new_page()
        page.insert_image(fitz.Rect(0,0,144,72),stream=self.image(),keep_proportion=False)
        self.process(doc,150,(300,300))

    def test_rotated_image_uses_image_axes_not_axis_aligned_page_box(self):
        doc=fitz.open()
        page=doc.new_page()
        page.insert_image(fitz.Rect(0,0,72,144),stream=self.image(1200,600),rotate=90)
        page.set_rotation(90)
        self.process(doc,150,(300,150))

    def test_nested_form_image_placement_is_measured(self):
        with fitz.open() as template:
            page=template.new_page(width=72,height=72)
            page.insert_image(page.rect,stream=self.image())
            doc=fitz.open()
            page=doc.new_page()
            page.show_pdf_page(fitz.Rect(0,0,144,144),template,0)
        self.process(doc,150,(300,300))

    def test_image_already_below_target_dpi_is_not_downsampled(self):
        doc=fitz.open()
        page=doc.new_page(width=1440,height=1440)
        page.insert_image(page.rect,stream=self.image())
        self.process(doc,300,(1200,1200))

    def test_user_unit_is_included_in_physical_image_dimensions(self):
        doc=fitz.open()
        page=doc.new_page()
        page.insert_image(fitz.Rect(0,0,72,72),stream=self.image())
        doc.xref_set_key(page.xref,'UserUnit','2')
        self.process(doc,300,(600,600))

    def test_shear_uses_vector_lengths_instead_of_bounding_box(self):
        doc=fitz.open()
        page=doc.new_page()
        page.insert_image(fitz.Rect(0,0,72,72),stream=self.image())
        name=page.get_images(full=True)[0][7]
        doc.update_stream(page.get_contents()[0],f'q 144 0 72 72 0 0 cm /{name} Do Q'.encode())
        self.process(doc,300,(600,600))


if __name__=='__main__':
    unittest.main()
