"""Image compression preserves transparency over non-white page backgrounds."""

from io import BytesIO
import tempfile
import unittest
from pathlib import Path

import fitz
from PIL import Image, ImageDraw

from ratools_pdf.pdf.processor import _compress_single_image, process_document


class ImageTransparencyTests(unittest.TestCase):
    def document(self, kind='soft', width=600, height=700):
        doc = fitz.open()
        self.addCleanup(doc.close)
        page = doc.new_page(width=300,height=350)
        page.draw_rect(page.rect,color=None,fill=(0,0,1))
        if kind == 'soft':
            image = Image.new('RGBA',(width,height),(255,0,0,128))
            ImageDraw.Draw(image).rectangle((0,0,width//4,height),fill=(0,255,0,0))
        elif kind == 'matte':
            image = Image.new('RGB',(width,height),(255,128,128))
        elif kind == 'key':
            image = Image.new('RGB',(width,height),'white')
            ImageDraw.Draw(image).rectangle((width//2,0,width,height),fill=(255,0,0))
        else:
            image = Image.new('RGB',(width,height),(255,0,0))
        data = BytesIO()
        image.save(data,format='PNG')
        xref = page.insert_image(page.rect,stream=data.getvalue())
        if kind == 'key':
            doc.xref_set_key(xref,'Mask','[255 255 255 255 255 255]')
        elif kind == 'matte':
            mask = doc.get_new_xref()
            doc.update_object(mask,f'<< /Type /XObject /Subtype /Image /Width {width} /Height {height} '
                              '/ColorSpace /DeviceGray /BitsPerComponent 8 /Matte [1 1 1] >>')
            doc.update_stream(mask,bytes([128])*(width*height))
            doc.xref_set_key(xref,'SMask',f'{mask} 0 R')
        elif kind == 'explicit':
            mask = doc.get_new_xref()
            doc.update_object(mask,f'<< /Type /XObject /Subtype /Image /Width {width} /Height {height} '
                              '/ImageMask true /BitsPerComponent 1 /Decode [0 1] >>')
            doc.update_stream(mask,bytes([0xF0])*(((width+7)//8)*height))
            doc.xref_set_key(xref,'Mask',f'{mask} 0 R')
        # A deliberately uncompressed stream guarantees downsampling can save space.
        raw = doc.xref_stream(xref)
        doc.update_stream(xref,raw,compress=0)
        return doc,xref

    def check_compression(self, kind):
        doc,xref = self.document(kind)
        before = doc[0].get_pixmap().samples
        self.assertTrue(_compress_single_image(doc,xref,300,350,90), 'fixture should be compressed')
        self.assertEqual(doc.xref_get_key(xref,'Width'),('int','300'))
        with fitz.open(stream=doc.tobytes(),filetype='pdf') as result:
            after = result[0].get_pixmap().samples
            self.assertEqual(len(before),len(after))
            difference = sum(abs(a-b) for a,b in zip(before,after))/len(before)
            self.assertLess(difference,3, 'transparency or background composition changed')

    def test_soft_mask_preserves_semitransparent_and_clear_regions(self):
        self.check_compression('soft')

    def test_color_key_mask_preserves_transparent_background(self):
        self.check_compression('key')

    def test_matte_soft_mask_preserves_unblended_colors(self):
        self.check_compression('matte')

    def test_explicit_stencil_mask_preserves_background(self):
        self.check_compression('explicit')

    def test_shared_mask_is_not_modified_when_one_image_is_compressed(self):
        doc,xref = self.document('soft')
        mask = int(doc.xref_get_key(xref,'SMask')[1].split()[0])
        original_mask = doc.xref_stream(mask)
        image = Image.new('RGB',(600,700),(0,255,0))
        data = BytesIO(); image.save(data,format='PNG')
        page = doc.new_page(width=300,height=350)
        page.draw_rect(page.rect,color=None,fill=(1,0,0))
        second = page.insert_image(page.rect,stream=data.getvalue())
        doc.xref_set_key(second,'SMask',f'{mask} 0 R')
        before = doc[1].get_pixmap().samples
        self.assertTrue(_compress_single_image(doc,xref,300,350,90))
        self.assertEqual(doc.xref_stream(mask),original_mask)
        self.assertEqual(doc[1].get_pixmap().samples,before)
        self.assertEqual(doc.xref_get_key(xref,'SMask')[0],'xref')
        self.assertNotEqual(doc.xref_get_key(xref,'SMask')[1],f'{mask} 0 R')

    def test_full_compression_pipeline_preserves_transparency_after_save(self):
        doc,xref = self.document('soft',width=1200,height=1400)
        before = doc[0].get_pixmap().samples
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory)/'source.pdf'
            output = Path(directory)/'output.pdf'
            doc.save(source)
            ok,message = process_document(str(source),str(output),{'compress_images'},processing_mode='force',
                                          compression_settings={'dpi':72,'quality':90})
            self.assertTrue(ok,message)
            with fitz.open(output) as result:
                xref_out = result[0].get_images()[0][0]
                self.assertLess(int(result.xref_get_key(xref_out,'Width')[1]),1200)
                after = result[0].get_pixmap().samples
                self.assertLess(sum(abs(a-b) for a,b in zip(before,after))/len(before),3)


if __name__ == '__main__':
    unittest.main()
