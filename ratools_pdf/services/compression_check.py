"""A headless check of the actual image codec and PDF pipeline in a release."""

import json
from io import BytesIO
from pathlib import Path

from ratools_pdf.services.file_safety import check_unlinked_path


def verify_image_compression(output_directory):
    root = check_unlinked_path(output_directory)
    root.mkdir(parents=True, exist_ok=False)
    report = {'ok': False}
    try:
        import fitz
        from PIL import Image, ImageDraw
        from ratools_pdf.pdf.processor import process_document
        image = Image.new('RGB',(1800,2200))
        draw = ImageDraw.Draw(image)
        for y in range(2200):
            draw.line((0,y,1800,y),fill=(y%256,(y*2)%256,220-y%220))
        data = BytesIO()
        image.save(data,format='JPEG',quality=95)
        source, output = root/'source.pdf', root/'compressed.pdf'
        with fitz.open() as doc:
            page=doc.new_page(width=432,height=576)
            page.insert_image(page.rect,stream=data.getvalue())
            doc.save(source)
        ok, message = process_document(str(source),str(output),{'compress_images'},
                                       processing_mode='force',compression_settings={'dpi':72,'quality':70})
        if not ok:
            raise RuntimeError(message)
        with fitz.open(output) as doc:
            info = doc[0].get_images(full=True)[0]
            pixmap = doc[0].get_pixmap()
            if info[2] >= 1800 or info[3] >= 2200 or not pixmap.samples:
                raise RuntimeError('Packaged image compression did not downsample/render')
            if output.stat().st_size >= source.stat().st_size:
                raise RuntimeError('Packaged compression did not reduce the PDF size')
            report.update(ok=True, image_width=info[2], image_height=info[3],
                          input_bytes=source.stat().st_size, output_bytes=output.stat().st_size)
    except Exception as exc:
        report['error'] = str(exc)
    (root/'verification.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
    return 0 if report['ok'] else 1
