import fitz

from ratools_pdf.pdf.navigation_geometry import resize_local_destinations


def _transform_rect(rect, scale, dx, dy):
    return fitz.Rect(
        rect.x0 * scale + dx,
        rect.y0 * scale + dy,
        rect.x1 * scale + dx,
        rect.y1 * scale + dy,
    )


def _transform_point(point, scale, dx, dy):
    if point is None:
        return None
    return fitz.Point(point.x * scale + dx, point.y * scale + dy)


def _get_oriented_target_rect(base_target_rect, src_rect):
    if src_rect.width > src_rect.height and base_target_rect.width < base_target_rect.height:
        return fitz.Rect(0, 0, base_target_rect.height, base_target_rect.width)
    if src_rect.width < src_rect.height and base_target_rect.width > base_target_rect.height:
        return fitz.Rect(0, 0, base_target_rect.height, base_target_rect.width)
    return fitz.Rect(0, 0, base_target_rect.width, base_target_rect.height)


def _paper_rect_exact(size_name):
    size = size_name.lower()
    if size == "a4":
        return fitz.Rect(0, 0, 210 / 25.4 * 72, 297 / 25.4 * 72)
    if size == "letter":
        return fitz.Rect(0, 0, 8.5 * 72, 11 * 72)
    return fitz.paper_rect(size_name)


def _visible_pdf_box(page):
    mupdf = fitz.mupdf
    obj = fitz._as_pdf_page(page.this).obj()
    media = fitz.Rect(mupdf.pdf_to_rect(mupdf.pdf_dict_get_inheritable(obj, fitz.PDF_NAME('MediaBox'))))
    crop = mupdf.pdf_dict_get_inheritable(obj, fitz.PDF_NAME('CropBox'))
    return (fitz.Rect(mupdf.pdf_to_rect(crop)) & media) if mupdf.pdf_is_array(crop) else media


def _transform_annotation_array(pdf, array, scale, dx, dy, coordinates):
    # A self-recursive closure would retain PdfDocument in a reference cycle,
    # keeping the input file locked even after Document.close() on Windows.
    mupdf = fitz.mupdf
    count = mupdf.pdf_array_len(array)
    result = mupdf.pdf_new_array(pdf, count)
    for index in range(count):
        value = mupdf.pdf_array_get(array, index)
        if mupdf.pdf_is_array(value):
            nested = _transform_annotation_array(pdf, value, scale, dx, dy, coordinates)
            mupdf.pdf_array_push(result, nested)
        elif mupdf.pdf_is_number(value):
            offset = (dx if index % 2 == 0 else dy) if coordinates else 0
            mupdf.pdf_array_push_real(result, mupdf.pdf_to_real(value) * scale + offset)
        else:
            raise ValueError('批注几何数组包含非数字，无法安全调整页面尺寸')
    return result


def _transform_annotation_geometry(page, scale, dx, dy):
    """Transform raw PDF geometry without rebuilding actions or appearances.

    Annots can contain both references and inline dictionaries. Raw PDF coordinates
    also avoid the different rotation conventions of links, annotations and widgets.
    """
    mupdf = fitz.mupdf
    pdf = fitz._as_pdf_document(page.parent)
    obj = fitz._as_pdf_page(page.this).obj()
    annots = mupdf.pdf_dict_get(obj, fitz.PDF_NAME('Annots'))

    for index in range(mupdf.pdf_array_len(annots)):
        annot = mupdf.pdf_array_get(annots, index)
        for key in ('Rect', 'QuadPoints', 'L', 'Vertices', 'InkList', 'CL', 'RD', 'Border'):
            value = mupdf.pdf_dict_get(annot, fitz.PDF_NAME(key))
            if mupdf.pdf_is_array(value):
                transformed = _transform_annotation_array(pdf, value, scale, dx, dy, key not in ('RD', 'Border'))
                mupdf.pdf_dict_put(annot, fitz.PDF_NAME(key), transformed)
        border = mupdf.pdf_dict_get(annot, fitz.PDF_NAME('BS'))
        if mupdf.pdf_is_dict(border):
            border = mupdf.pdf_copy_dict(border)
            mupdf.pdf_dict_put(annot, fitz.PDF_NAME('BS'), border)
        width = mupdf.pdf_dict_get(border, fitz.PDF_NAME('W'))
        if mupdf.pdf_is_number(width):
            mupdf.pdf_dict_put_real(border, fitz.PDF_NAME('W'), mupdf.pdf_to_real(width) * scale)
        dash = mupdf.pdf_dict_get(border, fitz.PDF_NAME('D'))
        if mupdf.pdf_is_array(dash):
            mupdf.pdf_dict_put(border, fitz.PDF_NAME('D'), _transform_annotation_array(pdf, dash, scale, dx, dy, False))


def _wrap_page_contents(page, scale, dx, dy, crop):
    doc = page.parent
    original = page.get_contents()
    prefix = (f'q\n{scale:g} 0 0 {scale:g} {dx:g} {dy:g} cm\n'
              f'{crop.x0:g} {crop.y0:g} {crop.width:g} {crop.height:g} re W n\n')
    wrappers = []
    for content in (prefix.encode('ascii'), b'\nQ\n'):
        xref = doc.get_new_xref()
        doc.update_object(xref, '<<>>')
        doc.update_stream(xref, content)
        wrappers.append(xref)
    references = [wrappers[0], *original, wrappers[1]]
    doc.xref_set_key(page.xref, 'Contents', '[' + ' '.join(f'{xref} 0 R' for xref in references) + ']')


def _resize_pages_with_padding(doc, target_rect):
    page_transforms = []
    for page in doc:
        src_rect = page.rect
        page_target_rect = _get_oriented_target_rect(target_rect, src_rect)

        if abs(src_rect.width - page_target_rect.width) <= 1 and abs(src_rect.height - page_target_rect.height) <= 1:
            page_transforms.append(None)
            continue

        scale = min(page_target_rect.width / src_rect.width, page_target_rect.height / src_rect.height)
        dx = (page_target_rect.width - src_rect.width * scale) / 2.0
        dy = (page_target_rect.height - src_rect.height * scale) / 2.0
        physical_target = (fitz.Rect(0, 0, page_target_rect.height, page_target_rect.width)
                           if page.rotation in (90, 270) else page_target_rect)
        crop = _visible_pdf_box(page)
        page_transforms.append({
            "scale": scale,
            "dx": dx,
            "dy": dy,
            "target_rect": page_target_rect,
            "physical_target": physical_target,
            "crop": crop,
            "pdf_dx": (physical_target.width - crop.width * scale) / 2 - crop.x0 * scale,
            "pdf_dy": (physical_target.height - crop.height * scale) / 2 - crop.y0 * scale,
        })

    if not any(page_transforms):
        return 0

    resize_local_destinations(doc, page_transforms)
    # 先调整每页内容和页面内对象坐标
    for page_index, page in enumerate(doc):
        transform = page_transforms[page_index]
        if transform is None:
            continue

        scale = transform["scale"]
        physical_target = transform['physical_target']
        crop = transform['crop']
        _wrap_page_contents(page, scale, transform['pdf_dx'], transform['pdf_dy'], crop)
        _transform_annotation_geometry(page, scale, transform['pdf_dx'], transform['pdf_dy'])
        page.set_mediabox(physical_target)
        page.set_cropbox(page.mediabox)

    return sum(1 for item in page_transforms if item is not None)

# 导出与导入书签 (CSV)


def resize_pages_with_padding(doc, target_rect):
    return _resize_pages_with_padding(doc, target_rect)

