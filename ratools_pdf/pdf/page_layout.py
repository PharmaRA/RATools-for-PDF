import fitz


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
        page_transforms.append({
            "scale": scale,
            "dx": dx,
            "dy": dy,
            "target_rect": page_target_rect,
        })

    if not any(page_transforms):
        return 0

    # 先调整每页内容和页面内对象坐标
    for page_index, page in enumerate(doc):
        transform = page_transforms[page_index]
        if transform is None:
            continue

        scale = transform["scale"]
        dx = transform["dx"]
        dy = transform["dy"]
        page_target_rect = transform["target_rect"]

        if not page.is_wrapped:
            page.wrap_contents()

        for xref in page.get_contents():
            old_stream = doc.xref_stream(xref).decode("latin1", "ignore")
            new_stream = f"q\n{scale} 0 0 {scale} {dx} {dy} cm\n{old_stream}\nQ\n"
            doc.update_stream(xref, new_stream.encode("latin1"))

        links = page.get_links()
        annots = list(page.annots() or [])

        page.set_mediabox(page_target_rect)
        page.set_cropbox(page.mediabox)

        for link in links:
            try:
                link["from"] = _transform_rect(fitz.Rect(link["from"]), scale, dx, dy)
                if link.get("kind") == fitz.LINK_GOTO and link.get("to") is not None:
                    dest_page = int(link.get("page", page_index))
                    dest_transform = page_transforms[dest_page] if 0 <= dest_page < len(page_transforms) else None
                    if dest_transform is not None:
                        d_scale = dest_transform["scale"]
                        d_dx = dest_transform["dx"]
                        d_dy = dest_transform["dy"]
                        to_point = link.get("to")
                        if hasattr(to_point, "x") and hasattr(to_point, "y"):
                            link["to"] = _transform_point(to_point, d_scale, d_dx, d_dy)
                page.update_link(link)
            except Exception:
                continue

        for annot in annots:
            try:
                annot.set_rect(_transform_rect(annot.rect, scale, dx, dy))
                annot.update()
            except Exception:
                continue

    # 再调整书签目标坐标
    toc = doc.get_toc(simple=False)
    if toc:
        toc_changed = False
        for item in toc:
            if len(item) < 4 or not isinstance(item[3], dict):
                continue
            dest = item[3]
            if dest.get("kind") != fitz.LINK_GOTO:
                continue

            target_page = item[2] - 1
            if not (0 <= target_page < len(page_transforms)):
                continue

            transform = page_transforms[target_page]
            to_point = dest.get("to")
            if transform is None or to_point is None or not hasattr(to_point, "x") or not hasattr(to_point, "y"):
                continue

            scale = transform["scale"]
            dx = transform["dx"]
            dy = transform["dy"]
            dest["to"] = _transform_point(to_point, scale, dx, dy)
            toc_changed = True

        if toc_changed:
            doc.set_toc(toc)

    return sum(1 for item in page_transforms if item is not None)

# 导出与导入书签 (CSV)


def resize_pages_with_padding(doc, target_rect):
    return _resize_pages_with_padding(doc, target_rect)

