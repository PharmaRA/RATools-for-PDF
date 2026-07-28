import csv
import json
from urllib.parse import unquote

import fitz


def export_bookmarks(pdf_path, csv_path):
    """Export bookmarks to CSV with enough action data for round-trip import."""
    doc = fitz.open(pdf_path)
    toc = doc.get_toc(simple=False)
    with open(csv_path, 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.writer(f)
        writer.writerow(['Level', 'Title', 'Page', 'Kind', 'Uri', 'File', 'TargetPage', 'ToX', 'ToY', 'Zoom', 'NewWindow'])
        for item in toc:
            lvl, title, page, dest = item
            if not isinstance(dest, dict):
                dest = {}
            point = dest.get('to')
            writer.writerow([
                lvl,
                title,
                page,
                dest.get('kind', fitz.LINK_NONE),
                dest.get('uri', ''),
                dest.get('file', ''),
                dest.get('page', ''),
                getattr(point, 'x', ''),
                getattr(point, 'y', ''),
                dest.get('zoom', ''),
                dest.get('newWindow', False),
            ])
    doc.close()


def import_bookmarks(pdf_path, csv_path, output_path):
    """Import bookmarks from CSV while remaining compatible with the old 3-column format."""
    doc = fitz.open(pdf_path)
    new_toc = []

    with open(csv_path, 'r', encoding='utf-8-sig') as f:
        reader = csv.DictReader(f)
        for row in reader:
            try:
                lvl = int(row.get('Level', 1))
                title = row.get('Title', '')
                page = int(row.get('Page', 1))
                page = max(1, min(page, doc.page_count))

                kind_text = str(row.get('Kind', '') or '').strip()
                if not kind_text:
                    new_toc.append([lvl, title, page])
                    continue

                kind = int(kind_text)
                if kind == fitz.LINK_URI:
                    dest = {
                        'kind': fitz.LINK_URI,
                        'uri': row.get('Uri', ''),
                    }
                elif kind == fitz.LINK_GOTO:
                    target_page = int(row.get('TargetPage', page - 1) or (page - 1))
                    dest = {
                        'kind': fitz.LINK_GOTO,
                        'page': max(0, min(target_page, doc.page_count - 1)),
                        'to': fitz.Point(float(row.get('ToX', 72.0) or 72.0), float(row.get('ToY', 36.0) or 36.0)),
                        'zoom': float(row.get('Zoom', 0.0) or 0.0),
                    }
                elif kind == fitz.LINK_GOTOR:
                    target_page = int(row.get('TargetPage', 0) or 0)
                    dest = {
                        'kind': fitz.LINK_GOTOR,
                        'file': row.get('File', ''),
                        'page': max(0, target_page),
                        'to': fitz.Point(float(row.get('ToX', 72.0) or 72.0), float(row.get('ToY', 36.0) or 36.0)),
                        'zoom': float(row.get('Zoom', 0.0) or 0.0),
                        'newWindow': str(row.get('NewWindow', 'false')).lower() == 'true',
                    }
                elif kind == fitz.LINK_LAUNCH:
                    dest = {
                        'kind': fitz.LINK_LAUNCH,
                        'file': row.get('File', ''),
                        'newWindow': str(row.get('NewWindow', 'false')).lower() == 'true',
                    }
                else:
                    dest = {'kind': fitz.LINK_NONE}

                new_toc.append([lvl, title, page, dest])
            except (ValueError, TypeError):
                continue

    if new_toc:
        for i in range(len(new_toc)):
            if i == 0:
                new_toc[i][0] = 1
            else:
                prev_lvl = new_toc[i - 1][0]
                if new_toc[i][0] > prev_lvl + 1:
                    new_toc[i][0] = prev_lvl + 1

    doc.set_toc(new_toc)
    doc.save(output_path, garbage=2, deflate=True, use_objstms=1)
    doc.close()


# 导出与导入超链接 (JSON)

# 外部链接：指向网页/邮件 (URI)、其它 PDF 文件 (GOTOR) 或外部程序/文件 (LAUNCH)。
# 文档内跳转 (GOTO) 属于内部链接，不计入“仅外链”范围。
_EXTERNAL_LINK_KINDS = {
    fitz.LINK_URI,
    fitz.LINK_GOTOR,
    getattr(fitz, "LINK_LAUNCH", None),
}
_EXTERNAL_LINK_KINDS.discard(None)


def _is_external_link_kind(kind):
    return kind in _EXTERNAL_LINK_KINDS


def _dest_page_index(dest):
    """读取目的地里的 0 基页码；缺失或无法解析时返回 None。"""
    value = dest.get("page")
    if value is None:
        return None
    try:
        return int(value)
    except (TypeError, ValueError):
        return None


def bookmark_named_dest_page(dest):
    """命名目标解析出的 0 基页码；悬空（指向不存在的命名目标）时返回 None。

    PyMuPDF 对 ``/GoTo`` + 命名目标一律报 LINK_NAMED：能解析时目的地里带
    ``page`` 键，悬空时只剩 ``nameddest``。
    """
    page_idx = _dest_page_index(dest)
    if page_idx is None or page_idx < 0:
        return None
    return page_idx


def is_bookmark_dest_invalid(dest, bm_page, page_count):
    """书签目标是否失效，供预检与处理共用以保证判定一致。

    三类失效：无动作的空书签、内部跳转指向不存在的页码、命名目标无法解析。
    """
    if not isinstance(dest, dict):
        return True

    kind = dest.get("kind", fitz.LINK_NONE)
    if kind == fitz.LINK_NONE:
        return True
    if kind not in (fitz.LINK_GOTO, fitz.LINK_NAMED):
        return False

    if kind == fitz.LINK_NAMED and bookmark_named_dest_page(dest) is None:
        return True

    try:
        bm_page = int(bm_page)
    except (TypeError, ValueError):
        return True
    return bm_page < 1 or bm_page > page_count


def _rects_overlap(rect_a, rect_b, min_ratio=0.5):
    """两个矩形是否被视为“同一区域”：交集面积占较小矩形的比例达到阈值即算重复。"""
    inter = rect_a & rect_b
    if inter.is_empty or inter.width <= 0 or inter.height <= 0:
        return False
    inter_area = inter.width * inter.height
    smaller_area = min(rect_a.width * rect_a.height, rect_b.width * rect_b.height)
    if smaller_area <= 0:
        return False
    return (inter_area / smaller_area) >= min_ratio


def _read_link_quad_points(doc, xref):
    """读取链接注释的 QuadPoints 数组，返回浮点数列表；无则返回 None。

    QuadPoints 为 8×N 个数字（每 8 个描述一个四边形），跨行链接会有多组。
    """
    if not xref:
        return None
    try:
        key_type, value = doc.xref_get_key(xref, "QuadPoints")
    except Exception:
        return None
    if key_type != "array" or not value:
        return None
    numbers = []
    for token in value.strip("[]").split():
        try:
            numbers.append(float(token))
        except ValueError:
            return None
    # 必须是完整的四边形组（每组 8 个数），否则视为无效。
    if not numbers or len(numbers) % 8 != 0:
        return None
    return numbers


def _write_link_quad_points(doc, xref, quad_points):
    """把 QuadPoints 数组写回链接注释，用于还原跨行链接的逐行可点区域。"""
    if not xref or not quad_points or len(quad_points) % 8 != 0:
        return
    # PDF 数字尽量输出整数形式，保持与原始注释一致的紧凑写法。
    tokens = []
    for value in quad_points:
        if float(value).is_integer():
            tokens.append(str(int(value)))
        else:
            tokens.append(repr(float(value)))
    array_str = "[" + " ".join(tokens) + "]"
    try:
        doc.xref_set_key(xref, "QuadPoints", array_str)
    except Exception:
        pass


def export_links(pdf_path, json_path, scope="all"):
    """Export links to JSON with enough destination data for round-trip import.

    scope: "all" 导出全部链接；"external" 仅导出外部链接 (URI/GOTOR/LAUNCH)。
    """
    only_external = scope == "external"
    doc = fitz.open(pdf_path)
    all_links = []

    for page in doc:
        for link in page.get_links():
            kind = link.get('kind', fitz.LINK_NONE)
            if only_external and not _is_external_link_kind(kind):
                continue
            rect = link['from']
            target_point = link.get('to')
            link_dict = {
                'page_index': page.number,
                'rect': [rect.x0, rect.y0, rect.x1, rect.y1],
                'kind': kind,
                'uri': link.get('uri', ''),
                'file': link.get('file', ''),
                'target_page': link.get('page', 0),
                'zoom': link.get('zoom', 0.0),
                'to': [getattr(target_point, 'x', 0.0), getattr(target_point, 'y', 0.0)] if target_point else None,
                'new_window': bool(link.get('newWindow', False)),
                # 跨行链接由多个四边形 (QuadPoints) 组成，link['from'] 只是它们的合并包围盒。
                # 保存原始 QuadPoints，导入时精确还原每行的可点区域，避免把行间内容一并圈入。
                'quad_points': _read_link_quad_points(doc, link.get('xref', 0)),
            }
            all_links.append(link_dict)

    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(all_links, f, indent=4, ensure_ascii=False)
    doc.close()


def _build_link_payload(ld, link_file_kind):
    """将 JSON 中的一条链接记录转换为 insert_link 所需的字典。"""
    rect = fitz.Rect(ld['rect'])
    kind = ld['kind']

    new_link = {"kind": kind, "from": rect}
    if kind == fitz.LINK_URI:
        new_link["uri"] = ld.get('uri', '')
    elif link_file_kind is not None and kind == link_file_kind:
        new_link["file"] = ld.get('file', '')
    elif kind in [fitz.LINK_GOTO, fitz.LINK_GOTOR]:
        new_link["page"] = ld.get('target_page', 0)
        new_link["zoom"] = ld.get('zoom', 0.0)
        to_value = ld.get('to')
        if isinstance(to_value, (list, tuple)) and len(to_value) >= 2:
            new_link["to"] = fitz.Point(float(to_value[0]), float(to_value[1]))
        if kind == fitz.LINK_GOTOR:
            new_link["file"] = ld.get('file', '')
            new_link["newWindow"] = bool(ld.get('new_window', False))
    elif kind == fitz.LINK_LAUNCH:
        new_link["file"] = ld.get('file', '')
        new_link["newWindow"] = bool(ld.get('new_window', False))

    return new_link, rect


def import_links(pdf_path, json_path, output_path, scope="all", mode="overwrite"):
    """Rebuild links from JSON while preserving internal destinations and new-window flags.

    scope: "all" 导入全部链接；"external" 仅导入外部链接 (URI/GOTOR/LAUNCH)。
    mode:  "overwrite" 覆盖——先移除页面上同范围的旧链接再写入；
           "incremental" 增量——保留旧链接，仅写入未与既有链接重叠的新链接。
    """
    only_external = scope == "external"
    incremental = mode == "incremental"
    doc = fitz.open(pdf_path)
    link_file_kind = getattr(fitz, "LINK_FILE", None)

    with open(json_path, 'r', encoding='utf-8') as f:
        links_data = json.load(f)

    # 覆盖模式：移除页面上处于导入范围内的旧链接。
    # “仅外链”范围只删除外链，保留文档内跳转；“全部”范围删除所有链接。
    if not incremental:
        for page in doc:
            for link in page.get_links():
                if only_external and not _is_external_link_kind(link.get('kind', fitz.LINK_NONE)):
                    continue
                page.delete_link(link)

    # 增量模式下，记录各页现存链接区域，用于跳过重复区域。
    existing_rects = {}
    if incremental:
        for page in doc:
            existing_rects[page.number] = [
                fitz.Rect(link['from']) for link in page.get_links()
            ]

    for ld in links_data:
        p_idx = ld.get('page_index', 0)
        if not (0 <= p_idx < doc.page_count):
            continue

        kind = ld.get('kind', fitz.LINK_NONE)
        if only_external and not _is_external_link_kind(kind):
            continue

        page = doc[p_idx]
        new_link, rect = _build_link_payload(ld, link_file_kind)

        if incremental:
            page_rects = existing_rects.setdefault(p_idx, [])
            if any(_rects_overlap(rect, existing) for existing in page_rects):
                continue

        quad_points = ld.get('quad_points')
        try:
            # insert_link 不返回 xref，且保存前 get_links() 读不到新链接，
            # 因此用 page.annot_xrefs() 做插入前后差集来定位新注释，
            # 以便把跨行链接的 QuadPoints 精确写回。
            xrefs_before = {item[0] for item in page.annot_xrefs()}
            page.insert_link(new_link)
            if quad_points:
                new_xrefs = {item[0] for item in page.annot_xrefs()} - xrefs_before
                for xref in new_xrefs:
                    _write_link_quad_points(doc, xref, quad_points)
            if incremental:
                existing_rects[p_idx].append(rect)
        except Exception:
            pass

    doc.save(output_path, garbage=2, deflate=True, use_objstms=1)
    doc.close()


