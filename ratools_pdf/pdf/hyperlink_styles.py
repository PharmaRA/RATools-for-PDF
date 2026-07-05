import csv
import json
import os
import re
import shutil
import subprocess
import sys
from pathlib import Path
from urllib.parse import unquote, urlparse

import fitz

from ratools_pdf.config.paths import get_resource_path
from ratools_pdf.pdf.font_embedding_providers import get_font_embedding_provider


def _processor_cls():
    from ratools_pdf.pdf.processor import PDFProcessor
    return PDFProcessor


def _is_text_blue(page, rect):
    text_dict = page.get_text("dict", clip=rect)
    for block in text_dict.get("blocks", []):
        for line in block.get("lines", []):
            for span in line.get("spans", []):
                color = span.get("color", 0)
                b = color & 0xFF
                g = (color >> 8) & 0xFF
                r = (color >> 16) & 0xFF
                if b > r + 40 and b > g + 40:
                    return True
    return False


def _overlay_text_color_in_rect(page, rect, color, skip_if_already_blue=False, erase_background=False):
    try:
        text_dict = page.get_text("dict", clip=rect)
    except Exception:
        return False

    changed = False
    for block in text_dict.get("blocks", []):
        for line in block.get("lines", []):
            for span in line.get("spans", []):
                try:
                    txt = span.get("text", "")
                    if not txt.strip():
                        continue

                    span_color = span.get("color", 0)
                    b = span_color & 0xFF
                    g = (span_color >> 8) & 0xFF
                    r = (span_color >> 16) & 0xFF
                    is_blue = b > r + 40 and b > g + 40
                    if skip_if_already_blue and is_blue:
                        continue

                    bbox = span.get("bbox", None)
                    if not bbox or len(bbox) != 4:
                        continue

                    span_rect = fitz.Rect(bbox)
                    origin = span.get("origin", None)
                    if not origin or len(origin) != 2:
                        origin = (span_rect.x0, span_rect.y1)
                    font_candidates = []
                    original_font = span.get("font", "")
                    if original_font:
                        font_candidates.append(original_font)

                    if any(ord(ch) > 127 for ch in txt):
                        font_candidates.extend(["china-s", "cjk", "helv"])
                    else:
                        font_candidates.append("helv")

                    inserted = False
                    for font_name in font_candidates:
                        try:
                            if erase_background and not inserted:
                                page.draw_rect(span_rect, color=None, fill=(1, 1, 1), overlay=True)

                            page.insert_text(
                                origin,
                                txt,
                                fontsize=span.get("size", 11),
                                fontname=font_name,
                                color=color,
                                overlay=True,
                            )
                            inserted = True
                            changed = True
                            break
                        except Exception:
                            continue
                except Exception:
                    continue

    return changed


def _link_has_visible_border(doc, link_obj):
    border = link_obj.border or {}
    if border.get("width", 0) > 0:
        return True

    xref = getattr(link_obj, "xref", 0)
    if not xref:
        return False

    try:
        raw_obj = doc.xref_object(xref)
    except Exception:
        return False

    bs_match = re.search(r"/BS\s*<<.*?/W\s+([0-9.]+)", raw_obj, re.S)
    if bs_match:
        try:
            return float(bs_match.group(1)) > 0
        except Exception:
            pass

    border_match = re.search(r"/Border\s*\[\s*[0-9.]+\s+[0-9.]+\s+([0-9.]+)", raw_obj)
    if border_match:
        try:
            return float(border_match.group(1)) > 0
        except Exception:
            pass

    return False


def _force_link_new_window(doc, xref):
    if not xref:
        return

    try:
        link_obj = doc.xref_object(xref)
        if "/NewWindow" in link_obj:
            link_obj = re.sub(r"/NewWindow\s+(true|false)", "/NewWindow true", link_obj)
        elif "/S /GoToR" in link_obj:
            link_obj = link_obj.replace("/S /GoToR", "/S /GoToR\n    /NewWindow true", 1)
        elif "/S /Launch" in link_obj:
            link_obj = link_obj.replace("/S /Launch", "/S /Launch\n    /NewWindow true", 1)
        doc.update_object(xref, link_obj)
    except Exception:
        pass


def _rects_intersect(a, b):
    return not (a.x1 <= b.x0 or a.x0 >= b.x1 or a.y1 <= b.y0 or a.y0 >= b.y1)


def _point_in_any_rect(point, rects):
    return any(rect.contains(point) for rect in rects)


def _make_text_block_blue(block_text):
    return _processor_cls()._make_text_block_color(block_text, (0.0, 0.0, 1.0))


def _make_text_block_color(block_text, color_rgb):
    r, g, b = color_rgb
    color_cmd = f"{r:g} {g:g} {b:g} rg"

    if " rg" in block_text:
        return re.sub(r"(?<![0-9.])-?[0-9.]+\s+-?[0-9.]+\s+-?[0-9.]+\s+rg", color_cmd, block_text, count=1)
    if " g" in block_text:
        return re.sub(r"(?<![0-9.])-?[0-9.]+\s+g", color_cmd, block_text, count=1)

    tj_pos = block_text.find("TJ")
    if tj_pos == -1:
        tj_pos = block_text.find("Tj")
    if tj_pos == -1:
        return block_text

    return block_text[:tj_pos] + color_cmd + "\n" + block_text[tj_pos:]


def _apply_text_color_via_content_stream(doc, page, target_rects, color_rgb, only_if_blue=False):
    if not target_rects:
        return False

    target_indexes = set()
    for trace_index, trace in enumerate(page.get_texttrace()):
        if trace.get("type") != 0:
            continue

        bbox = fitz.Rect(trace.get("bbox", (0, 0, 0, 0)))
        if not any(_processor_cls()._rects_intersect(bbox, rect) for rect in target_rects):
            continue

        chars = trace.get("chars", ())
        visible_char_count = 0
        inside_char_count = 0
        for ch in chars:
            if len(ch) < 4:
                continue
            unicode_codepoint = ch[0]
            if unicode_codepoint in (9, 10, 13, 32):
                continue

            char_bbox = ch[3]
            if not char_bbox or len(char_bbox) != 4:
                continue

            visible_char_count += 1
            rect = fitz.Rect(char_bbox)
            center = fitz.Point((rect.x0 + rect.x1) / 2.0, (rect.y0 + rect.y1) / 2.0)
            if _processor_cls()._point_in_any_rect(center, target_rects):
                inside_char_count += 1

        if visible_char_count == 0:
            continue
        if inside_char_count == 0:
            continue
        if inside_char_count != visible_char_count:
            continue

        color = trace.get("color", (0.0, 0.0, 0.0))
        if only_if_blue and isinstance(color, tuple) and len(color) >= 3:
            is_blue = color[2] > color[0] + 0.1 and color[2] > color[1] + 0.1
            if not is_blue:
                continue
        target_indexes.add(trace_index)

    if not target_indexes:
        return False

    text_block_index = 0
    changed = False
    content_xrefs = page.get_contents()
    if not isinstance(content_xrefs, (list, tuple)):
        content_xrefs = [content_xrefs]

    for xref in content_xrefs:
        stream_bytes = doc.xref_stream(xref)
        stream_text = stream_bytes.decode("latin1", "ignore")

        def replace_bt_block(match):
            nonlocal text_block_index, changed
            block = match.group(0)
            if "Tj" not in block and "TJ" not in block:
                return block

            current_index = text_block_index
            text_block_index += 1
            if current_index not in target_indexes:
                return block

            new_block = _processor_cls()._make_text_block_color(block, color_rgb)
            if new_block != block:
                changed = True
            return new_block

        new_stream_text = re.sub(r"BT.*?ET", replace_bt_block, stream_text, flags=re.S)
        if new_stream_text != stream_text:
            doc.update_stream(xref, new_stream_text.encode("latin1"))

    return changed


def _collect_page_state(page):
    links = page.get_links()
    annots = list(page.annots() or [])
    link_objs = []
    link_rects = []
    link_obj = page.first_link
    while link_obj:
        link_objs.append(link_obj)
        try:
            link_rects.append(link_obj.rect)
        except Exception:
            pass
        link_obj = link_obj.next
    return {
        "links": links,
        "annots": annots,
        "link_objs": link_objs,
        "link_rects": link_rects,
    }


def _apply_blue_text_via_content_stream(doc, page, link_rects=None):
    if link_rects is None:
        link_rects = []
        link_obj = page.first_link
        while link_obj:
            link_rects.append(link_obj.rect)
            link_obj = link_obj.next
    return _processor_cls()._apply_text_color_via_content_stream(
        doc,
        page,
        link_rects,
        (0.0, 0.0, 1.0),
        only_if_blue=False,
    )


def _apply_hyperlink_actions(doc, page, options, file_like_link_kinds, page_links=None):
    changed = False

    links = page_links if page_links is not None else page.get_links()
    for link in links:
        link_modified = False
        force_new_window = False
        kind = link.get("kind", fitz.LINK_NONE)

        if "link_abs_to_rel_path" in options and kind in file_like_link_kinds:
            file_path = link.get("file", "")
            decoded_file_path = unquote(file_path) if file_path else ""
            if decoded_file_path and (
                ":" in decoded_file_path or decoded_file_path.startswith("/") or decoded_file_path.startswith("\\")
            ):
                link["file"] = os.path.basename(decoded_file_path.replace("\\", "/"))
                link_modified = True

        if "link_inherit_zoom" in options and kind == fitz.LINK_GOTO:
            if link.get("zoom") != 0.0:
                link["zoom"] = 0.0
                link_modified = True

        if "link_open_new_window" in options and kind in [fitz.LINK_GOTOR, fitz.LINK_LAUNCH]:
            if not link.get("newWindow"):
                force_new_window = True

        if link_modified:
            page.update_link(link)
            changed = True

        if force_new_window:
            # Do not call update_link() just to set NewWindow. On some Windows/PyMuPDF
            # combinations, rebuilding external-file links can normalize relative /F paths
            # into absolute paths. Patch the raw action object instead so /F and /UF stay intact.
            _processor_cls()._force_link_new_window(doc, link.get("xref", 0))
            changed = True

    return changed


def _apply_hyperlink_styles(doc, page, options, link_objs=None, link_rects=None):
    changed = False

    if "link_text_blue" in options:
        if _processor_cls()._apply_blue_text_via_content_stream(doc, page, link_rects=link_rects):
            changed = True

    iterable = link_objs if link_objs is not None else []
    if link_objs is None:
        tmp = page.first_link
        while tmp:
            iterable.append(tmp)
            tmp = tmp.next

    for link_obj in iterable:
        link_changed = False
        has_border = _processor_cls()._link_has_visible_border(doc, link_obj)

        if "link_remove_border" in options:
            if has_border:
                link_obj.set_border(width=0)
                link_changed = True
        elif "link_black_border" in options:
            link_obj.set_border(width=1.0)
            link_obj.set_colors(stroke=(0, 0, 0))
            link_changed = True
        elif "link_bordered_to_blue_border" in options:
            if has_border:
                link_obj.set_border(width=1.0)
                link_obj.set_colors(stroke=(0, 0, 1))
                link_changed = True
        elif "link_unbordered_blue_to_blue_border" in options:
            if not has_border and _processor_cls()._is_text_blue(page, link_obj.rect):
                link_obj.set_border(width=1.0)
                link_obj.set_colors(stroke=(0, 0, 1))
                link_changed = True

        if link_changed:
            changed = True

    return changed

# ====================================================
# 六大核心合规清理模块入口
# ====================================================


def apply_hyperlink_styles(doc, page, options, link_objs=None, link_rects=None):
    return _apply_hyperlink_styles(doc, page, options, link_objs, link_rects)

