import fitz
import os
import sys
import subprocess
import shutil
import csv
import json
import re
from urllib.parse import unquote
from pathlib import Path

from app_paths import get_resource_path


class PDFProcessor:

    @staticmethod
    def _get_gs_path():
        if sys.platform == "win32":
            gs_exe = get_resource_path("plugins", "ghostscript", "bin", "gswin64c.exe")
        elif sys.platform == "darwin":
            gs_exe = get_resource_path("plugins", "ghostscript", "bin", "gs")
        else:
            gs_exe = "gs"
        return gs_exe

    @staticmethod
    def _embed_fonts_with_gs(input_pdf, output_pdf):
        PDFProcessor._rewrite_with_gs(input_pdf, output_pdf, embed_fonts=True, linearize=False)

    @staticmethod
    def _rewrite_with_gs(input_pdf, output_pdf, embed_fonts=False, linearize=False):
        gs_exe = PDFProcessor._get_gs_path()
        if sys.platform in ["win32", "darwin"] and not os.path.exists(gs_exe):
            raise FileNotFoundError(f"未找到 Ghostscript 引擎！\n请确保已将引擎文件放置在: {gs_exe}")

        cmd = [gs_exe, "-sDEVICE=pdfwrite", "-dCompatibilityLevel=1.7", "-dNOPAUSE", "-dQUIET", "-dBATCH"]

        if embed_fonts:
            cmd.extend(["-dPDFSETTINGS=/prepress", "-dSubsetFonts=true", "-dEmbedAllFonts=true"])
        if linearize:
            cmd.append("-dFastWebView=true")

        cmd.extend([f"-sOutputFile={output_pdf}", input_pdf])

        startupinfo = None
        if sys.platform == "win32":
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW

        result = subprocess.run(cmd, startupinfo=startupinfo, capture_output=True, text=True)
        if result.returncode != 0:
            raise RuntimeError(f"Ghostscript 执行失败: {result.stderr}")

    @staticmethod
    def _linearize_with_gs(input_pdf, output_pdf):
        PDFProcessor._rewrite_with_gs(input_pdf, output_pdf, embed_fonts=False, linearize=True)

    @staticmethod
    def _mark_change(change_list, label):
        if label not in change_list:
            change_list.append(label)

    @staticmethod
    def _increase_change_count(change_counts, label, amount=1):
        if amount <= 0:
            return
        change_counts[label] = change_counts.get(label, 0) + amount

    @staticmethod
    def _format_change_summary(change_counts, ordered_labels):
        parts = []
        for label in ordered_labels:
            if label in change_counts:
                count = change_counts[label]
                if count > 1:
                    parts.append(f"{label}({count}处)")
                else:
                    parts.append(label)
            else:
                parts.append(label)
        return "、".join(parts)

    # 导出与导入书签 (CSV)
    @staticmethod
    def export_bookmarks(pdf_path, csv_path):
        """将书签导出为 CSV (格式: 级别, 标题, 页码)"""
        doc = fitz.open(pdf_path)
        toc = doc.get_toc(simple=False)
        # 必须使用 utf-8-sig，防止 Excel 打开中文乱码
        with open(csv_path, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f)
            writer.writerow(['Level', 'Title', 'Page'])
            for item in toc:
                lvl, title, page, _ = item
                writer.writerow([lvl, title, page])
        doc.close()

    @staticmethod
    def import_bookmarks(pdf_path, csv_path, output_path):
        """读取 CSV 结构并强制写入 PDF 书签，带有防断层与防越界保护"""
        doc = fitz.open(pdf_path)
        new_toc = []

        with open(csv_path, 'r', encoding='utf-8-sig') as f:
            reader = csv.DictReader(f)
            for row in reader:
                try:
                    lvl = int(row.get('Level', 1))
                    title = row.get('Title', '')
                    page = int(row.get('Page', 1))

                    # 防越界：强制限制在真实页码范围内
                    page = max(1, min(page, doc.page_count))
                    new_toc.append([lvl, title, page])
                except ValueError:
                    continue  # 跳过无法解析格式的脏数据行

        # 防断层算法：拉平非法的书签级别跨越 (例如: 1 -> 3 会报错，强制转为 1 -> 2)
        if new_toc:
            for i in range(len(new_toc)):
                if i == 0:
                    new_toc[i][0] = 1
                else:
                    prev_lvl = new_toc[i - 1][0]
                    if new_toc[i][0] > prev_lvl + 1:
                        new_toc[i][0] = prev_lvl + 1

        doc.set_toc(new_toc)
        doc.save(output_path, garbage=3, deflate=True)
        doc.close()


    # 导出与导入超链接 (JSON)
    @staticmethod
    def export_links(pdf_path, json_path):
        """将超链接的物理坐标与动作类型提取至 JSON 文件"""
        doc = fitz.open(pdf_path)
        all_links = []

        for page in doc:
            for link in page.get_links():
                rect = link['from']
                link_dict = {
                    'page_index': page.number,  # PyMuPDF 中页面索引从 0 开始
                    'rect': [rect.x0, rect.y0, rect.x1, rect.y1],
                    'kind': link.get('kind', fitz.LINK_NONE),
                    'uri': link.get('uri', ''),
                    'file': link.get('file', ''),
                    'target_page': link.get('page', 0),  # 仅 GOTO 有效
                    'zoom': link.get('zoom', 0.0)
                }
                all_links.append(link_dict)

        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(all_links, f, indent=4, ensure_ascii=False)
        doc.close()

    @staticmethod
    def import_links(pdf_path, json_path, output_path):
        """清除原有链接，根据 JSON 精准复原超链接布局"""
        doc = fitz.open(pdf_path)
        link_file_kind = getattr(fitz, "LINK_FILE", None)

        with open(json_path, 'r', encoding='utf-8') as f:
            links_data = json.load(f)

        # 先清空原有超链接，防止重复叠加
        for page in doc:
            for link in page.get_links():
                page.delete_link(link)

        for ld in links_data:
            p_idx = ld.get('page_index', 0)
            if 0 <= p_idx < doc.page_count:
                page = doc[p_idx]
                rect = fitz.Rect(ld['rect'])
                kind = ld['kind']

                # 构建 PyMuPDF 所需的动作字典
                new_link = {"kind": kind, "from": rect}
                if kind == fitz.LINK_URI:
                    new_link["uri"] = ld.get('uri', '')
                elif link_file_kind is not None and kind == link_file_kind:
                    new_link["file"] = ld.get('file', '')
                elif kind in [fitz.LINK_GOTO, fitz.LINK_GOTOR]:
                    new_link["page"] = ld.get('target_page', 0)
                    new_link["zoom"] = ld.get('zoom', 0.0)
                    if kind == fitz.LINK_GOTOR:
                        new_link["file"] = ld.get('file', '')

                try:
                    page.insert_link(new_link)
                except Exception:
                    pass  # 忽略错误坐标导致无法注入的脏链接

        doc.save(output_path, garbage=3, deflate=True)
        doc.close()

    @staticmethod
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

    @staticmethod
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

    @staticmethod
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

    @staticmethod
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

    @staticmethod
    def _rects_intersect(a, b):
        return not (a.x1 <= b.x0 or a.x0 >= b.x1 or a.y1 <= b.y0 or a.y0 >= b.y1)

    @staticmethod
    def _point_in_any_rect(point, rects):
        return any(rect.contains(point) for rect in rects)

    @staticmethod
    def _make_text_block_blue(block_text):
        return PDFProcessor._make_text_block_color(block_text, (0.0, 0.0, 1.0))

    @staticmethod
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

    @staticmethod
    def _apply_text_color_via_content_stream(doc, page, target_rects, color_rgb, only_if_blue=False):
        if not target_rects:
            return False

        target_indexes = set()
        for trace_index, trace in enumerate(page.get_texttrace()):
            if trace.get("type") != 0:
                continue

            bbox = fitz.Rect(trace.get("bbox", (0, 0, 0, 0)))
            if not any(PDFProcessor._rects_intersect(bbox, rect) for rect in target_rects):
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
                if PDFProcessor._point_in_any_rect(center, target_rects):
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

                new_block = PDFProcessor._make_text_block_color(block, color_rgb)
                if new_block != block:
                    changed = True
                return new_block

            new_stream_text = re.sub(r"BT.*?ET", replace_bt_block, stream_text, flags=re.S)
            if new_stream_text != stream_text:
                doc.update_stream(xref, new_stream_text.encode("latin1"))

        return changed

    @staticmethod
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

    @staticmethod
    def _apply_blue_text_via_content_stream(doc, page, link_rects=None):
        if link_rects is None:
            link_rects = []
            link_obj = page.first_link
            while link_obj:
                link_rects.append(link_obj.rect)
                link_obj = link_obj.next
        return PDFProcessor._apply_text_color_via_content_stream(
            doc,
            page,
            link_rects,
            (0.0, 0.0, 1.0),
            only_if_blue=False,
        )

    @staticmethod
    def _apply_hyperlink_actions(doc, page, options, file_like_link_kinds, page_links=None):
        changed = False

        links = page_links if page_links is not None else page.get_links()
        for link in links:
            link_modified = False
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
                    link["newWindow"] = True
                    link_modified = True

            if not link_modified:
                continue

            page.update_link(link)
            if "link_open_new_window" in options and kind in [fitz.LINK_GOTOR, fitz.LINK_LAUNCH]:
                PDFProcessor._force_link_new_window(doc, link.get("xref", 0))
            changed = True

        return changed

    @staticmethod
    def _apply_hyperlink_styles(doc, page, options, link_objs=None, link_rects=None):
        changed = False

        if "link_text_blue" in options:
            if PDFProcessor._apply_blue_text_via_content_stream(doc, page, link_rects=link_rects):
                changed = True

        iterable = link_objs if link_objs is not None else []
        if link_objs is None:
            tmp = page.first_link
            while tmp:
                iterable.append(tmp)
                tmp = tmp.next

        for link_obj in iterable:
            link_changed = False
            has_border = PDFProcessor._link_has_visible_border(doc, link_obj)

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
                if not has_border and PDFProcessor._is_text_blue(page, link_obj.rect):
                    link_obj.set_border(width=1.0)
                    link_obj.set_colors(stroke=(0, 0, 1))
                    link_changed = True

            if link_changed:
                changed = True

        return changed

    # ====================================================
    # 六大核心合规清理模块入口
    # ====================================================
    @staticmethod
    def process_document(input_path, output_path, options):
        try:
            doc = fitz.open(input_path)
            applied_changes = []
            change_counts = {}
            link_file_kind = getattr(fitz, "LINK_FILE", None)
            file_like_link_kinds = {fitz.LINK_GOTOR}
            if link_file_kind is not None:
                file_like_link_kinds.add(link_file_kind)

            if doc.needs_pass: return False, "❌ 文件已加密"
            changed = False
            catalog_xref = doc.pdf_catalog()

            needs_gs_engine = any(opt in options for opt in [
                "embed_nonstandard_fonts",
                "convert_pdf_version"
            ])

            if "title_from_filename" in options:
                base_name = Path(input_path).stem
                meta = doc.metadata
                if meta.get("title") != base_name:
                    meta["title"] = base_name;
                    doc.set_metadata(meta);
                    changed = True
                    PDFProcessor._mark_change(applied_changes, "标题同步为文件名")

            if "open_page_first" in options or "zoom_default" in options:
                if doc.page_count > 0:
                    page0_xref = doc[0].xref
                    # /XYZ null null null 表示使用阅读器默认缩放，不强制 Fit/固定倍率
                    action_str = f"[{page0_xref} 0 R /XYZ null null null]"
                    doc.xref_set_key(catalog_xref, "OpenAction", action_str);
                    changed = True
                    if "open_page_first" in options:
                        PDFProcessor._mark_change(applied_changes, "打开页设为第一页")
                    if "zoom_default" in options:
                        PDFProcessor._mark_change(applied_changes, "打开缩放设为默认")

            if "page_layout_default" in options:
                # 恢复为 PDF 阅读器默认行为：移除显式 PageLayout 设置
                doc.xref_set_key(catalog_xref, "PageLayout", "null");
                changed = True
                PDFProcessor._mark_change(applied_changes, "页面布局恢复默认")

            if "initial_view_bookmarks_and_page" in options:
                has_bookmarks = len(doc.get_toc(simple=False)) > 0
                page_mode = "/UseOutlines" if has_bookmarks else "/UseNone"
                doc.xref_set_key(catalog_xref, "PageMode", page_mode)
                changed = True
                PDFProcessor._mark_change(applied_changes, "初始视图设为书签/页面")

            if "collapse_all_bookmarks" in options:
                toc = doc.get_toc(simple=False)
                if toc:
                    for item in toc:
                        if isinstance(item[-1], dict): item[-1]["collapse"] = True
                    doc.set_toc(toc);
                    changed = True
                    PDFProcessor._mark_change(applied_changes, "折叠全部书签")

            if "page_size_a4" in options or "page_size_letter" in options:
                target_rect = fitz.paper_rect("a4") if "page_size_a4" in options else fitz.paper_rect(
                    "letter")
                for page in doc:
                    if abs(page.rect.width - target_rect.width) > 1 or abs(page.rect.height - target_rect.height) > 1:
                        page.set_mediabox(target_rect)
                        page.set_cropbox(target_rect)
                        changed = True
                        PDFProcessor._mark_change(applied_changes, "页面尺寸标准化")
                        PDFProcessor._increase_change_count(change_counts, "页面尺寸标准化")

            def _to_point(value):
                if hasattr(value, "x") and hasattr(value, "y"):
                    return fitz.Point(float(value.x), float(value.y))
                if isinstance(value, (tuple, list)) and len(value) >= 2:
                    try:
                        return fitz.Point(float(value[0]), float(value[1]))
                    except Exception:
                        return fitz.Point(72.0, 36.0)
                return fitz.Point(72.0, 36.0)

            def _normalize_bookmark_dest(dest, kind):
                if not isinstance(dest, dict):
                    dest = {}

                if kind == fitz.LINK_GOTO:
                    try:
                        page_idx = int(dest.get("page", 0))
                    except Exception:
                        page_idx = 0
                    if page_idx < 0:
                        page_idx = 0

                    try:
                        zoom = float(dest.get("zoom", 0.0))
                    except Exception:
                        zoom = 0.0

                    return {
                        "kind": fitz.LINK_GOTO,
                        "page": page_idx,
                        "to": _to_point(dest.get("to")),
                        "zoom": zoom,
                    }

                if kind == fitz.LINK_GOTOR:
                    try:
                        page_idx = int(dest.get("page", 0))
                    except Exception:
                        page_idx = 0
                    if page_idx < 0:
                        page_idx = 0

                    try:
                        zoom = float(dest.get("zoom", 0.0))
                    except Exception:
                        zoom = 0.0

                    file_path = dest.get("file", "")
                    if file_path is None:
                        file_path = ""

                    return {
                        "kind": fitz.LINK_GOTOR,
                        "file": str(file_path),
                        "page": page_idx,
                        "to": _to_point(dest.get("to")),
                        "zoom": zoom,
                        "newWindow": bool(dest.get("newWindow", False)),
                    }

                if kind == fitz.LINK_LAUNCH:
                    file_path = dest.get("file", "")
                    if file_path is None:
                        file_path = ""
                    return {
                        "kind": fitz.LINK_LAUNCH,
                        "file": str(file_path),
                        "newWindow": bool(dest.get("newWindow", False)),
                    }

                if kind == fitz.LINK_URI:
                    uri = dest.get("uri", "")
                    if uri is None:
                        uri = ""
                    return {
                        "kind": fitz.LINK_URI,
                        "uri": str(uri),
                    }

                return {"kind": fitz.LINK_NONE}

            bookmark_options = ["bookmark_inherit_zoom", "bookmark_open_new_window", "bookmark_remove_external_links",
                                 "bookmark_remove_invalid",
                                 "bookmark_remove_unknown_actions"]
            if any(opt in options for opt in bookmark_options):
                toc = doc.get_toc(simple=False)
                if toc:
                    new_toc = []
                    toc_modified = False
                    for item in toc:
                        lvl, title, bm_page, dest = item
                        if not isinstance(lvl, int):
                            try:
                                lvl = int(lvl)
                            except Exception:
                                lvl = 1
                        if lvl < 1:
                            lvl = 1

                        if not isinstance(bm_page, int):
                            try:
                                bm_page = int(bm_page)
                            except Exception:
                                bm_page = 1

                        kind = dest.get("kind", fitz.LINK_NONE)
                        dest = _normalize_bookmark_dest(dest, kind)
                        delete_it = False

                        if "bookmark_remove_external_links" in options and kind == fitz.LINK_URI: delete_it = True
                        if "bookmark_remove_invalid" in options:
                            if kind == fitz.LINK_NONE or (
                                    kind == fitz.LINK_GOTO and (bm_page < 1 or bm_page > doc.page_count)): delete_it = True
                        if "bookmark_remove_unknown_actions" in options:
                            if kind not in [fitz.LINK_GOTO, fitz.LINK_GOTOR, fitz.LINK_LAUNCH]: delete_it = True

                        if delete_it:
                            toc_modified = True;
                            continue

                        if "bookmark_inherit_zoom" in options and kind == fitz.LINK_GOTO:
                            if dest.get("zoom") != 0.0: dest["zoom"] = 0.0; toc_modified = True
                        if "bookmark_open_new_window" in options and kind in [fitz.LINK_GOTOR, fitz.LINK_LAUNCH]:
                            if not dest.get("newWindow"): dest["newWindow"] = True; toc_modified = True

                        if kind == fitz.LINK_GOTO:
                            if bm_page < 1:
                                bm_page = 1; toc_modified = True
                            elif bm_page > doc.page_count:
                                bm_page = doc.page_count; toc_modified = True

                        new_toc.append([lvl, title, bm_page, dest])

                    if toc_modified:
                        if new_toc:
                            for i in range(len(new_toc)):
                                if i == 0:
                                    new_toc[i][0] = 1
                                else:
                                    prev_lvl = new_toc[i - 1][0]
                                    if new_toc[i][0] > prev_lvl + 1: new_toc[i][0] = prev_lvl + 1
                        try:
                            doc.set_toc(new_toc)
                        except Exception:
                            # 容错兜底：若目的地结构异常导致写入失败，降级为基础书签（标题+页码）
                            fallback_toc = []
                            prev_lvl = 1
                            for lvl, title, bm_page, _dest in new_toc:
                                if not isinstance(lvl, int):
                                    lvl = prev_lvl
                                if lvl < 1:
                                    lvl = 1
                                if lvl > prev_lvl + 1:
                                    lvl = prev_lvl + 1

                                if not isinstance(bm_page, int):
                                    try:
                                        bm_page = int(bm_page)
                                    except Exception:
                                        bm_page = 1
                                bm_page = max(1, min(bm_page, doc.page_count))

                                fallback_toc.append([lvl, title, bm_page])
                                prev_lvl = lvl

                            doc.set_toc(fallback_toc)
                        changed = True
                        PDFProcessor._mark_change(applied_changes, "书签规则已更新")

            hyperlink_options = ["link_abs_to_rel_path", "link_inherit_zoom",
                                 "link_open_new_window", "link_text_blue",
                                 "link_black_border", "link_bordered_to_blue_border", "link_unbordered_blue_to_blue_border",
                                 "link_remove_border"]
            if any(opt in options for opt in hyperlink_options):
                for page in doc:
                    page_state = PDFProcessor._collect_page_state(page)
                    if PDFProcessor._apply_hyperlink_actions(
                        doc,
                        page,
                        options,
                        file_like_link_kinds,
                        page_links=page_state["links"],
                    ):
                        changed = True
                        PDFProcessor._mark_change(applied_changes, "超链接动作已更新")
                    if PDFProcessor._apply_hyperlink_styles(
                        doc,
                        page,
                        options,
                        link_objs=page_state["link_objs"],
                        link_rects=page_state["link_rects"],
                    ):
                        changed = True
                        PDFProcessor._mark_change(applied_changes, "超链接外观已更新")

            cleanup_options = ["cleanup_remove_external_uri", "cleanup_remove_external_uri_and_text_black",
                               "cleanup_remove_invalid_links", "cleanup_remove_invalid_links_and_text_black",
                               "cleanup_remove_unknown_action_links",
                               "cleanup_remove_dynamic_content", "cleanup_remove_attachments", "cleanup_remove_tags", "cleanup_remove_annotations",
                               "cleanup_remove_metadata", "cleanup_remove_all_links_bookmarks"]
            if any(opt in options for opt in cleanup_options):
                external_uri_opts = {"cleanup_remove_external_uri", "cleanup_remove_external_uri_and_text_black"}
                selected_cleanup_opts = {opt for opt in options if opt in cleanup_options}

                # 性能快路径：仅删除外部 URI（可选去色）时，避免扫描注释和其他重逻辑
                if selected_cleanup_opts and selected_cleanup_opts.issubset(external_uri_opts):
                    for page in doc:
                        decolor_rects = []
                        removed_count = 0

                        for link in page.get_links():
                            if link.get("kind", fitz.LINK_NONE) != fitz.LINK_URI:
                                continue
                            if "cleanup_remove_external_uri_and_text_black" in options:
                                try:
                                    decolor_rects.append(fitz.Rect(link.get("from")))
                                except Exception:
                                    pass
                            try:
                                page.delete_link(link)
                                removed_count += 1
                                changed = True
                                PDFProcessor._mark_change(applied_changes, "已删除外部URI链接")
                                PDFProcessor._increase_change_count(change_counts, "已删除外部URI链接")
                            except Exception:
                                pass

                        # 仅在需要去色时触发内容流改色
                        if decolor_rects and "cleanup_remove_external_uri_and_text_black" in options:
                            if PDFProcessor._apply_text_color_via_content_stream(
                                doc,
                                page,
                                decolor_rects,
                                (0.0, 0.0, 0.0),
                                only_if_blue=True,
                            ):
                                changed = True
                                PDFProcessor._mark_change(applied_changes, "已将链接文本恢复为黑色")

                        # 兼容兜底：若仍有 URI 链接残留，再做一次注释级删除
                        if removed_count > 0 and any(
                            l.get("kind", fitz.LINK_NONE) == fitz.LINK_URI for l in page.get_links()
                        ):
                            for annot in page.annots() or []:
                                try:
                                    if annot.type[0] != 8:
                                        continue
                                    uri = getattr(annot, "uri", "") or ""
                                    if not uri and hasattr(annot, "info"):
                                        uri = annot.info.get("uri", "") or ""
                                    if uri:
                                        page.delete_annot(annot)
                                        changed = True
                                        PDFProcessor._mark_change(applied_changes, "已删除外部URI链接")
                                        PDFProcessor._increase_change_count(change_counts, "已删除外部URI链接")
                                except Exception:
                                    pass

                elif "cleanup_remove_all_links_bookmarks" in options:
                    doc.set_toc([])
                    for page in doc:
                        page_state = PDFProcessor._collect_page_state(page)
                        # 直接删除 Link 注释，避免部分 PDF 中 delete_link 命中不到
                        for annot in page_state["annots"]:
                            try:
                                if annot.type[0] == 8:  # 8 代表 LINK 注释
                                    page.delete_annot(annot)
                            except Exception:
                                pass
                        # 兜底：再按 get_links 删除一遍
                        for link in page_state["links"]:
                            try:
                                page.delete_link(link)
                            except Exception:
                                pass
                    changed = True
                    PDFProcessor._mark_change(applied_changes, "已删除全部链接和书签")
                else:
                    for page in doc:
                        page_state = PDFProcessor._collect_page_state(page)
                        decolor_rects = []

                        def _is_span_blue(span_color_int: int) -> bool:
                            # span["color"] 是 0xRRGGBB
                            b = span_color_int & 0xFF
                            g = (span_color_int >> 8) & 0xFF
                            r = (span_color_int >> 16) & 0xFF
                            return b > r + 40 and b > g + 40

                        def _overlay_black_text_in_rect(rect: fitz.Rect):
                            try:
                                text_dict = page.get_text("dict", clip=rect)
                            except Exception:
                                return
                            for block in text_dict.get("blocks", []):
                                for line in block.get("lines", []):
                                    for span in line.get("spans", []):
                                        try:
                                            txt = span.get("text", "")
                                            if not txt.strip():
                                                continue
                                            if not _is_span_blue(span.get("color", 0)):
                                                continue
                                            bbox = span.get("bbox", None)
                                            if not bbox or len(bbox) != 4:
                                                continue
                                            span_rect = fitz.Rect(bbox)
                                            # 叠加黑字覆盖蓝字（不重写内容流，尽量低风险）
                                            page.insert_textbox(
                                                span_rect,
                                                txt,
                                                fontsize=span.get("size", 11),
                                                fontname="helv",
                                                color=(0, 0, 0),
                                                overlay=True,
                                            )
                                            changed = True
                                        except Exception:
                                            continue

                        # 外部 URI 链接：优先用 delete_annot 方式确保真的移除可点击行为
                        if (
                            "cleanup_remove_external_uri" in options
                            or "cleanup_remove_external_uri_and_text_black" in options
                        ):
                            for annot in page_state["annots"]:
                                try:
                                    if annot.type[0] != 8:
                                        continue
                                    uri = ""
                                    # PyMuPDF 不同版本可能用不同字段暴露 uri
                                    if hasattr(annot, "uri"):
                                        uri = getattr(annot, "uri") or ""
                                    if not uri and hasattr(annot, "info"):
                                        uri = annot.info.get("uri", "") or ""
                                    if uri:
                                        if "cleanup_remove_external_uri_and_text_black" in options:
                                            decolor_rects.append(annot.rect)
                                        page.delete_annot(annot)
                                        changed = True
                                        PDFProcessor._mark_change(applied_changes, "已删除外部URI链接")
                                        PDFProcessor._increase_change_count(change_counts, "已删除外部URI链接")
                                except Exception:
                                    pass

                        for link in page_state["links"]:
                            kind = link.get("kind", fitz.LINK_NONE)
                            delete_it = False
                            if kind == fitz.LINK_URI and (
                                    "cleanup_remove_external_uri" in options or "cleanup_remove_external_uri_and_text_black" in options): delete_it = True
                            if kind == fitz.LINK_NONE and (
                                    "cleanup_remove_invalid_links" in options or "cleanup_remove_invalid_links_and_text_black" in options): delete_it = True
                            if "cleanup_remove_unknown_action_links" in options and kind not in [
                                fitz.LINK_GOTO, fitz.LINK_GOTOR, fitz.LINK_LAUNCH]: delete_it = True
                            if delete_it:
                                if kind == fitz.LINK_URI and "cleanup_remove_external_uri_and_text_black" in options:
                                    try:
                                        decolor_rects.append(fitz.Rect(link.get("from")))
                                    except Exception:
                                        pass
                                if kind == fitz.LINK_NONE and "cleanup_remove_invalid_links_and_text_black" in options:
                                    try:
                                        decolor_rects.append(fitz.Rect(link.get("from")))
                                    except Exception:
                                        pass
                                page.delete_link(link)
                                changed = True
                                if kind == fitz.LINK_URI:
                                    PDFProcessor._mark_change(applied_changes, "已删除外部URI链接")
                                    PDFProcessor._increase_change_count(change_counts, "已删除外部URI链接")
                                elif kind == fitz.LINK_NONE:
                                    PDFProcessor._mark_change(applied_changes, "已删除失效链接")
                                    PDFProcessor._increase_change_count(change_counts, "已删除失效链接")
                                else:
                                    PDFProcessor._mark_change(applied_changes, "已删除未知动作链接")
                                    PDFProcessor._increase_change_count(change_counts, "已删除未知动作链接")

                        # 去色：对刚刚删除的外部 URI 区域叠加黑色文字
                        if decolor_rects and (
                            "cleanup_remove_external_uri_and_text_black" in options
                            or "cleanup_remove_invalid_links_and_text_black" in options
                        ):
                            if PDFProcessor._apply_text_color_via_content_stream(
                                doc,
                                page,
                                decolor_rects,
                                (0.0, 0.0, 0.0),
                                only_if_blue=True,
                            ):
                                changed = True
                                PDFProcessor._mark_change(applied_changes, "已将链接文本恢复为黑色")

                if "cleanup_remove_annotations" in options:
                    for page in doc:
                        annots = list(page.annots() or [])
                        for annot in annots:
                            try:
                                page.delete_annot(annot)
                                changed = True
                                PDFProcessor._mark_change(applied_changes, "已删除PDF注释")
                                PDFProcessor._increase_change_count(change_counts, "已删除PDF注释")
                            except Exception:
                                pass

                if "cleanup_remove_dynamic_content" in options:
                    doc.xref_set_key(catalog_xref, "Names", "null");
                    changed = True
                    PDFProcessor._mark_change(applied_changes, "已删除动态内容/JavaScript")
                if "cleanup_remove_attachments" in options:
                    if doc.embfile_count() > 0:
                        attachment_count = doc.embfile_count()
                        for emb in doc.embfile_names(): doc.embfile_del(emb)
                        changed = True
                        PDFProcessor._mark_change(applied_changes, "已删除文档附件")
                        PDFProcessor._increase_change_count(change_counts, "已删除文档附件", attachment_count)
                if "cleanup_remove_tags" in options:
                    doc.xref_set_key(catalog_xref, "StructTreeRoot", "null")
                    doc.xref_set_key(catalog_xref, "MarkInfo", "null");
                    changed = True
                    PDFProcessor._mark_change(applied_changes, "已删除文档标签")
                if "cleanup_remove_metadata" in options:
                    doc.set_metadata({});
                    doc.xref_set_key(catalog_xref, "PieceInfo", "null");
                    changed = True
                    PDFProcessor._mark_change(applied_changes, "已删除文档元数据")

            is_linear = "fast_web_view" in options
            needs_gs_rewrite = needs_gs_engine or is_linear
            embed_fonts = "embed_nonstandard_fonts" in options

            if changed:
                if needs_gs_rewrite:
                    temp_pdf = str(output_path) + ".tmp.pdf"
                    doc.save(temp_pdf, garbage=3, deflate=True)
                    doc.close()
                    try:
                        PDFProcessor._rewrite_with_gs(
                            temp_pdf,
                            output_path,
                            embed_fonts=embed_fonts,
                            linearize=is_linear,
                        )
                        if embed_fonts:
                            PDFProcessor._mark_change(applied_changes, "已重写并嵌入字体")
                        if "convert_pdf_version" in options:
                            PDFProcessor._mark_change(applied_changes, "已转换PDF版本")
                        if is_linear:
                            PDFProcessor._mark_change(applied_changes, "已启用快速网页浏览")
                    finally:
                        if os.path.exists(temp_pdf):
                            os.remove(temp_pdf)
                else:
                    doc.save(output_path, garbage=3, deflate=True)
                    doc.close()
            else:
                doc.close()
                if needs_gs_rewrite:
                    PDFProcessor._rewrite_with_gs(
                        input_path,
                        output_path,
                        embed_fonts=embed_fonts,
                        linearize=is_linear,
                    )
                    if embed_fonts:
                        PDFProcessor._mark_change(applied_changes, "已重写并嵌入字体")
                    if "convert_pdf_version" in options:
                        PDFProcessor._mark_change(applied_changes, "已转换PDF版本")
                    if is_linear:
                        PDFProcessor._mark_change(applied_changes, "已启用快速网页浏览")
                else:
                    shutil.copy2(input_path, output_path)

            if applied_changes:
                return True, f"✅ 处理成功；修改项：{PDFProcessor._format_change_summary(change_counts, applied_changes)}"
            return True, "✅ 处理成功；无实际修改"

        except FileNotFoundError as e:
            return False, f"⚠️ 缺少引擎组件: {str(e)}"
        except Exception as e:
            return False, f"❌ 处理失败: {str(e)}"
