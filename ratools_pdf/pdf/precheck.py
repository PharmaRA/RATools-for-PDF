import os
import re
from pathlib import Path
from urllib.parse import unquote, urlparse

import fitz


def _processor_cls():
    from ratools_pdf.pdf.processor import PDFProcessor
    return PDFProcessor


def _option_title(option_id):
    return _processor_cls().PRECHECK_OPTION_TITLES.get(option_id, option_id)


def _precheck_option_matches_selected(option_id, selected_options):
    if selected_options is None:
        return True
    selected_options = set(selected_options or [])
    if option_id in selected_options:
        return True
    return bool(_processor_cls().PRECHECK_OPTION_ALIASES.get(option_id, set()) & selected_options)


def _selected_precheck_option_id(option_id, selected_options):
    if selected_options is None or option_id in set(selected_options or []):
        return option_id
    for alias in _processor_cls().PRECHECK_OPTION_ALIASES.get(option_id, set()):
        if alias in selected_options:
            return alias
    return option_id


def _filtered_precheck_options(selected_options):
    if selected_options is None:
        return None
    selected_options = set(selected_options or [])
    filtered = set()
    for option_id in selected_options:
        if option_id in _processor_cls().PRECHECK_DETECTABLE_OPTIONS:
            filtered.add(option_id)
    for option_id, aliases in _processor_cls().PRECHECK_OPTION_ALIASES.items():
        if option_id in selected_options or aliases & selected_options:
            filtered.add(option_id)
            filtered.update(aliases & selected_options)
    return filtered

PRECHECK_OPTION_TITLES = {
    "open_page_first": "设为首页打开",
    "page_layout_default": "重置页面布局",
    "zoom_default": "重置缩放比例",
    "initial_view_bookmarks_and_page": "设置导览标签",
    "collapse_all_bookmarks": "折叠所有书签",
    "title_from_filename": "同步文件名为标题",
    "bookmark_inherit_zoom": "书签设为承前缩放",
    "bookmark_open_new_window": "书签动作改为新窗口打开",
    "bookmark_remove_external_links": "删除书签中的外部链接",
    "bookmark_remove_invalid": "删除失效书签",
    "bookmark_remove_unknown_actions": "删除未知动作书签",
    "link_abs_to_rel_path": "外部文件链接转相对路径",
    "link_inherit_zoom": "超链接设为承前缩放",
    "link_open_new_window": "超链接动作改为新窗口打开",
    "link_text_blue": "链接文本设为蓝色",
    "link_black_border": "链接区域加黑框",
    "link_bordered_to_blue_border": "标准化有框链接",
    "link_unbordered_blue_to_blue_border": "标准化无框蓝字链接",
    "link_remove_border": "清除所有链接边框",
    "cleanup_remove_external_uri": "删除外部URI链接",
    "cleanup_remove_external_uri_and_text_black": "删除外部URI链接并去色",
    "cleanup_remove_invalid_links": "清理失效超链接",
    "cleanup_remove_invalid_links_and_text_black": "清理失效链接并去色",
    "cleanup_remove_unknown_action_links": "清理非标准动作链接",
    "cleanup_remove_dynamic_content": "彻底清除动态内容 (JS/3D)",
    "cleanup_remove_attachments": "移除所有内嵌附件",
    "cleanup_remove_tags": "移除结构化标签",
    "cleanup_remove_annotations": "清理所有高亮/批注",
    "cleanup_remove_metadata": "清空文档元数据",
    "cleanup_remove_all_links_bookmarks": "移除全部链接和书签",
    "convert_pdf_version": "PDF版本转换",
    "remove_pdf_restrictions": "PDF解除权限限制",
    "fast_web_view": "启用线性化 (快速网页浏览)",
}

PRECHECK_DETECTABLE_OPTIONS = {
    "open_page_first",
    "page_layout_default",
    "zoom_default",
    "initial_view_bookmarks_and_page",
    "collapse_all_bookmarks",
    "title_from_filename",
    "bookmark_inherit_zoom",
    "bookmark_open_new_window",
    "bookmark_remove_external_links",
    "bookmark_remove_invalid",
    "bookmark_remove_unknown_actions",
    "link_abs_to_rel_path",
    "link_inherit_zoom",
    "link_open_new_window",
    "cleanup_remove_external_uri",
    "cleanup_remove_external_uri_and_text_black",
    "cleanup_remove_invalid_links",
    "cleanup_remove_invalid_links_and_text_black",
    "cleanup_remove_unknown_action_links",
    "cleanup_remove_dynamic_content",
    "cleanup_remove_attachments",
    "cleanup_remove_tags",
    "cleanup_remove_annotations",
    "cleanup_remove_metadata",
    "cleanup_remove_all_links_bookmarks",
    "convert_pdf_version",
    "remove_pdf_restrictions",
    "fast_web_view",
}

PRECHECK_OPTION_ALIASES = {
    "cleanup_remove_external_uri": {"cleanup_remove_external_uri_and_text_black"},
    "cleanup_remove_invalid_links": {"cleanup_remove_invalid_links_and_text_black"},
}

NON_PROCESSING_OPTIONS = {"filename_ectd_format", "embed_nonstandard_fonts"}

BASE14_FONT_NAMES = {
    "Courier",
    "Courier-Bold",
    "Courier-Oblique",
    "Courier-BoldOblique",
    "Helvetica",
    "Helvetica-Bold",
    "Helvetica-Oblique",
    "Helvetica-BoldOblique",
    "Times-Roman",
    "Times-Bold",
    "Times-Italic",
    "Times-BoldItalic",
    "Symbol",
    "ZapfDingbats",
}


def _normalize_font_name(font_name):
    name = str(font_name or "").strip()
    if re.match(r"^[A-Z]{6}\+", name):
        name = name.split("+", 1)[1]
    return name


def _is_base14_font(font_name):
    return _processor_cls()._normalize_font_name(font_name) in _processor_cls().BASE14_FONT_NAMES


def _format_font_page_numbers(page_numbers):
    parsed_pages = set()
    for page in page_numbers:
        try:
            page_number = int(page)
        except (TypeError, ValueError):
            continue
        if page_number > 0:
            parsed_pages.add(page_number)

    pages = sorted(parsed_pages)
    if not pages:
        return "页码未知"

    ranges = []
    start = pages[0]
    prev = pages[0]
    for page in pages[1:]:
        if page == prev + 1:
            prev = page
            continue
        ranges.append(f"{start}-{prev}" if start != prev else str(start))
        start = page
        prev = page
    ranges.append(f"{start}-{prev}" if start != prev else str(start))
    return f"第{','.join(ranges)}页"


def _font_object_has_embedded_file(doc, xref, seen=None):
    if not xref:
        return False, False
    if seen is None:
        seen = set()
    if xref in seen:
        return False, True
    seen.add(xref)

    try:
        obj = doc.xref_object(int(xref))
    except Exception:
        return False, False

    if re.search(r"/FontFile(?:2|3)?\s+\d+\s+0\s+R", obj):
        return True, True

    refs = []
    refs.extend(int(match) for match in re.findall(r"/FontDescriptor\s+(\d+)\s+0\s+R", obj))
    refs.extend(int(match) for match in re.findall(r"/DescendantFonts\s*\[\s*(\d+)\s+0\s+R", obj))

    known = True
    for ref in refs:
        embedded, child_known = _processor_cls()._font_object_has_embedded_file(doc, ref, seen)
        known = known and child_known
        if embedded:
            return True, True
    return False, known


def _font_tuple_value(font_tuple, index):
    if len(font_tuple) <= index:
        return ""
    return font_tuple[index]


def _font_tuple_embedded_fallback(font_tuple):
    ext = str(_processor_cls()._font_tuple_value(font_tuple, 1) or "").strip().lower()
    if ext and ext not in {"n/a", "none", "null"}:
        return True, True
    if ext in {"n/a", "none", "null"}:
        return False, True
    return False, False


def _collect_font_precheck_findings(doc):
    fonts = {}
    for page_index, page in enumerate(doc, start=1):
        try:
            page_fonts = page.get_fonts(full=True)
        except Exception:
            continue

        for font_tuple in page_fonts:
            original_name = str(_processor_cls()._font_tuple_value(font_tuple, 3) or "").strip()
            if not original_name:
                original_name = str(_processor_cls()._font_tuple_value(font_tuple, 4) or "").strip()
            normalized_name = _processor_cls()._normalize_font_name(original_name)
            if not normalized_name:
                continue

            try:
                xref = int(_processor_cls()._font_tuple_value(font_tuple, 0) or 0)
            except Exception:
                xref = 0
            embedded, known = _processor_cls()._font_object_has_embedded_file(doc, xref)
            if not embedded and known:
                embedded, known = _processor_cls()._font_tuple_embedded_fallback(font_tuple)

            entry = fonts.setdefault(normalized_name, {
                "font_name": original_name,
                "original_names": set(),
                "normalized_name": normalized_name,
                "pages": set(),
                "has_unembedded": False,
                "has_embedded": False,
                "embedding_unknown": False,
                "base14": _processor_cls()._is_base14_font(normalized_name),
            })
            entry["original_names"].add(original_name)
            entry["pages"].add(page_index)
            if not known:
                entry["embedding_unknown"] = True
            elif embedded:
                entry["has_embedded"] = True
            else:
                entry["has_unembedded"] = True

    findings = []
    for normalized_name, entry in sorted(fonts.items(), key=lambda item: item[0].lower()):
        embedding_status_known = not entry["embedding_unknown"]
        has_unembedded = entry["has_unembedded"]
        has_unknown_embedding_status = entry["embedding_unknown"]
        embedded = entry["has_embedded"] and not entry["has_unembedded"] and embedding_status_known
        if entry["has_unembedded"]:
            embedded = False
        substitution_risk = bool((not entry["base14"]) and has_unembedded)

        if embedded and entry["base14"] and not entry["embedding_unknown"]:
            continue

        findings.append({
            "font_name": sorted(entry["original_names"])[0],
            "original_names": sorted(entry["original_names"]),
            "normalized_name": normalized_name,
            "pages": sorted(entry["pages"]),
            "embedded": embedded,
            "base14": entry["base14"],
            "substitution_risk": substitution_risk,
            "embedding_status_known": embedding_status_known,
            "has_unembedded": has_unembedded,
            "has_unknown_embedding_status": has_unknown_embedding_status,
        })

    unembedded_count = sum(1 for item in findings if item["has_unembedded"])
    non_base14_count = sum(1 for item in findings if not item["base14"])
    risk_count = sum(1 for item in findings if item["substitution_risk"])
    unknown_count = sum(1 for item in findings if item["has_unknown_embedding_status"])

    summary_parts = []
    if unembedded_count:
        summary_parts.append(f"未嵌入字体 {unembedded_count} 个")
    if non_base14_count:
        summary_parts.append(f"非标准字体 {non_base14_count} 个")
    if risk_count:
        summary_parts.append(f"替代字体风险 {risk_count} 个")
    if unknown_count:
        summary_parts.append(f"嵌入状态未知字体 {unknown_count} 个")

    detail_parts = []
    for item in findings:
        labels = []
        if item["has_unembedded"]:
            labels.append("未嵌入")
        elif item["embedded"]:
            labels.append("已嵌入")
        else:
            labels.append("嵌入状态未知")
        if item["has_unknown_embedding_status"] and item["has_unembedded"]:
            labels.append("嵌入状态未知")
        labels.append("Base14" if item["base14"] else "非Base14")
        if item["substitution_risk"]:
            labels.append("替代风险")
        detail_parts.append(
            f"{item['normalized_name']}({_processor_cls()._format_font_page_numbers(item['pages'])}，{'，'.join(labels)})"
        )

    return {
        "font_summary": "，".join(summary_parts),
        "font_details": "; ".join(detail_parts),
        "font_findings": findings,
    }


def _font_precheck_has_embedding_risk(font_precheck):
    for item in font_precheck.get("font_findings", []) or []:
        if (
            item.get("has_unembedded")
            or item.get("substitution_risk")
            or item.get("has_unknown_embedding_status")
        ):
            return True
    return False


def _collect_font_precheck_for_path(pdf_path):
    doc = None
    try:
        doc = fitz.open(pdf_path)
        if doc.needs_pass:
            return {
                "available": False,
                "error": "该PDF需要密码，无法执行字体风险预检",
                "font_summary": "",
                "font_details": "",
                "font_findings": [],
            }
        font_precheck = _processor_cls()._collect_font_precheck_findings(doc)
        font_precheck["available"] = True
        font_precheck["error"] = ""
        return font_precheck
    except Exception as exc:
        return {
            "available": False,
            "error": str(exc),
            "font_summary": "",
            "font_details": "",
            "font_findings": [],
        }
    finally:
        if doc is not None:
            try:
                doc.close()
            except Exception:
                pass


def _add_precheck_suggestion(suggestions, option_id, reason):
    if option_id in suggestions:
        return
    suggestions[option_id] = {
        "matched": True,
        "title": _processor_cls().PRECHECK_OPTION_TITLES.get(option_id, option_id),
        "reason": reason,
    }


def _add_precheck_report_finding(suggestions, finding_id, title, reason):
    if finding_id in suggestions:
        existing_reason = suggestions[finding_id].get("reason", "")
        if reason and reason not in existing_reason:
            suggestions[finding_id]["reason"] = f"{existing_reason}；{reason}" if existing_reason else reason
        return
    suggestions[finding_id] = {
        "matched": True,
        "title": title,
        "reason": reason,
        "report_only": True,
    }


def _resolve_external_file_target(base_dir, file_path):
    raw_path = str(file_path or "").strip()
    if not raw_path:
        return "", "", False

    parsed = urlparse(unquote(raw_path))
    if parsed.scheme.lower() == "file":
        if parsed.netloc:
            file_part = f"//{parsed.netloc}{parsed.path}"
        else:
            file_part = parsed.path
    else:
        file_part = raw_path.split("#", 1)[0]

    decoded_path = unquote(file_part).strip()
    if not decoded_path:
        return "", "", False
    if re.match(r"^/[A-Za-z]:[\\/]", decoded_path):
        decoded_path = decoded_path[1:]

    is_absolute = (
        os.path.isabs(decoded_path)
        or bool(re.match(r"^[A-Za-z]:[\\/]", decoded_path))
        or decoded_path.startswith("\\\\")
        or decoded_path.startswith("//")
    )
    normalized_path = os.path.normpath(decoded_path.replace("\\", os.sep).replace("/", os.sep))
    target_path = normalized_path if is_absolute else os.path.normpath(os.path.join(base_dir, normalized_path))
    return decoded_path, target_path, not is_absolute


def _read_target_pdf_page_count(target_path):
    target_doc = None
    try:
        target_doc = fitz.open(target_path)
        if target_doc.needs_pass:
            return None
        return target_doc.page_count
    except Exception:
        return None
    finally:
        if target_doc is not None:
            try:
                target_doc.close()
            except Exception:
                pass


def _add_link_target_integrity_findings(suggestions, base_dir, kind, file_path, target_page=None, source_label="链接"):
    link_file_kind = getattr(fitz, "LINK_FILE", None)
    is_file_link = kind in [fitz.LINK_GOTOR, fitz.LINK_LAUNCH] or (
        link_file_kind is not None and kind == link_file_kind
    )
    if not is_file_link:
        return

    if not str(file_path or "").strip():
        finding_id = (
            "link_target_integrity_gotor_missing_file"
            if kind == fitz.LINK_GOTOR
            else "link_target_integrity_missing_file"
        )
        title = (
            "链接目标完整性检查：GoToR目标文件不存在"
            if kind == fitz.LINK_GOTOR
            else "链接目标完整性检查：外部文件链接目标不存在"
        )
        _processor_cls()._add_precheck_report_finding(
            suggestions,
            finding_id,
            title,
            f"{source_label} 未指定目标文件",
        )
        return

    decoded_path, target_path, is_relative = _processor_cls()._resolve_external_file_target(base_dir, file_path)
    if not decoded_path:
        return

    if not os.path.isfile(target_path):
        if kind == fitz.LINK_GOTOR:
            _processor_cls()._add_precheck_report_finding(
                suggestions,
                "link_target_integrity_gotor_missing_file",
                "链接目标完整性检查：GoToR目标文件不存在",
                f"{source_label} 指向的目标文件不存在: {decoded_path}",
            )
        else:
            _processor_cls()._add_precheck_report_finding(
                suggestions,
                "link_target_integrity_missing_file",
                "链接目标完整性检查：外部文件链接目标不存在",
                f"{source_label} 指向的目标文件不存在: {decoded_path}",
            )
        if is_relative:
            _processor_cls()._add_precheck_report_finding(
                suggestions,
                "link_target_integrity_broken_relative_path",
                "链接目标完整性检查：eCTD相对路径链接断链",
                f"{source_label} 指向的相对路径不存在: {decoded_path}",
            )
        return

    if kind != fitz.LINK_GOTOR:
        return

    page_count = _processor_cls()._read_target_pdf_page_count(target_path)
    if page_count is None:
        return
    try:
        target_page_index = int(target_page if target_page is not None else 0)
    except Exception:
        target_page_index = -1
    if target_page_index < 0 or target_page_index >= page_count:
        _processor_cls()._add_precheck_report_finding(
            suggestions,
            "link_target_integrity_gotor_page_out_of_range",
            "链接目标完整性检查：GoToR目标页码越界",
            f"{source_label} 指向 {decoded_path} 的第 {target_page_index + 1} 页，超过目标文件页数 {page_count}",
        )


def _catalog_key(doc, catalog_xref, key):
    try:
        return doc.xref_get_key(catalog_xref, key)
    except Exception:
        return ("null", "null")


def _catalog_key_is_present(doc, catalog_xref, key):
    kind, value = _processor_cls()._catalog_key(doc, catalog_xref, key)
    return kind != "null" and value != "null"


def _dereference_xref_value(doc, value):
    match = re.match(r"^(\d+)\s+0\s+R$", str(value or "").strip())
    if not match:
        return str(value or "")
    try:
        return doc.xref_object(int(match.group(1)))
    except Exception:
        return str(value or "")


def _catalog_key_resolved_value(doc, catalog_xref, key):
    _kind, value = _processor_cls()._catalog_key(doc, catalog_xref, key)
    return _processor_cls()._dereference_xref_value(doc, value)


def _pdf_has_signature(input_path):
    """判断 PDF 是否已包含数字签名。

    只读检测，任何异常都按“无签名”处理，避免阻塞批量处理流程。
    """
    if not input_path or not os.path.isfile(input_path):
        return False

    doc = None
    try:
        doc = fitz.open(input_path)
        if doc.needs_pass:
            # 加密文档无法可靠读取签名域，交由后续处理流程按加密文件处理
            return False

        # 首选：AcroForm 的 SigFlags。位 1 (SignaturesExist) 置位表示文档存在签名域，
        # 常规监管 PDF 中该标志置位基本等同于已签名。
        try:
            if doc.get_sigflags() > 0:
                return True
        except Exception:
            pass

        # 兜底：扫描签名控件，命中实际已签名的签名域。
        signature_widget_type = getattr(fitz, "PDF_WIDGET_TYPE_SIGNATURE", None)
        for page in doc:
            for widget in (page.widgets() or []):
                try:
                    if signature_widget_type is not None and widget.field_type != signature_widget_type:
                        continue
                    if getattr(widget, "is_signed", None):
                        return True
                    if str(getattr(widget, "field_value", "") or "").strip():
                        return True
                except Exception:
                    continue
        return False
    except Exception:
        return False
    finally:
        if doc is not None:
            try:
                doc.close()
            except Exception:
                pass


# 批注类型编号到中文名称的映射。用于批注复核提示的明细展示。
# 注意：PyMuPDF 的 page.annots() 本身不返回 Link 注释，链接由 get_links() 处理，
# 因此这里以类型编号判定，不依赖易混淆的“type[0] == 8”写法。
ANNOTATION_TYPE_LABELS = {
    0: "便签(文本注释)",
    2: "自由文本",
    3: "直线",
    4: "矩形",
    5: "圆形",
    6: "多边形",
    7: "折线",
    8: "高亮",
    9: "下划线",
    10: "波浪线",
    11: "删除线",
    12: "涂黑(密文)",
    13: "图章",
    14: "标注",
    15: "手绘",
    17: "文件附件",
    18: "声音",
    19: "影片",
    22: "屏幕",
}

# Word 转 PDF 时，交叉引用、书签或超链接失效会残留固定的占位错误文本。
# 这里只匹配 Word 生成的固定文案，不匹配正文中正常出现的“错误”“error”，避免误报。
WORD_BROKEN_REFERENCE_REGEXES = [
    re.compile(r"错误\s*[！!]\s*未找到引用源"),
    re.compile(r"错误\s*[！!]\s*未定义书签"),
    re.compile(r"错误\s*[！!]\s*超链接引用无效"),
    re.compile(r"错误\s*[！!]\s*链接无效"),
    re.compile(r"错误\s*[！!]\s*不是有效的(?:文件名|链接)"),
    re.compile(r"错误\s*[！!]\s*文档中没有指定样式的文字"),
    re.compile(r"错误\s*[！!]\s*不能通过编辑域代码创建对象"),
    re.compile(r"Error!\s*Reference source not found", re.IGNORECASE),
    re.compile(r"Error!\s*Bookmark not defined", re.IGNORECASE),
    re.compile(r"Error!\s*Hyperlink reference not valid", re.IGNORECASE),
    re.compile(r"Error!\s*No text of specified style in document", re.IGNORECASE),
    re.compile(r"Error!\s*Objects cannot be created from editing field codes", re.IGNORECASE),
    re.compile(r"Error!\s*Not a valid (?:filename|link)", re.IGNORECASE),
]


def _collect_annotation_findings(doc):
    """扫描 PDF 中的批注（便签、高亮、下划线等），仅作人工复核提示，不自动处理。

    统计非链接类批注（PyMuPDF 的 page.annots() 本身不返回 Link 注释）；
    Popup 依附于其它批注，不单独计数。返回结构包含总数、页数、按类型的明细。
    """
    link_type = getattr(fitz, "PDF_ANNOT_LINK", 1)
    popup_type = getattr(fitz, "PDF_ANNOT_POPUP", 16)

    type_counts = {}
    matched_pages = set()
    total = 0
    for page_index, page in enumerate(doc, start=1):
        try:
            annots = page.annots()
        except Exception:
            continue
        if not annots:
            continue
        for annot in annots:
            try:
                type_tuple = annot.type
            except Exception:
                continue
            if isinstance(type_tuple, (tuple, list)) and type_tuple:
                type_number = type_tuple[0]
                type_name = type_tuple[1] if len(type_tuple) > 1 else ""
            else:
                type_number = None
                type_name = ""
            if type_number in (link_type, popup_type):
                continue
            total += 1
            matched_pages.add(page_index)
            label = ANNOTATION_TYPE_LABELS.get(type_number, str(type_name or "批注"))
            type_counts[label] = type_counts.get(label, 0) + 1

    if total == 0:
        return {"has_annotations": False, "count": 0, "summary": "", "details": ""}

    detail_parts = [
        f"{label} {count} 个"
        for label, count in sorted(type_counts.items(), key=lambda item: item[0])
    ]
    return {
        "has_annotations": True,
        "count": total,
        "summary": f"发现 {total} 个批注，分布在 {len(matched_pages)} 页",
        "details": "、".join(detail_parts),
    }


def _collect_broken_reference_findings(doc):
    """扫描 Word 转 PDF 后残留的失效引用/链接占位错误文本。

    仅匹配 Word 固定生成的错误文案（如“错误！未找到引用源。”“Error! Reference source not found.”），
    不匹配正文中正常出现的“错误”“error”，避免误报。返回结构包含命中页码与命中文案样例。
    """
    matched_pages = {}  # page_number -> set(命中文案)
    for page_index, page in enumerate(doc, start=1):
        try:
            text = page.get_text("text")
        except Exception:
            continue
        if not text:
            continue
        for regex in WORD_BROKEN_REFERENCE_REGEXES:
            match = regex.search(text)
            if match:
                matched_pages.setdefault(page_index, set()).add(match.group(0).strip())

    if not matched_pages:
        return {"has_broken_reference": False, "count": 0, "summary": "", "details": ""}

    pages = sorted(matched_pages)
    detail_parts = [
        f"第{page_number}页：{'、'.join(sorted(matched_pages[page_number]))}"
        for page_number in pages
    ]
    return {
        "has_broken_reference": True,
        "count": len(pages),
        "summary": f"发现 {len(pages)} 页存在疑似失效引用/链接占位文本（Word 转 PDF 常见问题）",
        "details": "; ".join(detail_parts),
    }


def _collect_annotation_findings_for_path(pdf_path):
    """按路径执行批注检测，供独立的“批注检测”按钮使用，不牵连完整预检。"""
    if not os.path.exists(pdf_path):
        return {"available": False, "error": "文件不存在", "has_annotations": False, "count": 0, "summary": "", "details": ""}
    if not os.path.isfile(pdf_path):
        return {"available": False, "error": "不是PDF文件", "has_annotations": False, "count": 0, "summary": "", "details": ""}

    doc = None
    try:
        doc = fitz.open(pdf_path)
        if doc.needs_pass:
            return {"available": False, "error": "该PDF需要密码，无法执行批注检测", "has_annotations": False, "count": 0, "summary": "", "details": ""}
        findings = _processor_cls()._collect_annotation_findings(doc)
        findings["available"] = True
        findings["error"] = ""
        return findings
    except Exception as exc:
        return {"available": False, "error": str(exc), "has_annotations": False, "count": 0, "summary": "", "details": ""}
    finally:
        if doc is not None:
            try:
                doc.close()
            except Exception:
                pass


def _collect_broken_reference_findings_for_path(pdf_path):
    """按路径执行失效引用/链接文本检测，供独立的“失效引用检测”按钮使用，不牵连完整预检。"""
    if not os.path.exists(pdf_path):
        return {"available": False, "error": "文件不存在", "has_broken_reference": False, "count": 0, "summary": "", "details": ""}
    if not os.path.isfile(pdf_path):
        return {"available": False, "error": "不是PDF文件", "has_broken_reference": False, "count": 0, "summary": "", "details": ""}

    doc = None
    try:
        doc = fitz.open(pdf_path)
        if doc.needs_pass:
            return {"available": False, "error": "该PDF需要密码，无法执行失效引用检测", "has_broken_reference": False, "count": 0, "summary": "", "details": ""}
        findings = _processor_cls()._collect_broken_reference_findings(doc)
        findings["available"] = True
        findings["error"] = ""
        return findings
    except Exception as exc:
        return {"available": False, "error": str(exc), "has_broken_reference": False, "count": 0, "summary": "", "details": ""}
    finally:
        if doc is not None:
            try:
                doc.close()
            except Exception:
                pass


def build_precheck_report(input_path, selected_options=None):
    selected_options = _processor_cls()._filtered_precheck_options(selected_options)

    def wants(*option_ids):
        if selected_options is None:
            return True
        return any(_processor_cls()._precheck_option_matches_selected(option_id, selected_options) for option_id in option_ids)

    def add_suggestion(option_id, reason):
        if not wants(option_id):
            return
        suggestion_id = _processor_cls()._selected_precheck_option_id(option_id, selected_options)
        _processor_cls()._add_precheck_suggestion(suggestions, suggestion_id, reason)

    report = {
        "available": False,
        "file_path": input_path,
        "file_name": os.path.basename(input_path),
        "suggestions": {},
        "font_summary": "",
        "font_details": "",
        "font_findings": [],
        "annotation_summary": "",
        "annotation_details": "",
        "broken_reference_summary": "",
        "broken_reference_details": "",
        "error": "",
    }

    if not os.path.exists(input_path):
        report["error"] = "文件不存在"
        return report
    if not os.path.isfile(input_path):
        report["error"] = "不是PDF文件"
        return report

    suggestions = report["suggestions"]
    doc = None
    try:
        doc = fitz.open(input_path)
        if doc.needs_pass:
            report["error"] = "文件需要打开密码，无法预检内部结构"
            return report

        report["available"] = True
        base_dir = os.path.dirname(os.path.abspath(input_path))
        catalog_xref = doc.pdf_catalog()
        toc = doc.get_toc(simple=False)
        has_bookmarks = len(toc) > 0
        meta = doc.metadata or {}
        base_name = Path(input_path).stem

        if wants("title_from_filename") and (meta.get("title") or "") != base_name:
            add_suggestion("title_from_filename", "PDF标题属性与文件名不一致或为空")

        if wants("initial_view_bookmarks_and_page"):
            page_mode_kind, page_mode_value = _processor_cls()._catalog_key(doc, catalog_xref, "PageMode")
            if has_bookmarks and not (page_mode_kind == "name" and page_mode_value == "/UseOutlines"):
                add_suggestion("initial_view_bookmarks_and_page", "文档包含书签，但打开时未设置为显示书签面板")
            elif not has_bookmarks and page_mode_kind == "name" and page_mode_value != "/UseNone":
                add_suggestion("initial_view_bookmarks_and_page", "文档不含书签，但初始导览标签不是页面视图")

        if wants("page_layout_default") and _processor_cls()._catalog_key_is_present(doc, catalog_xref, "PageLayout"):
            add_suggestion("page_layout_default", "文档设置了显式页面布局")

        if wants("open_page_first", "zoom_default"):
            open_action_kind, open_action_value = _processor_cls()._catalog_key(doc, catalog_xref, "OpenAction")
        else:
            open_action_kind, open_action_value = "null", "null"
        if open_action_kind != "null" and doc.page_count > 0:
            open_action_value = _processor_cls()._dereference_xref_value(doc, open_action_value)
            first_page_ref = f"{doc[0].xref} 0 R"
            compact_action = open_action_value.replace(" ", "")
            if wants("open_page_first") and first_page_ref.replace(" ", "") not in compact_action:
                add_suggestion("open_page_first", "文档打开动作没有指向第一页")
            if wants("zoom_default") and "/XYZnullnullnull" not in compact_action:
                add_suggestion("zoom_default", "文档打开动作使用了固定缩放或非默认视图")

        bookmark_wanted = wants(
            "collapse_all_bookmarks",
            "bookmark_inherit_zoom",
            "bookmark_open_new_window",
            "bookmark_remove_external_links",
            "bookmark_remove_invalid",
            "bookmark_remove_unknown_actions",
            "cleanup_remove_all_links_bookmarks",
        )
        if has_bookmarks and bookmark_wanted:
            if selected_options is not None and wants("cleanup_remove_all_links_bookmarks"):
                add_suggestion("cleanup_remove_all_links_bookmarks", "文档包含可清理的书签")
            if wants("collapse_all_bookmarks") and any(isinstance(item[-1], dict) and item[-1].get("collapse") is not True for item in toc):
                add_suggestion("collapse_all_bookmarks", "文档包含未折叠的书签")

            for item in toc:
                try:
                    _level, _title, bm_page, dest = item
                except Exception:
                    continue
                if not isinstance(dest, dict):
                    dest = {}
                kind = dest.get("kind", fitz.LINK_NONE)
                if wants("bookmark_inherit_zoom") and kind == fitz.LINK_GOTO and dest.get("zoom", 0.0) not in [0, 0.0, None]:
                    add_suggestion("bookmark_inherit_zoom", "部分内部书签使用了固定缩放比例")
                if wants("bookmark_open_new_window") and kind in [fitz.LINK_GOTOR, fitz.LINK_LAUNCH] and not dest.get("newWindow"):
                    add_suggestion("bookmark_open_new_window", "部分外部文件书签未设置为新窗口打开")
                if selected_options is None:
                    _processor_cls()._add_link_target_integrity_findings(
                        suggestions,
                        base_dir,
                        kind,
                        dest.get("file", ""),
                        dest.get("page", 0),
                        "书签",
                    )
                if wants("bookmark_remove_external_links") and kind == fitz.LINK_URI:
                    add_suggestion("bookmark_remove_external_links", "书签中包含外部URI链接")
                if wants("bookmark_remove_invalid") and (
                    kind == fitz.LINK_NONE or (kind == fitz.LINK_GOTO and (bm_page < 1 or bm_page > doc.page_count))
                ):
                    add_suggestion("bookmark_remove_invalid", "书签中存在失效目标")
                if wants("bookmark_remove_unknown_actions") and kind not in [fitz.LINK_GOTO, fitz.LINK_GOTOR, fitz.LINK_LAUNCH]:
                    add_suggestion("bookmark_remove_unknown_actions", "书签中存在非标准动作")

        link_file_kind = getattr(fitz, "LINK_FILE", None)
        file_like_link_kinds = {fitz.LINK_GOTOR}
        if link_file_kind is not None:
            file_like_link_kinds.add(link_file_kind)

        has_non_link_annotation = False
        link_wanted = wants(
            "link_abs_to_rel_path",
            "link_inherit_zoom",
            "link_open_new_window",
            "cleanup_remove_external_uri",
            "cleanup_remove_external_uri_and_text_black",
            "cleanup_remove_invalid_links",
            "cleanup_remove_invalid_links_and_text_black",
            "cleanup_remove_unknown_action_links",
            "cleanup_remove_all_links_bookmarks",
            "cleanup_remove_annotations",
        )
        if link_wanted:
            for page in doc:
                links = page.get_links()
                if selected_options is not None and wants("cleanup_remove_all_links_bookmarks") and links:
                    add_suggestion("cleanup_remove_all_links_bookmarks", "页面中包含可清理的链接")
                for link in links:
                    kind = link.get("kind", fitz.LINK_NONE)
                    if wants("link_abs_to_rel_path") and kind in file_like_link_kinds:
                        file_path = link.get("file", "") or ""
                        decoded_file_path = unquote(file_path)
                        if decoded_file_path and (
                            ":" in decoded_file_path
                            or decoded_file_path.startswith("/")
                            or decoded_file_path.startswith("\\")
                        ):
                            add_suggestion("link_abs_to_rel_path", "外部文件链接中包含绝对路径")
                    if wants("link_inherit_zoom") and kind == fitz.LINK_GOTO and link.get("zoom", 0.0) not in [0, 0.0, None]:
                        add_suggestion("link_inherit_zoom", "部分内部超链接使用了固定缩放比例")
                    if wants("link_open_new_window") and kind in [fitz.LINK_GOTOR, fitz.LINK_LAUNCH] and not link.get("newWindow"):
                        add_suggestion("link_open_new_window", "部分外部文件链接未设置为新窗口打开")
                    if selected_options is None:
                        _processor_cls()._add_link_target_integrity_findings(
                            suggestions,
                            base_dir,
                            kind,
                            link.get("file", ""),
                            link.get("page", 0),
                            "页面链接",
                        )
                    if wants("cleanup_remove_external_uri") and kind == fitz.LINK_URI:
                        add_suggestion("cleanup_remove_external_uri", "页面中包含外部URI链接")
                    if wants("cleanup_remove_invalid_links") and kind == fitz.LINK_NONE:
                        add_suggestion("cleanup_remove_invalid_links", "页面中存在无有效动作的链接区域")
                    if wants("cleanup_remove_unknown_action_links") and kind not in [fitz.LINK_GOTO, fitz.LINK_GOTOR, fitz.LINK_LAUNCH]:
                        add_suggestion("cleanup_remove_unknown_action_links", "页面中存在非标准链接动作")

                for annot in page.annots() or []:
                    try:
                        if annot.type[0] == 8:
                            uri = getattr(annot, "uri", "") or ""
                            if not uri and hasattr(annot, "info"):
                                uri = annot.info.get("uri", "") or ""
                            if wants("cleanup_remove_external_uri") and uri:
                                add_suggestion("cleanup_remove_external_uri", "页面注释中包含外部URI链接")
                        else:
                            has_non_link_annotation = True
                    except Exception:
                        continue

        if wants("cleanup_remove_annotations") and has_non_link_annotation:
            add_suggestion("cleanup_remove_annotations", "文档包含高亮、文本框或其它批注")

        if wants("cleanup_remove_attachments") and doc.embfile_count() > 0:
            add_suggestion("cleanup_remove_attachments", f"文档包含 {doc.embfile_count()} 个内嵌附件")

        if wants("cleanup_remove_tags") and (
            _processor_cls()._catalog_key_is_present(doc, catalog_xref, "StructTreeRoot")
            or _processor_cls()._catalog_key_is_present(doc, catalog_xref, "MarkInfo")
        ):
            add_suggestion("cleanup_remove_tags", "文档包含结构化标签信息")

        if wants("cleanup_remove_metadata"):
            metadata_values = [
                value
                for key, value in meta.items()
                if key not in ["format", "encryption"] and str(value or "").strip()
            ]
            if metadata_values or _processor_cls()._catalog_key_is_present(doc, catalog_xref, "PieceInfo"):
                add_suggestion("cleanup_remove_metadata", "文档包含可清理的元数据")

        if wants("cleanup_remove_dynamic_content"):
            catalog_object = ""
            try:
                catalog_object = doc.xref_object(catalog_xref)
            except Exception:
                catalog_object = ""
            names_kind, _names_value = _processor_cls()._catalog_key(doc, catalog_xref, "Names")
            names_value = _processor_cls()._catalog_key_resolved_value(doc, catalog_xref, "Names")
            dynamic_probe = f"{catalog_object}\n{names_value}"
            if names_kind != "null" and any(marker in dynamic_probe for marker in ["/JavaScript", "/JS", "/RichMedia", "/3D"]):
                add_suggestion("cleanup_remove_dynamic_content", "文档包含 JavaScript、3D 或富媒体入口")

        if wants("convert_pdf_version"):
            pdf_version = _processor_cls()._read_pdf_header_version(input_path)
            if pdf_version and pdf_version != "1.7":
                add_suggestion("convert_pdf_version", f"当前PDF版本为 {pdf_version}，不是 1.7")

        if wants("fast_web_view") and not _processor_cls()._is_pdf_linearized(input_path):
            add_suggestion("fast_web_view", "文档未启用线性化快速网页浏览")

        if wants("remove_pdf_restrictions") and _processor_cls()._qpdf_reports_restrictions(input_path):
            add_suggestion("remove_pdf_restrictions", "文档存在打印、复制或编辑权限限制")

        if selected_options is None:
            try:
                font_precheck = _processor_cls()._collect_font_precheck_findings(doc)
                report["font_summary"] = font_precheck.get("font_summary", "")
                report["font_details"] = font_precheck.get("font_details", "")
                report["font_findings"] = font_precheck.get("font_findings", [])
                if report["font_summary"]:
                    reason = report["font_summary"]
                    if report["font_details"]:
                        reason = f"{reason}；明细：{report['font_details']}"
                    _processor_cls()._add_precheck_report_finding(
                        suggestions,
                        "font_precheck_review",
                        "字体预检：需要复核",
                        reason,
                    )
            except Exception:
                pass

            try:
                annotation_precheck = _processor_cls()._collect_annotation_findings(doc)
                if annotation_precheck.get("has_annotations"):
                    report["annotation_summary"] = annotation_precheck.get("summary", "")
                    report["annotation_details"] = annotation_precheck.get("details", "")
                    reason = report["annotation_summary"]
                    if report["annotation_details"]:
                        reason = f"{reason}；明细：{report['annotation_details']}"
                    _processor_cls()._add_precheck_report_finding(
                        suggestions,
                        "annotation_precheck_review",
                        "批注检查：需要复核",
                        reason,
                    )
            except Exception:
                pass

            try:
                broken_reference_precheck = _processor_cls()._collect_broken_reference_findings(doc)
                if broken_reference_precheck.get("has_broken_reference"):
                    report["broken_reference_summary"] = broken_reference_precheck.get("summary", "")
                    report["broken_reference_details"] = broken_reference_precheck.get("details", "")
                    reason = report["broken_reference_summary"]
                    if report["broken_reference_details"]:
                        reason = f"{reason}；明细：{report['broken_reference_details']}"
                    _processor_cls()._add_precheck_report_finding(
                        suggestions,
                        "broken_reference_precheck_review",
                        "失效引用/链接文本检查：需要复核",
                        reason,
                    )
            except Exception:
                pass

        return report
    except Exception as e:
        report["available"] = False
        report["error"] = str(e)
        return report
    finally:
        if doc is not None:
            try:
                doc.close()
            except Exception:
                pass


def resolve_processing_options(input_path, options, processing_mode="smart"):
    selected_options = set(options or [])
    selected_options.discard("embed_nonstandard_fonts")
    mode = str(processing_mode or "smart").lower()
    if mode not in {"smart", "force"}:
        mode = "smart"

    if mode == "force":
        return {
            "mode": "force",
            "options": selected_options,
            "skipped": [],
            "forced_unsupported": [],
            "suggested": [],
            "precheck_available": None,
            "log": "处理模式: 全部处理（强制执行全部勾选规则）",
        }

    actionable_selected = selected_options - _processor_cls().NON_PROCESSING_OPTIONS
    unsupported = actionable_selected - _processor_cls().PRECHECK_DETECTABLE_OPTIONS
    report = _processor_cls().build_precheck_report(input_path, selected_options=actionable_selected)
    if not report.get("available"):
        reason = report.get("error") or "预检不可用"
        return {
            "mode": "smart",
            "options": selected_options,
            "skipped": [],
            "forced_unsupported": sorted(unsupported),
            "suggested": [],
            "precheck_available": False,
            "log": f"处理模式: 智能处理；预检不可用，执行全部已勾选规则。原因: {reason}",
        }

    suggested = {
        option_id
        for option_id, item in report.get("suggestions", {}).items()
        if not item.get("report_only")
    }
    effective = set(selected_options & _processor_cls().NON_PROCESSING_OPTIONS)
    effective.update(actionable_selected & suggested)
    effective.update(unsupported)

    skipped = sorted(
        option_id
        for option_id in actionable_selected
        if option_id in _processor_cls().PRECHECK_DETECTABLE_OPTIONS and option_id not in effective
    )
    forced_unsupported = sorted(unsupported)
    log_parts = ["处理模式: 智能处理（仅处理预检发现的问题）"]
    if suggested:
        log_parts.append(
            "预检命中: " + "、".join(_processor_cls()._option_title(option_id) for option_id in sorted(suggested))
        )
    else:
        log_parts.append("预检命中: 无")
    if forced_unsupported:
        log_parts.append(
            "无法可靠预检但已执行: "
            + "、".join(_processor_cls()._option_title(option_id) for option_id in forced_unsupported)
        )
    if skipped:
        log_parts.append(
            "已跳过未命中规则: "
            + "、".join(_processor_cls()._option_title(option_id) for option_id in skipped)
        )
    return {
        "mode": "smart",
        "options": effective,
        "skipped": skipped,
        "forced_unsupported": forced_unsupported,
        "suggested": sorted(suggested),
        "precheck_available": True,
        "log": "；".join(log_parts),
    }


