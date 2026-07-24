import fitz
import os
import sys
import subprocess
import shutil
import csv
import json
import re
import time
from urllib.parse import unquote, urlparse
from pathlib import Path

from ratools_pdf.config.paths import get_resource_path
from ratools_pdf.pdf.font_embedding_providers import get_font_embedding_provider
from ratools_pdf.pdf import bookmarks_links, hyperlink_styles, page_layout, precheck, qpdf


class _PhaseProfiler:
    """轻量阶段计时器。通过环境变量 RATOOLS_PROFILE=1 开启。

    用法:
        prof = _PhaseProfiler(page_count)
        with prof.phase("预检"):
            ...
        msg = prof.summary()  # 关闭时返回空字符串
    """

    def __init__(self, page_count=0):
        self.enabled = str(os.environ.get("RATOOLS_PROFILE", "")).strip().lower() in ("1", "true", "yes", "on")
        self.page_count = page_count
        self._phases = []  # [(name, seconds), ...]
        self._t0 = time.perf_counter() if self.enabled else 0.0
        self._tag = f"[PID {os.getpid()}]"

    def _log(self, text):
        # 即时打印到 stderr 并 flush，保证卡住时也能看到已完成/正在进行的阶段
        try:
            sys.stderr.write(text + "\n")
            sys.stderr.flush()
        except Exception:
            pass

    class _Timer:
        def __init__(self, profiler, name):
            self.profiler = profiler
            self.name = name
            self.start = 0.0

        def __enter__(self):
            if self.profiler.enabled:
                self.start = time.perf_counter()
                self.profiler._log(f"⏱{self.profiler._tag} ▶ 开始: {self.name}")
            return self

        def __exit__(self, exc_type, exc, tb):
            if self.profiler.enabled:
                elapsed = time.perf_counter() - self.start
                self.profiler._phases.append((self.name, elapsed))
                status = "✔ 结束" if exc_type is None else "✖ 异常"
                self.profiler._log(f"⏱{self.profiler._tag} {status}: {self.name} 耗时 {elapsed:.2f}s")
            return False

    def phase(self, name):
        return _PhaseProfiler._Timer(self, name)

    def set_page_count(self, page_count):
        self.page_count = page_count

    def summary(self):
        if not self.enabled or not self._phases:
            return ""
        total = time.perf_counter() - self._t0
        parts = [f"{name} {sec:.2f}s" for name, sec in self._phases]
        head = f"⏱ 计时[{self.page_count}页 总{total:.2f}s]"
        return head + ": " + " | ".join(parts)


class PDFProcessor:
    PRECHECK_OPTION_TITLES = precheck.PRECHECK_OPTION_TITLES
    PRECHECK_DETECTABLE_OPTIONS = precheck.PRECHECK_DETECTABLE_OPTIONS
    PRECHECK_OPTION_ALIASES = precheck.PRECHECK_OPTION_ALIASES
    NON_PROCESSING_OPTIONS = precheck.NON_PROCESSING_OPTIONS
    BASE14_FONT_NAMES = precheck.BASE14_FONT_NAMES

    @staticmethod
    def _get_qpdf_path():
        return qpdf._get_qpdf_path()

    @staticmethod
    def _rewrite_with_qpdf(input_pdf, output_pdf, force_version=None, linearize=False, decrypt_restrictions=False):
        return qpdf._rewrite_with_qpdf(input_pdf, output_pdf, force_version, linearize, decrypt_restrictions)

    @staticmethod
    def _format_qpdf_error(error):
        return qpdf._format_qpdf_error(error)

    @staticmethod
    def _option_title(option_id):
        return precheck._option_title(option_id)

    @staticmethod
    def _precheck_option_matches_selected(option_id, selected_options):
        return precheck._precheck_option_matches_selected(option_id, selected_options)

    @staticmethod
    def _selected_precheck_option_id(option_id, selected_options):
        return precheck._selected_precheck_option_id(option_id, selected_options)

    @staticmethod
    def _filtered_precheck_options(selected_options):
        return precheck._filtered_precheck_options(selected_options)

    @staticmethod
    def _normalize_font_name(font_name):
        return precheck._normalize_font_name(font_name)

    @staticmethod
    def _is_base14_font(font_name):
        return precheck._is_base14_font(font_name)

    @staticmethod
    def _format_font_page_numbers(page_numbers):
        return precheck._format_font_page_numbers(page_numbers)

    @staticmethod
    def _font_object_has_embedded_file(doc, xref, seen=None):
        return precheck._font_object_has_embedded_file(doc, xref, seen)

    @staticmethod
    def _font_tuple_value(font_tuple, index):
        return precheck._font_tuple_value(font_tuple, index)

    @staticmethod
    def _font_tuple_embedded_fallback(font_tuple):
        return precheck._font_tuple_embedded_fallback(font_tuple)

    @staticmethod
    def _collect_font_precheck_findings(doc):
        return precheck._collect_font_precheck_findings(doc)

    @staticmethod
    def _font_precheck_has_embedding_risk(font_precheck):
        return precheck._font_precheck_has_embedding_risk(font_precheck)

    @staticmethod
    def _collect_annotation_findings(doc):
        return precheck._collect_annotation_findings(doc)

    @staticmethod
    def _collect_broken_reference_findings(doc):
        return precheck._collect_broken_reference_findings(doc)

    @staticmethod
    def _collect_annotation_findings_for_path(pdf_path):
        return precheck._collect_annotation_findings_for_path(pdf_path)

    @staticmethod
    def _collect_broken_reference_findings_for_path(pdf_path):
        return precheck._collect_broken_reference_findings_for_path(pdf_path)

    @staticmethod
    def _collect_font_precheck_for_path(pdf_path):
        return precheck._collect_font_precheck_for_path(pdf_path)

    @staticmethod
    def _add_precheck_suggestion(suggestions, option_id, reason):
        return precheck._add_precheck_suggestion(suggestions, option_id, reason)

    @staticmethod
    def _add_precheck_report_finding(suggestions, finding_id, title, reason):
        return precheck._add_precheck_report_finding(suggestions, finding_id, title, reason)

    @staticmethod
    def _resolve_external_file_target(base_dir, file_path):
        return precheck._resolve_external_file_target(base_dir, file_path)

    @staticmethod
    def _read_target_pdf_page_count(target_path):
        return precheck._read_target_pdf_page_count(target_path)

    @staticmethod
    def _add_link_target_integrity_findings(suggestions, base_dir, kind, file_path, target_page=None, source_label="链接"):
        return precheck._add_link_target_integrity_findings(suggestions, base_dir, kind, file_path, target_page, source_label)

    @staticmethod
    def _catalog_key(doc, catalog_xref, key):
        return precheck._catalog_key(doc, catalog_xref, key)

    @staticmethod
    def _catalog_key_is_present(doc, catalog_xref, key):
        return precheck._catalog_key_is_present(doc, catalog_xref, key)

    @staticmethod
    def _read_pdf_header_version(input_path):
        return qpdf._read_pdf_header_version(input_path)

    @staticmethod
    def _is_pdf_linearized(input_path):
        return qpdf._is_pdf_linearized(input_path)

    @staticmethod
    def _qpdf_encryption_info(input_path):
        return qpdf._qpdf_encryption_info(input_path)

    @staticmethod
    def _dereference_xref_value(doc, value):
        return precheck._dereference_xref_value(doc, value)

    @staticmethod
    def _catalog_key_resolved_value(doc, catalog_xref, key):
        return precheck._catalog_key_resolved_value(doc, catalog_xref, key)

    @staticmethod
    def _qpdf_reports_restrictions(input_path):
        return qpdf._qpdf_reports_restrictions(input_path)

    @staticmethod
    def build_precheck_report(input_path, selected_options=None):
        return precheck.build_precheck_report(input_path, selected_options)

    @staticmethod
    def _pdf_has_signature(input_path):
        return precheck._pdf_has_signature(input_path)

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

    @staticmethod
    def resolve_processing_options(input_path, options, processing_mode="smart"):
        return precheck.resolve_processing_options(input_path, options, processing_mode)

    @staticmethod
    def _run_font_embedding_workflow(pdf_path):
        before = PDFProcessor._collect_font_precheck_for_path(pdf_path)
        if not before.get("available"):
            return False, f"字体风险预检失败：{before.get('error') or '无法读取PDF'}", False
        if not PDFProcessor._font_precheck_has_embedding_risk(before):
            return True, "字体风险预检通过，无需调用外部后端", False

        provider = get_font_embedding_provider()
        temp_output = f"{pdf_path}.font_embed.tmp.pdf"
        try:
            result = provider.embed_missing_fonts(pdf_path, temp_output, before)
            if not result.success:
                return False, f"{result.provider_name} 处理失败：{result.message}", False

            after = PDFProcessor._collect_font_precheck_for_path(temp_output)
            if not after.get("available"):
                return False, f"字体修复后验证失败：{after.get('error') or '无法读取PDF'}", False
            if PDFProcessor._font_precheck_has_embedding_risk(after):
                detail = after.get("font_details") or after.get("font_summary") or "仍存在字体嵌入风险"
                provider_detail = result.message.strip() if result.message else ""
                if provider_detail:
                    return False, f"{result.provider_name} 返回成功，但后验证未通过：{detail}；后端返回：{provider_detail}", False
                return False, f"{result.provider_name} 返回成功，但后验证未通过：{detail}", False

            os.replace(temp_output, pdf_path)
            return True, f"{result.provider_name} 字体修复后验证通过", True
        finally:
            if os.path.exists(temp_output):
                try:
                    os.remove(temp_output)
                except Exception:
                    pass

    @staticmethod
    def _transform_rect(rect, scale, dx, dy):
        return page_layout._transform_rect(rect, scale, dx, dy)

    @staticmethod
    def _transform_point(point, scale, dx, dy):
        return page_layout._transform_point(point, scale, dx, dy)

    @staticmethod
    def _get_oriented_target_rect(base_target_rect, src_rect):
        return page_layout._get_oriented_target_rect(base_target_rect, src_rect)

    @staticmethod
    def _paper_rect_exact(size_name):
        return page_layout._paper_rect_exact(size_name)

    @staticmethod
    def _resize_pages_with_padding(doc, target_rect):
        return page_layout._resize_pages_with_padding(doc, target_rect)

    @staticmethod
    def export_bookmarks(pdf_path, csv_path):
        return bookmarks_links.export_bookmarks(pdf_path, csv_path)

    @staticmethod
    def import_bookmarks(pdf_path, csv_path, output_path):
        return bookmarks_links.import_bookmarks(pdf_path, csv_path, output_path)

    @staticmethod
    def export_links(pdf_path, json_path, scope="all"):
        return bookmarks_links.export_links(pdf_path, json_path, scope=scope)

    @staticmethod
    def import_links(pdf_path, json_path, output_path, scope="all", mode="overwrite"):
        return bookmarks_links.import_links(pdf_path, json_path, output_path, scope=scope, mode=mode)

    @staticmethod
    def _is_text_blue(page, rect):
        return hyperlink_styles._is_text_blue(page, rect)

    @staticmethod
    def _overlay_text_color_in_rect(page, rect, color, skip_if_already_blue=False, erase_background=False):
        return hyperlink_styles._overlay_text_color_in_rect(page, rect, color, skip_if_already_blue, erase_background)

    @staticmethod
    def _link_has_visible_border(doc, link_obj):
        return hyperlink_styles._link_has_visible_border(doc, link_obj)

    @staticmethod
    def _force_link_new_window(doc, xref):
        return hyperlink_styles._force_link_new_window(doc, xref)

    @staticmethod
    def _rects_intersect(a, b):
        return hyperlink_styles._rects_intersect(a, b)

    @staticmethod
    def _point_in_any_rect(point, rects):
        return hyperlink_styles._point_in_any_rect(point, rects)

    @staticmethod
    def _make_text_block_blue(block_text):
        return hyperlink_styles._make_text_block_blue(block_text)

    @staticmethod
    def _make_text_block_color(block_text, color_rgb):
        return hyperlink_styles._make_text_block_color(block_text, color_rgb)

    @staticmethod
    def _apply_text_color_via_content_stream(doc, page, target_rects, color_rgb, only_if_blue=False):
        return hyperlink_styles._apply_text_color_via_content_stream(doc, page, target_rects, color_rgb, only_if_blue)

    @staticmethod
    def _collect_page_state(page):
        return hyperlink_styles._collect_page_state(page)

    @staticmethod
    def _apply_blue_text_via_content_stream(doc, page, link_rects=None):
        return hyperlink_styles._apply_blue_text_via_content_stream(doc, page, link_rects)

    @staticmethod
    def _apply_hyperlink_actions(doc, page, options, file_like_link_kinds, page_links=None):
        return hyperlink_styles._apply_hyperlink_actions(doc, page, options, file_like_link_kinds, page_links)

    @staticmethod
    def _apply_hyperlink_styles(doc, page, options, link_objs=None, link_rects=None):
        return hyperlink_styles._apply_hyperlink_styles(doc, page, options, link_objs, link_rects)

    @staticmethod
    def process_document(input_path, output_path, options, processing_mode="smart"):
        prof = _PhaseProfiler()
        try:
            with prof.phase("预检/选项解析"):
                mode_resolution = PDFProcessor.resolve_processing_options(input_path, options, processing_mode)
            options = set(mode_resolution["options"])
            mode_log = mode_resolution.get("log", "")

            with prof.phase("打开文档"):
                doc = fitz.open(input_path)
            prof.set_page_count(doc.page_count)
            applied_changes = []
            change_counts = {}
            link_file_kind = getattr(fitz, "LINK_FILE", None)
            file_like_link_kinds = {fitz.LINK_GOTOR}
            if link_file_kind is not None:
                file_like_link_kinds.add(link_file_kind)

            if doc.needs_pass: return False, "❌ 文件已加密"
            changed = False
            catalog_xref = doc.pdf_catalog()

            if "title_from_filename" in options:
                base_name = Path(input_path).stem
                meta = doc.metadata
                if meta.get("title") != base_name:
                    meta["title"] = base_name;
                    doc.set_metadata(meta);
                    changed = True
                    PDFProcessor._mark_change(applied_changes, "标题同步为文件名")
            elif "fast_web_view" in options:
                base_name = Path(input_path).stem
                meta = doc.metadata
                if not (meta.get("title") or "").strip():
                    meta["title"] = base_name
                    doc.set_metadata(meta)
                    changed = True
                    PDFProcessor._mark_change(applied_changes, "标题补全为文件名")

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
                target_rect = PDFProcessor._paper_rect_exact("a4") if "page_size_a4" in options else PDFProcessor._paper_rect_exact(
                    "letter")
                with prof.phase("页面尺寸标准化"):
                    resized_pages = PDFProcessor._resize_pages_with_padding(doc, target_rect)
                if resized_pages > 0:
                    changed = True
                    PDFProcessor._mark_change(applied_changes, "页面尺寸标准化")
                    PDFProcessor._increase_change_count(change_counts, "页面尺寸标准化", resized_pages)

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
                with prof.phase("超链接遍历"):
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
            _cleanup_active = any(opt in options for opt in cleanup_options)
            _cleanup_ctx = prof.phase("清理遍历") if _cleanup_active else None
            if _cleanup_ctx is not None:
                _cleanup_ctx.__enter__()
            if _cleanup_active:
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

                if _cleanup_ctx is not None:
                    _cleanup_ctx.__exit__(None, None, None)
                    _cleanup_ctx = None

                if "cleanup_remove_annotations" in options:
                    with prof.phase("删除注释遍历"):
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
            force_pdf_version = "1.7" if "convert_pdf_version" in options else None
            remove_pdf_restrictions = "remove_pdf_restrictions" in options
            needs_qpdf_rewrite = bool(is_linear or force_pdf_version or remove_pdf_restrictions)

            if changed:
                if needs_qpdf_rewrite:
                    temp_pdf = str(output_path) + ".tmp.pdf"
                    with prof.phase("保存(deflate+garbage2+objstms)"):
                        doc.save(temp_pdf, garbage=2, deflate=True, use_objstms=1)
                        doc.close()
                    try:
                        with prof.phase("qpdf重写"):
                            PDFProcessor._rewrite_with_qpdf(
                                temp_pdf,
                                output_path,
                                force_version=force_pdf_version,
                                linearize=is_linear,
                                decrypt_restrictions=remove_pdf_restrictions,
                            )
                        if remove_pdf_restrictions:
                            PDFProcessor._mark_change(applied_changes, "已解除PDF权限限制")
                        if force_pdf_version:
                            PDFProcessor._mark_change(applied_changes, "已转换PDF版本")
                        if is_linear:
                            PDFProcessor._mark_change(applied_changes, "已启用快速网页浏览")
                    finally:
                        if os.path.exists(temp_pdf):
                            os.remove(temp_pdf)
                else:
                    with prof.phase("保存(deflate+garbage2+objstms)"):
                        doc.save(output_path, garbage=2, deflate=True, use_objstms=1)
                        doc.close()
            else:
                doc.close()
                if needs_qpdf_rewrite:
                    PDFProcessor._rewrite_with_qpdf(
                        input_path,
                        output_path,
                        force_version=force_pdf_version,
                        linearize=is_linear,
                        decrypt_restrictions=remove_pdf_restrictions,
                    )
                    if remove_pdf_restrictions:
                        PDFProcessor._mark_change(applied_changes, "已解除PDF权限限制")
                    if force_pdf_version:
                        PDFProcessor._mark_change(applied_changes, "已转换PDF版本")
                    if is_linear:
                        PDFProcessor._mark_change(applied_changes, "已启用快速网页浏览")
                else:
                    shutil.copy2(input_path, output_path)

            if applied_changes:
                result_msg = f"✅ 处理成功；修改项：{PDFProcessor._format_change_summary(change_counts, applied_changes)}"
            else:
                result_msg = "✅ 处理成功；无实际修改"
            if mode_log:
                result_msg = f"{result_msg}；{mode_log}"
            prof_summary = prof.summary()
            if prof_summary:
                result_msg = f"{result_msg}\n    {prof_summary}"
            return True, result_msg

        except FileNotFoundError as e:
            return False, f"⚠️ 缺少引擎组件: {str(e)}"
        except Exception as e:
            if "remove_pdf_restrictions" in options:
                return False, f"❌ 处理失败: {PDFProcessor._format_qpdf_error(e)}"
            return False, f"❌ 处理失败: {str(e)}"
