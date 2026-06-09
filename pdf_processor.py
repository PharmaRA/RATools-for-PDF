import fitz
import os
import sys
import subprocess
import shutil
import csv
import json
import re
from urllib.parse import unquote, urlparse
from pathlib import Path

from app_paths import get_resource_path
from font_embedding_providers import get_font_embedding_provider


class PDFProcessor:

    @staticmethod
    def _get_qpdf_path():
        if sys.platform == "win32":
            candidates = [
                get_resource_path("plugins", "qpdf", "qpdf.exe"),
                os.environ.get("QPDF_PATH", ""),
                r"D:\Program Files\qpdf 11.9.1\bin\qpdf.exe",
                r"C:\Program Files\qpdf\bin\qpdf.exe",
            ]
            for candidate in candidates:
                if candidate and os.path.exists(candidate):
                    return candidate
            return "qpdf.exe"
        return "qpdf"

    @staticmethod
    def _rewrite_with_qpdf(input_pdf, output_pdf, force_version=None, linearize=False, decrypt_restrictions=False):
        qpdf_exe = PDFProcessor._get_qpdf_path()
        if sys.platform == "win32" and qpdf_exe != "qpdf.exe" and not os.path.exists(qpdf_exe):
            raise FileNotFoundError(f"未找到 qpdf 工具！\n请确保已安装或设置 QPDF_PATH: {qpdf_exe}")

        cmd = [qpdf_exe]
        if decrypt_restrictions:
            cmd.append("--decrypt")
        if linearize:
            cmd.append("--linearize")
        if force_version:
            cmd.append(f"--force-version={force_version}")
        cmd.extend([input_pdf, output_pdf])

        startupinfo = None
        if sys.platform == "win32":
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW

        result = subprocess.run(cmd, startupinfo=startupinfo, capture_output=True, text=True)
        if result.returncode != 0:
            raise RuntimeError(f"qpdf 执行失败: {result.stderr}")

    @staticmethod
    def _format_qpdf_error(error):
        text = str(error).strip()
        lowered = text.lower()
        password_markers = ["password", "invalid password", "incorrect password", "requires a password"]
        if any(marker in lowered for marker in password_markers):
            return "该PDF需要密码，当前模式不支持输入密码解锁"
        if text.startswith("qpdf 执行失败:"):
            detail = text.split(":", 1)[1].strip()
            return f"未能移除PDF权限限制：{detail}" if detail else "未能移除PDF权限限制"
        return text

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

    @staticmethod
    def _option_title(option_id):
        return PDFProcessor.PRECHECK_OPTION_TITLES.get(option_id, option_id)

    @staticmethod
    def _precheck_option_matches_selected(option_id, selected_options):
        if selected_options is None:
            return True
        selected_options = set(selected_options or [])
        if option_id in selected_options:
            return True
        return bool(PDFProcessor.PRECHECK_OPTION_ALIASES.get(option_id, set()) & selected_options)

    @staticmethod
    def _selected_precheck_option_id(option_id, selected_options):
        if selected_options is None or option_id in set(selected_options or []):
            return option_id
        for alias in PDFProcessor.PRECHECK_OPTION_ALIASES.get(option_id, set()):
            if alias in selected_options:
                return alias
        return option_id

    @staticmethod
    def _filtered_precheck_options(selected_options):
        if selected_options is None:
            return None
        selected_options = set(selected_options or [])
        filtered = set()
        for option_id in selected_options:
            if option_id in PDFProcessor.PRECHECK_DETECTABLE_OPTIONS:
                filtered.add(option_id)
        for option_id, aliases in PDFProcessor.PRECHECK_OPTION_ALIASES.items():
            if option_id in selected_options or aliases & selected_options:
                filtered.add(option_id)
                filtered.update(aliases & selected_options)
        return filtered

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

    @staticmethod
    def _normalize_font_name(font_name):
        name = str(font_name or "").strip()
        if re.match(r"^[A-Z]{6}\+", name):
            name = name.split("+", 1)[1]
        return name

    @staticmethod
    def _is_base14_font(font_name):
        return PDFProcessor._normalize_font_name(font_name) in PDFProcessor.BASE14_FONT_NAMES

    @staticmethod
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

    @staticmethod
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
            embedded, child_known = PDFProcessor._font_object_has_embedded_file(doc, ref, seen)
            known = known and child_known
            if embedded:
                return True, True
        return False, known

    @staticmethod
    def _font_tuple_value(font_tuple, index):
        if len(font_tuple) <= index:
            return ""
        return font_tuple[index]

    @staticmethod
    def _font_tuple_embedded_fallback(font_tuple):
        ext = str(PDFProcessor._font_tuple_value(font_tuple, 1) or "").strip().lower()
        if ext and ext not in {"n/a", "none", "null"}:
            return True, True
        if ext in {"n/a", "none", "null"}:
            return False, True
        return False, False

    @staticmethod
    def _collect_font_precheck_findings(doc):
        fonts = {}
        for page_index, page in enumerate(doc, start=1):
            try:
                page_fonts = page.get_fonts(full=True)
            except Exception:
                continue

            for font_tuple in page_fonts:
                original_name = str(PDFProcessor._font_tuple_value(font_tuple, 3) or "").strip()
                if not original_name:
                    original_name = str(PDFProcessor._font_tuple_value(font_tuple, 4) or "").strip()
                normalized_name = PDFProcessor._normalize_font_name(original_name)
                if not normalized_name:
                    continue

                try:
                    xref = int(PDFProcessor._font_tuple_value(font_tuple, 0) or 0)
                except Exception:
                    xref = 0
                embedded, known = PDFProcessor._font_object_has_embedded_file(doc, xref)
                if not embedded and known:
                    embedded, known = PDFProcessor._font_tuple_embedded_fallback(font_tuple)

                entry = fonts.setdefault(normalized_name, {
                    "font_name": original_name,
                    "original_names": set(),
                    "normalized_name": normalized_name,
                    "pages": set(),
                    "has_unembedded": False,
                    "has_embedded": False,
                    "embedding_unknown": False,
                    "base14": PDFProcessor._is_base14_font(normalized_name),
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
                f"{item['normalized_name']}({PDFProcessor._format_font_page_numbers(item['pages'])}，{'，'.join(labels)})"
            )

        return {
            "font_summary": "，".join(summary_parts),
            "font_details": "; ".join(detail_parts),
            "font_findings": findings,
        }

    @staticmethod
    def _font_precheck_has_embedding_risk(font_precheck):
        for item in font_precheck.get("font_findings", []) or []:
            if (
                item.get("has_unembedded")
                or item.get("substitution_risk")
                or item.get("has_unknown_embedding_status")
            ):
                return True
        return False

    @staticmethod
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
            font_precheck = PDFProcessor._collect_font_precheck_findings(doc)
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

    @staticmethod
    def _add_precheck_suggestion(suggestions, option_id, reason):
        if option_id in suggestions:
            return
        suggestions[option_id] = {
            "matched": True,
            "title": PDFProcessor.PRECHECK_OPTION_TITLES.get(option_id, option_id),
            "reason": reason,
        }

    @staticmethod
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

    @staticmethod
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

    @staticmethod
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

    @staticmethod
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
            PDFProcessor._add_precheck_report_finding(
                suggestions,
                finding_id,
                title,
                f"{source_label} 未指定目标文件",
            )
            return

        decoded_path, target_path, is_relative = PDFProcessor._resolve_external_file_target(base_dir, file_path)
        if not decoded_path:
            return

        if not os.path.isfile(target_path):
            if kind == fitz.LINK_GOTOR:
                PDFProcessor._add_precheck_report_finding(
                    suggestions,
                    "link_target_integrity_gotor_missing_file",
                    "链接目标完整性检查：GoToR目标文件不存在",
                    f"{source_label} 指向的目标文件不存在: {decoded_path}",
                )
            else:
                PDFProcessor._add_precheck_report_finding(
                    suggestions,
                    "link_target_integrity_missing_file",
                    "链接目标完整性检查：外部文件链接目标不存在",
                    f"{source_label} 指向的目标文件不存在: {decoded_path}",
                )
            if is_relative:
                PDFProcessor._add_precheck_report_finding(
                    suggestions,
                    "link_target_integrity_broken_relative_path",
                    "链接目标完整性检查：eCTD相对路径链接断链",
                    f"{source_label} 指向的相对路径不存在: {decoded_path}",
                )
            return

        if kind != fitz.LINK_GOTOR:
            return

        page_count = PDFProcessor._read_target_pdf_page_count(target_path)
        if page_count is None:
            return
        try:
            target_page_index = int(target_page if target_page is not None else 0)
        except Exception:
            target_page_index = -1
        if target_page_index < 0 or target_page_index >= page_count:
            PDFProcessor._add_precheck_report_finding(
                suggestions,
                "link_target_integrity_gotor_page_out_of_range",
                "链接目标完整性检查：GoToR目标页码越界",
                f"{source_label} 指向 {decoded_path} 的第 {target_page_index + 1} 页，超过目标文件页数 {page_count}",
            )

    @staticmethod
    def _catalog_key(doc, catalog_xref, key):
        try:
            return doc.xref_get_key(catalog_xref, key)
        except Exception:
            return ("null", "null")

    @staticmethod
    def _catalog_key_is_present(doc, catalog_xref, key):
        kind, value = PDFProcessor._catalog_key(doc, catalog_xref, key)
        return kind != "null" and value != "null"

    @staticmethod
    def _read_pdf_header_version(input_path):
        try:
            with open(input_path, "rb") as f:
                header = f.read(32)
        except Exception:
            return ""
        match = re.search(rb"%PDF-(\d+\.\d+)", header)
        return match.group(1).decode("ascii") if match else ""

    @staticmethod
    def _is_pdf_linearized(input_path):
        try:
            with open(input_path, "rb") as f:
                header = f.read(4096)
        except Exception:
            return False
        return b"/Linearized" in header

    @staticmethod
    def _qpdf_encryption_info(input_path):
        qpdf_exe = PDFProcessor._get_qpdf_path()
        if sys.platform == "win32" and qpdf_exe != "qpdf.exe" and not os.path.exists(qpdf_exe):
            return ""

        startupinfo = None
        if sys.platform == "win32":
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW

        try:
            result = subprocess.run(
                [qpdf_exe, "--show-encryption", input_path],
                startupinfo=startupinfo,
                capture_output=True,
                text=True,
                timeout=15,
            )
        except Exception:
            return ""
        return f"{result.stdout}\n{result.stderr}".strip()

    @staticmethod
    def _dereference_xref_value(doc, value):
        match = re.match(r"^(\d+)\s+0\s+R$", str(value or "").strip())
        if not match:
            return str(value or "")
        try:
            return doc.xref_object(int(match.group(1)))
        except Exception:
            return str(value or "")

    @staticmethod
    def _catalog_key_resolved_value(doc, catalog_xref, key):
        _kind, value = PDFProcessor._catalog_key(doc, catalog_xref, key)
        return PDFProcessor._dereference_xref_value(doc, value)

    @staticmethod
    def _qpdf_reports_restrictions(input_path):
        info = PDFProcessor._qpdf_encryption_info(input_path).lower()
        if not info or "file is not encrypted" in info:
            return False
        return ": not allowed" in info

    @staticmethod
    def build_precheck_report(input_path, selected_options=None):
        selected_options = PDFProcessor._filtered_precheck_options(selected_options)

        def wants(*option_ids):
            if selected_options is None:
                return True
            return any(PDFProcessor._precheck_option_matches_selected(option_id, selected_options) for option_id in option_ids)

        def add_suggestion(option_id, reason):
            if not wants(option_id):
                return
            suggestion_id = PDFProcessor._selected_precheck_option_id(option_id, selected_options)
            PDFProcessor._add_precheck_suggestion(suggestions, suggestion_id, reason)

        report = {
            "available": False,
            "file_path": input_path,
            "file_name": os.path.basename(input_path),
            "suggestions": {},
            "font_summary": "",
            "font_details": "",
            "font_findings": [],
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
                page_mode_kind, page_mode_value = PDFProcessor._catalog_key(doc, catalog_xref, "PageMode")
                if has_bookmarks and not (page_mode_kind == "name" and page_mode_value == "/UseOutlines"):
                    add_suggestion("initial_view_bookmarks_and_page", "文档包含书签，但打开时未设置为显示书签面板")
                elif not has_bookmarks and page_mode_kind == "name" and page_mode_value != "/UseNone":
                    add_suggestion("initial_view_bookmarks_and_page", "文档不含书签，但初始导览标签不是页面视图")

            if wants("page_layout_default") and PDFProcessor._catalog_key_is_present(doc, catalog_xref, "PageLayout"):
                add_suggestion("page_layout_default", "文档设置了显式页面布局")

            if wants("open_page_first", "zoom_default"):
                open_action_kind, open_action_value = PDFProcessor._catalog_key(doc, catalog_xref, "OpenAction")
            else:
                open_action_kind, open_action_value = "null", "null"
            if open_action_kind != "null" and doc.page_count > 0:
                open_action_value = PDFProcessor._dereference_xref_value(doc, open_action_value)
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
                        PDFProcessor._add_link_target_integrity_findings(
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
                            PDFProcessor._add_link_target_integrity_findings(
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
                PDFProcessor._catalog_key_is_present(doc, catalog_xref, "StructTreeRoot")
                or PDFProcessor._catalog_key_is_present(doc, catalog_xref, "MarkInfo")
            ):
                add_suggestion("cleanup_remove_tags", "文档包含结构化标签信息")

            if wants("cleanup_remove_metadata"):
                metadata_values = [
                    value
                    for key, value in meta.items()
                    if key not in ["format", "encryption"] and str(value or "").strip()
                ]
                if metadata_values or PDFProcessor._catalog_key_is_present(doc, catalog_xref, "PieceInfo"):
                    add_suggestion("cleanup_remove_metadata", "文档包含可清理的元数据")

            if wants("cleanup_remove_dynamic_content"):
                catalog_object = ""
                try:
                    catalog_object = doc.xref_object(catalog_xref)
                except Exception:
                    catalog_object = ""
                names_kind, _names_value = PDFProcessor._catalog_key(doc, catalog_xref, "Names")
                names_value = PDFProcessor._catalog_key_resolved_value(doc, catalog_xref, "Names")
                dynamic_probe = f"{catalog_object}\n{names_value}"
                if names_kind != "null" and any(marker in dynamic_probe for marker in ["/JavaScript", "/JS", "/RichMedia", "/3D"]):
                    add_suggestion("cleanup_remove_dynamic_content", "文档包含 JavaScript、3D 或富媒体入口")

            if wants("convert_pdf_version"):
                pdf_version = PDFProcessor._read_pdf_header_version(input_path)
                if pdf_version and pdf_version != "1.7":
                    add_suggestion("convert_pdf_version", f"当前PDF版本为 {pdf_version}，不是 1.7")

            if wants("fast_web_view") and not PDFProcessor._is_pdf_linearized(input_path):
                add_suggestion("fast_web_view", "文档未启用线性化快速网页浏览")

            if wants("remove_pdf_restrictions") and PDFProcessor._qpdf_reports_restrictions(input_path):
                add_suggestion("remove_pdf_restrictions", "文档存在打印、复制或编辑权限限制")

            if selected_options is None:
                try:
                    font_precheck = PDFProcessor._collect_font_precheck_findings(doc)
                    report["font_summary"] = font_precheck.get("font_summary", "")
                    report["font_details"] = font_precheck.get("font_details", "")
                    report["font_findings"] = font_precheck.get("font_findings", [])
                    if report["font_summary"]:
                        reason = report["font_summary"]
                        if report["font_details"]:
                            reason = f"{reason}；明细：{report['font_details']}"
                        PDFProcessor._add_precheck_report_finding(
                            suggestions,
                            "font_precheck_review",
                            "字体预检：需要复核",
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
                "log": "处理模式: 强制执行全部勾选规则",
            }

        actionable_selected = selected_options - PDFProcessor.NON_PROCESSING_OPTIONS
        unsupported = actionable_selected - PDFProcessor.PRECHECK_DETECTABLE_OPTIONS
        report = PDFProcessor.build_precheck_report(input_path, selected_options=actionable_selected)
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
        effective = set(selected_options & PDFProcessor.NON_PROCESSING_OPTIONS)
        effective.update(actionable_selected & suggested)
        effective.update(unsupported)

        skipped = sorted(
            option_id
            for option_id in actionable_selected
            if option_id in PDFProcessor.PRECHECK_DETECTABLE_OPTIONS and option_id not in effective
        )
        forced_unsupported = sorted(unsupported)
        log_parts = ["处理模式: 智能处理（仅处理预检发现的问题）"]
        if suggested:
            log_parts.append(
                "预检命中: " + "、".join(PDFProcessor._option_title(option_id) for option_id in sorted(suggested))
            )
        else:
            log_parts.append("预检命中: 无")
        if forced_unsupported:
            log_parts.append(
                "无法可靠预检但已执行: "
                + "、".join(PDFProcessor._option_title(option_id) for option_id in forced_unsupported)
            )
        if skipped:
            log_parts.append(
                "已跳过未命中规则: "
                + "、".join(PDFProcessor._option_title(option_id) for option_id in skipped)
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
        return fitz.Rect(
            rect.x0 * scale + dx,
            rect.y0 * scale + dy,
            rect.x1 * scale + dx,
            rect.y1 * scale + dy,
        )

    @staticmethod
    def _transform_point(point, scale, dx, dy):
        if point is None:
            return None
        return fitz.Point(point.x * scale + dx, point.y * scale + dy)

    @staticmethod
    def _get_oriented_target_rect(base_target_rect, src_rect):
        if src_rect.width > src_rect.height and base_target_rect.width < base_target_rect.height:
            return fitz.Rect(0, 0, base_target_rect.height, base_target_rect.width)
        if src_rect.width < src_rect.height and base_target_rect.width > base_target_rect.height:
            return fitz.Rect(0, 0, base_target_rect.height, base_target_rect.width)
        return fitz.Rect(0, 0, base_target_rect.width, base_target_rect.height)

    @staticmethod
    def _paper_rect_exact(size_name):
        size = size_name.lower()
        if size == "a4":
            return fitz.Rect(0, 0, 210 / 25.4 * 72, 297 / 25.4 * 72)
        if size == "letter":
            return fitz.Rect(0, 0, 8.5 * 72, 11 * 72)
        return fitz.paper_rect(size_name)

    @staticmethod
    def _resize_pages_with_padding(doc, target_rect):
        page_transforms = []
        for page in doc:
            src_rect = page.rect
            page_target_rect = PDFProcessor._get_oriented_target_rect(target_rect, src_rect)

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
                    link["from"] = PDFProcessor._transform_rect(fitz.Rect(link["from"]), scale, dx, dy)
                    if link.get("kind") == fitz.LINK_GOTO and link.get("to") is not None:
                        dest_page = int(link.get("page", page_index))
                        dest_transform = page_transforms[dest_page] if 0 <= dest_page < len(page_transforms) else None
                        if dest_transform is not None:
                            d_scale = dest_transform["scale"]
                            d_dx = dest_transform["dx"]
                            d_dy = dest_transform["dy"]
                            to_point = link.get("to")
                            if hasattr(to_point, "x") and hasattr(to_point, "y"):
                                link["to"] = PDFProcessor._transform_point(to_point, d_scale, d_dx, d_dy)
                    page.update_link(link)
                except Exception:
                    continue

            for annot in annots:
                try:
                    annot.set_rect(PDFProcessor._transform_rect(annot.rect, scale, dx, dy))
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
                dest["to"] = PDFProcessor._transform_point(to_point, scale, dx, dy)
                toc_changed = True

            if toc_changed:
                doc.set_toc(toc)

        return sum(1 for item in page_transforms if item is not None)

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
    def process_document(input_path, output_path, options, processing_mode="smart"):
        try:
            mode_resolution = PDFProcessor.resolve_processing_options(input_path, options, processing_mode)
            options = set(mode_resolution["options"])
            mode_log = mode_resolution.get("log", "")

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
            force_pdf_version = "1.7" if "convert_pdf_version" in options else None
            remove_pdf_restrictions = "remove_pdf_restrictions" in options
            needs_qpdf_rewrite = bool(is_linear or force_pdf_version or remove_pdf_restrictions)

            if changed:
                if needs_qpdf_rewrite:
                    temp_pdf = str(output_path) + ".tmp.pdf"
                    doc.save(temp_pdf, garbage=3, deflate=True)
                    doc.close()
                    try:
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
                    doc.save(output_path, garbage=3, deflate=True)
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
            return True, result_msg

        except FileNotFoundError as e:
            return False, f"⚠️ 缺少引擎组件: {str(e)}"
        except Exception as e:
            if "remove_pdf_restrictions" in options:
                return False, f"❌ 处理失败: {PDFProcessor._format_qpdf_error(e)}"
            return False, f"❌ 处理失败: {str(e)}"
