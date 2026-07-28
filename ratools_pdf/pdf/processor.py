import os
import shutil
import sys
import time
from pathlib import Path

import fitz

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


def _mark_change(change_list, label):
    if label not in change_list:
        change_list.append(label)


def _increase_change_count(change_counts, label, amount=1):
    if amount <= 0:
        return
    change_counts[label] = change_counts.get(label, 0) + amount


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


class _PipelineContext:
    """process_document 各 step 之间共享的可变状态。"""

    def __init__(self, doc, options, input_path, prof):
        self.doc = doc
        self.options = options
        self.input_path = input_path
        self.prof = prof
        self.changed = False
        self.applied_changes = []
        self.change_counts = {}
        self.catalog_xref = doc.pdf_catalog()

        link_file_kind = getattr(fitz, "LINK_FILE", None)
        self.file_like_link_kinds = {fitz.LINK_GOTOR}
        if link_file_kind is not None:
            self.file_like_link_kinds.add(link_file_kind)

    def mark(self, label, count=0):
        """记录一次变更；count > 0 时同时累计次数统计。"""
        self.changed = True
        _mark_change(self.applied_changes, label)
        if count:
            _increase_change_count(self.change_counts, label, count)


# ============================================================================
# 管线步骤。每个 step 接收 _PipelineContext，按需读取 ctx.options 并写回变更。
# 步骤顺序即 process_document 的执行顺序，与历史行为保持一致。
# ============================================================================


def _step_document_title(ctx):
    """title_from_filename：标题同步为文件名；fast_web_view：空标题补全。"""
    doc, options = ctx.doc, ctx.options
    if "title_from_filename" in options:
        base_name = Path(ctx.input_path).stem
        meta = doc.metadata
        if meta.get("title") != base_name:
            meta["title"] = base_name
            doc.set_metadata(meta)
            ctx.mark("标题同步为文件名")
    elif "fast_web_view" in options:
        base_name = Path(ctx.input_path).stem
        meta = doc.metadata
        if not (meta.get("title") or "").strip():
            meta["title"] = base_name
            doc.set_metadata(meta)
            ctx.mark("标题补全为文件名")


def _step_open_action(ctx):
    """open_page_first / zoom_default：打开动作指向第一页 + 阅读器默认缩放。"""
    doc, options = ctx.doc, ctx.options
    if "open_page_first" not in options and "zoom_default" not in options:
        return
    if doc.page_count <= 0:
        return
    page0_xref = doc[0].xref
    # /XYZ null null null 表示使用阅读器默认缩放，不强制 Fit/固定倍率
    action_str = f"[{page0_xref} 0 R /XYZ null null null]"
    doc.xref_set_key(ctx.catalog_xref, "OpenAction", action_str)
    if "open_page_first" in options:
        ctx.mark("打开页设为第一页")
    if "zoom_default" in options:
        ctx.mark("打开缩放设为默认")


def _step_page_layout_reset(ctx):
    """page_layout_default：移除显式 PageLayout，恢复阅读器默认行为。"""
    if "page_layout_default" not in ctx.options:
        return
    ctx.doc.xref_set_key(ctx.catalog_xref, "PageLayout", "null")
    ctx.mark("页面布局恢复默认")


def _step_page_mode(ctx):
    """initial_view_bookmarks_and_page：有书签显示书签面板，否则纯页面视图。"""
    if "initial_view_bookmarks_and_page" not in ctx.options:
        return
    doc = ctx.doc
    has_bookmarks = len(doc.get_toc(simple=False)) > 0
    page_mode = "/UseOutlines" if has_bookmarks else "/UseNone"
    doc.xref_set_key(ctx.catalog_xref, "PageMode", page_mode)
    ctx.mark("初始视图设为书签/页面")


def _step_collapse_bookmarks(ctx):
    """collapse_all_bookmarks：全部书签设为折叠。"""
    if "collapse_all_bookmarks" not in ctx.options:
        return
    doc = ctx.doc
    toc = doc.get_toc(simple=False)
    if not toc:
        return
    for item in toc:
        if isinstance(item[-1], dict):
            item[-1]["collapse"] = True
    doc.set_toc(toc)
    ctx.mark("折叠全部书签")


def _step_page_size(ctx):
    """page_size_a4 / page_size_letter：页面等比缩放并居中留白到目标纸型。"""
    options = ctx.options
    if "page_size_a4" not in options and "page_size_letter" not in options:
        return
    size_name = "a4" if "page_size_a4" in options else "letter"
    target_rect = page_layout._paper_rect_exact(size_name)
    with ctx.prof.phase("页面尺寸标准化"):
        resized_pages = page_layout._resize_pages_with_padding(ctx.doc, target_rect)
    if resized_pages > 0:
        ctx.mark("页面尺寸标准化", count=resized_pages)


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
    """把 get_toc 返回的目的地字典规整为 set_toc 可靠接受的形状。

    命名目标（LINK_NAMED）没有 set_toc 能写回的形状：能解析出页码的按内部跳转
    输出，悬空的降级为空目的地，交由失效书签规则处理。
    """
    if not isinstance(dest, dict):
        dest = {}

    if kind == fitz.LINK_NAMED:
        named_page = bookmarks_links.bookmark_named_dest_page(dest)
        if named_page is None:
            return {"kind": fitz.LINK_NONE}
        kind = fitz.LINK_GOTO
        dest = dict(dest, kind=fitz.LINK_GOTO, page=named_page)

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


def _normalize_toc_levels(toc):
    """书签层级修复：首项归 1，后项最多比前项深一级（set_toc 的硬性要求）。"""
    for i in range(len(toc)):
        if i == 0:
            toc[i][0] = 1
        else:
            prev_lvl = toc[i - 1][0]
            if toc[i][0] > prev_lvl + 1:
                toc[i][0] = prev_lvl + 1


_BOOKMARK_RULE_OPTIONS = (
    "bookmark_inherit_zoom",
    "bookmark_open_new_window",
    "bookmark_remove_external_links",
    "bookmark_remove_invalid",
    "bookmark_remove_unknown_actions",
)


def _step_bookmark_rules(ctx):
    """书签规则组：删除外链/失效/未知动作、承前缩放、新窗口打开。"""
    options = ctx.options
    if not any(opt in options for opt in _BOOKMARK_RULE_OPTIONS):
        return
    doc = ctx.doc
    toc = doc.get_toc(simple=False)
    if not toc:
        return

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

        raw_dest = dest if isinstance(dest, dict) else {}
        kind = raw_dest.get("kind", fitz.LINK_NONE)
        # 失效判定必须用原始目的地：规整会把命名目标折叠成 GOTO，丢掉判定依据
        dest_invalid = bookmarks_links.is_bookmark_dest_invalid(raw_dest, bm_page, doc.page_count)
        dest = _normalize_bookmark_dest(raw_dest, kind)
        # 命名目标已在规整时解析为内部跳转，后续按 GOTO 处理
        kind = dest.get("kind", fitz.LINK_NONE)
        delete_it = False

        if "bookmark_remove_external_links" in options and kind == fitz.LINK_URI:
            delete_it = True
        if "bookmark_remove_invalid" in options and dest_invalid:
            delete_it = True
        if "bookmark_remove_unknown_actions" in options:
            if kind not in [fitz.LINK_GOTO, fitz.LINK_GOTOR, fitz.LINK_LAUNCH]:
                delete_it = True

        if delete_it:
            toc_modified = True
            continue

        if "bookmark_inherit_zoom" in options and kind == fitz.LINK_GOTO:
            if dest.get("zoom") != 0.0:
                dest["zoom"] = 0.0
                toc_modified = True
        if "bookmark_open_new_window" in options and kind in [fitz.LINK_GOTOR, fitz.LINK_LAUNCH]:
            if not dest.get("newWindow"):
                dest["newWindow"] = True
                toc_modified = True

        if kind == fitz.LINK_GOTO:
            if bm_page < 1:
                bm_page = 1
                toc_modified = True
            elif bm_page > doc.page_count:
                bm_page = doc.page_count
                toc_modified = True

        new_toc.append([lvl, title, bm_page, dest])

    if not toc_modified:
        return

    if new_toc:
        _normalize_toc_levels(new_toc)
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
    ctx.mark("书签规则已更新")


_HYPERLINK_RULE_OPTIONS = (
    "link_abs_to_rel_path",
    "link_inherit_zoom",
    "link_open_new_window",
    "link_text_blue",
    "link_black_border",
    "link_bordered_to_blue_border",
    "link_unbordered_blue_to_blue_border",
    "link_remove_border",
)


def _step_hyperlink_rules(ctx):
    """超链接规则组：动作调整（路径/缩放/新窗口）与外观调整（颜色/边框）。"""
    options = ctx.options
    if not any(opt in options for opt in _HYPERLINK_RULE_OPTIONS):
        return
    doc = ctx.doc
    with ctx.prof.phase("超链接遍历"):
        for page in doc:
            page_state = hyperlink_styles._collect_page_state(page)
            if hyperlink_styles._apply_hyperlink_actions(
                doc,
                page,
                options,
                ctx.file_like_link_kinds,
                page_links=page_state["links"],
            ):
                ctx.mark("超链接动作已更新")
            if hyperlink_styles._apply_hyperlink_styles(
                doc,
                page,
                options,
                link_objs=page_state["link_objs"],
                link_rects=page_state["link_rects"],
            ):
                ctx.mark("超链接外观已更新")


_CLEANUP_PAGE_OPTIONS = (
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
)

_EXTERNAL_URI_OPTS = {"cleanup_remove_external_uri", "cleanup_remove_external_uri_and_text_black"}

_LINK_ANNOT_TYPE = 8  # PDF 注释类型编号：Link


def _annot_uri(annot):
    """兼容不同 PyMuPDF 版本的 Link 注释 uri 读取。"""
    uri = getattr(annot, "uri", "") or ""
    if not uri and hasattr(annot, "info"):
        uri = annot.info.get("uri", "") or ""
    return uri


def _decolor_rects_to_black(ctx, page, rects):
    """把指定区域内的蓝色文字改回黑色（内容流级）。"""
    if not rects:
        return
    if hyperlink_styles._apply_text_color_via_content_stream(
        ctx.doc, page, rects, (0.0, 0.0, 0.0), only_if_blue=True,
    ):
        ctx.mark("已将链接文本恢复为黑色")


def _cleanup_external_uri_fast_path(ctx):
    """性能快路径：仅删除外部 URI（可选去色）时，避免扫描注释和其他重逻辑。"""
    options = ctx.options
    decolor = "cleanup_remove_external_uri_and_text_black" in options
    for page in ctx.doc:
        decolor_rects = []
        removed_count = 0

        for link in page.get_links():
            if link.get("kind", fitz.LINK_NONE) != fitz.LINK_URI:
                continue
            if decolor:
                try:
                    decolor_rects.append(fitz.Rect(link.get("from")))
                except Exception:
                    pass
            try:
                page.delete_link(link)
                removed_count += 1
                ctx.mark("已删除外部URI链接", count=1)
            except Exception:
                pass

        if decolor:
            _decolor_rects_to_black(ctx, page, decolor_rects)

        # 兼容兜底：若仍有 URI 链接残留，再做一次注释级删除
        if removed_count > 0 and any(
            l.get("kind", fitz.LINK_NONE) == fitz.LINK_URI for l in page.get_links()
        ):
            for annot in page.annots() or []:
                try:
                    if annot.type[0] != _LINK_ANNOT_TYPE:
                        continue
                    if _annot_uri(annot):
                        page.delete_annot(annot)
                        ctx.mark("已删除外部URI链接", count=1)
                except Exception:
                    pass


def _cleanup_all_links_bookmarks(ctx):
    """一键删除全部链接与书签（不动普通批注）。"""
    doc = ctx.doc
    doc.set_toc([])
    for page in doc:
        page_state = hyperlink_styles._collect_page_state(page)
        # 直接删除 Link 注释，避免部分 PDF 中 delete_link 命中不到
        for annot in page_state["annots"]:
            try:
                if annot.type[0] == _LINK_ANNOT_TYPE:
                    page.delete_annot(annot)
            except Exception:
                pass
        # 兜底：再按 get_links 删除一遍
        for link in page_state["links"]:
            try:
                page.delete_link(link)
            except Exception:
                pass
    ctx.mark("已删除全部链接和书签")


def _cleanup_links_general_path(ctx):
    """通用清理路径：按选项删除 URI/失效/未知动作链接，并按需去色。"""
    options = ctx.options
    for page in ctx.doc:
        page_state = hyperlink_styles._collect_page_state(page)
        decolor_rects = []

        # 外部 URI 链接：优先用 delete_annot 方式确保真的移除可点击行为
        if (
            "cleanup_remove_external_uri" in options
            or "cleanup_remove_external_uri_and_text_black" in options
        ):
            for annot in page_state["annots"]:
                try:
                    if annot.type[0] != _LINK_ANNOT_TYPE:
                        continue
                    if _annot_uri(annot):
                        if "cleanup_remove_external_uri_and_text_black" in options:
                            decolor_rects.append(annot.rect)
                        page.delete_annot(annot)
                        ctx.mark("已删除外部URI链接", count=1)
                except Exception:
                    pass

        for link in page_state["links"]:
            kind = link.get("kind", fitz.LINK_NONE)
            delete_it = False
            if kind == fitz.LINK_URI and (
                    "cleanup_remove_external_uri" in options
                    or "cleanup_remove_external_uri_and_text_black" in options):
                delete_it = True
            if kind == fitz.LINK_NONE and (
                    "cleanup_remove_invalid_links" in options
                    or "cleanup_remove_invalid_links_and_text_black" in options):
                delete_it = True
            if "cleanup_remove_unknown_action_links" in options and kind not in [
                    fitz.LINK_GOTO, fitz.LINK_GOTOR, fitz.LINK_LAUNCH]:
                delete_it = True
            if not delete_it:
                continue

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
            if kind == fitz.LINK_URI:
                ctx.mark("已删除外部URI链接", count=1)
            elif kind == fitz.LINK_NONE:
                ctx.mark("已删除失效链接", count=1)
            else:
                ctx.mark("已删除未知动作链接", count=1)

        if (
            "cleanup_remove_external_uri_and_text_black" in options
            or "cleanup_remove_invalid_links_and_text_black" in options
        ):
            _decolor_rects_to_black(ctx, page, decolor_rects)


def _step_cleanup(ctx):
    """清理规则组：链接/书签/批注/附件/标签/元数据/动态内容。"""
    options = ctx.options
    selected_cleanup_opts = {opt for opt in options if opt in _CLEANUP_PAGE_OPTIONS}
    if not selected_cleanup_opts:
        return
    doc = ctx.doc

    with ctx.prof.phase("清理遍历"):
        if selected_cleanup_opts.issubset(_EXTERNAL_URI_OPTS):
            _cleanup_external_uri_fast_path(ctx)
        elif "cleanup_remove_all_links_bookmarks" in options:
            _cleanup_all_links_bookmarks(ctx)
        else:
            _cleanup_links_general_path(ctx)

    if "cleanup_remove_annotations" in options:
        with ctx.prof.phase("删除注释遍历"):
            for page in doc:
                annots = list(page.annots() or [])
                for annot in annots:
                    try:
                        page.delete_annot(annot)
                        ctx.mark("已删除PDF注释", count=1)
                    except Exception:
                        pass

    if "cleanup_remove_dynamic_content" in options:
        doc.xref_set_key(ctx.catalog_xref, "Names", "null")
        ctx.mark("已删除动态内容/JavaScript")
    if "cleanup_remove_attachments" in options:
        if doc.embfile_count() > 0:
            attachment_count = doc.embfile_count()
            for emb in doc.embfile_names():
                doc.embfile_del(emb)
            ctx.mark("已删除文档附件", count=attachment_count)
    if "cleanup_remove_tags" in options:
        doc.xref_set_key(ctx.catalog_xref, "StructTreeRoot", "null")
        doc.xref_set_key(ctx.catalog_xref, "MarkInfo", "null")
        ctx.mark("已删除文档标签")
    if "cleanup_remove_metadata" in options:
        doc.set_metadata({})
        doc.xref_set_key(ctx.catalog_xref, "PieceInfo", "null")
        ctx.mark("已删除文档元数据")


# 处理管线：按序执行。步骤内部自行判断选项是否命中。
_PIPELINE_STEPS = (
    _step_document_title,
    _step_open_action,
    _step_page_layout_reset,
    _step_page_mode,
    _step_collapse_bookmarks,
    _step_page_size,
    _step_bookmark_rules,
    _step_hyperlink_rules,
    _step_cleanup,
)


def _finalize_output(ctx, input_path, output_path):
    """保存阶段：按需经 qpdf 重写（线性化/版本转换/解限制），否则直接保存或复制。"""
    options = ctx.options
    doc = ctx.doc
    prof = ctx.prof

    is_linear = "fast_web_view" in options
    force_pdf_version = "1.7" if "convert_pdf_version" in options else None
    remove_pdf_restrictions = "remove_pdf_restrictions" in options
    needs_qpdf_rewrite = bool(is_linear or force_pdf_version or remove_pdf_restrictions)

    def _mark_qpdf_changes():
        if remove_pdf_restrictions:
            _mark_change(ctx.applied_changes, "已解除PDF权限限制")
        if force_pdf_version:
            _mark_change(ctx.applied_changes, "已转换PDF版本")
        if is_linear:
            _mark_change(ctx.applied_changes, "已启用快速网页浏览")

    if ctx.changed:
        if needs_qpdf_rewrite:
            temp_pdf = str(output_path) + ".tmp.pdf"
            with prof.phase("保存(deflate+garbage2+objstms)"):
                doc.save(temp_pdf, garbage=2, deflate=True, use_objstms=1)
                doc.close()
            try:
                with prof.phase("qpdf重写"):
                    qpdf._rewrite_with_qpdf(
                        temp_pdf,
                        output_path,
                        force_version=force_pdf_version,
                        linearize=is_linear,
                        decrypt_restrictions=remove_pdf_restrictions,
                    )
                _mark_qpdf_changes()
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
            qpdf._rewrite_with_qpdf(
                input_path,
                output_path,
                force_version=force_pdf_version,
                linearize=is_linear,
                decrypt_restrictions=remove_pdf_restrictions,
            )
            _mark_qpdf_changes()
        else:
            shutil.copy2(input_path, output_path)


def process_document(input_path, output_path, options, processing_mode="smart"):
    """处理单个 PDF：解析选项 → 依序执行规则步骤 → 保存输出。

    返回 (success: bool, message: str)。message 为用户可见的处理结果文本。
    """
    prof = _PhaseProfiler()
    try:
        with prof.phase("预检/选项解析"):
            mode_resolution = precheck.resolve_processing_options(input_path, options, processing_mode)
        options = set(mode_resolution["options"])
        mode_log = mode_resolution.get("log", "")

        with prof.phase("打开文档"):
            doc = fitz.open(input_path)
        prof.set_page_count(doc.page_count)

        if doc.needs_pass:
            return False, "❌ 文件已加密"

        ctx = _PipelineContext(doc, options, input_path, prof)
        for step in _PIPELINE_STEPS:
            step(ctx)

        _finalize_output(ctx, input_path, output_path)

        if ctx.applied_changes:
            result_msg = f"✅ 处理成功；修改项：{_format_change_summary(ctx.change_counts, ctx.applied_changes)}"
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
            return False, f"❌ 处理失败: {qpdf._format_qpdf_error(e)}"
        return False, f"❌ 处理失败: {str(e)}"


class PDFProcessor:
    """pdf 层对外入口（controllers 只应使用本类的公开方法）。

    历史上此类还挂着 ~50 个转发内部函数的 static method 与常量别名，
    现已收缩：内部函数请直接从 ratools_pdf.pdf 各子模块导入。
    """

    # --- 批量处理主入口 ---
    process_document = staticmethod(process_document)

    # --- 预检 ---
    build_precheck_report = staticmethod(precheck.build_precheck_report)
    resolve_processing_options = staticmethod(precheck.resolve_processing_options)
    _pdf_has_signature = staticmethod(precheck._pdf_has_signature)

    # --- 书签 / 链接导入导出 ---
    export_bookmarks = staticmethod(bookmarks_links.export_bookmarks)
    import_bookmarks = staticmethod(bookmarks_links.import_bookmarks)
    export_links = staticmethod(bookmarks_links.export_links)
    import_links = staticmethod(bookmarks_links.import_links)
