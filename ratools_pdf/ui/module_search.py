"""功能模块与规则即时搜索引擎（纯标准库，无 Qt / fitz 依赖）。

负责跨模块检索处理规则、无损压缩配置与专用工具入口，并为结果列表提供
相关性排序及安全的 HTML 高亮片段生成。
"""

from __future__ import annotations

import html
from dataclasses import dataclass

from ratools_pdf.config import rules_catalog

LOSSLESS_OPTION_IDS = frozenset({
    "compress_standard",
    "compress_aggressive",
    "generate_object_streams",
    "realloc_flate",
})

MODULE_KEYWORDS: dict[str, str] = {
    "初始视图与文档属性": "视图 打开 首页 缩放 导览 属性 元数据 标题 initial view metadata",
    "页面与字体标准化": "页面 尺寸 大小 纸张 字体 旋转 批注 固化 展平 a4 letter font",
    "书签管理": "书签 目录 大纲 跳转 bookmark outline",
    "超链接处理": "超链接 链接 路径 边框 蓝字 样式 link hyperlink border",
    "内容合规与安全性": "合规 安全 清理 删除 移除 水印 批注 元数据 附件 脚本 watermark security",
    "文件级优化与输出": "压缩 瘦身 体积 大小 无损 有损 dpi jpeg 线性化 版本 文件名 优化 输出 ectd",
    "文档检测": "检测 扫描 只读 批注 失效 引用 链接文本 inspect check",
}

OPTION_KEYWORDS: dict[str, str] = {
    "open_page_first": "首页 第一页 打开",
    "page_layout_default": "页面布局 默认",
    "zoom_default": "缩放比例 默认 视图",
    "initial_view_bookmarks_and_page": "导览 书签面板 页面",
    "collapse_all_bookmarks": "折叠 收起 书签",
    "title_from_filename": "文件名 标题 元数据 title",
    "page_size_a4": "a4 纸张 尺寸 页面大小",
    "page_size_letter": "letter 信纸 纸张 尺寸 页面大小",
    "flatten_rotation": "旋转 展平 归一化 方向",
    "flatten_annotations": "展平 固化 批注 注释 高亮 印章",
    "bookmark_inherit_zoom": "书签 承前缩放 inherit zoom",
    "bookmark_open_new_window": "书签 新窗口 打开",
    "bookmark_remove_external_links": "书签 外部链接 uri",
    "bookmark_remove_invalid": "书签 失效 空书签",
    "bookmark_remove_unknown_actions": "书签 非标准 动作",
    "link_abs_to_rel_path": "链接 绝对路径 相对路径",
    "link_inherit_zoom": "超链接 链接 承前缩放 inherit zoom",
    "link_open_new_window": "超链接 链接 新窗口 打开",
    "link_text_blue": "超链接 链接 蓝色 文本颜色",
    "link_black_border": "超链接 链接 黑框 边框 1px",
    "link_bordered_to_blue_border": "有框 蓝框 黑字 标准化",
    "link_unbordered_blue_to_blue_border": "无框 蓝字 蓝框 黑字 标准化",
    "link_remove_border": "清除 去除 链接 边框",
    "cleanup_remove_external_uri": "删除 外部 uri 链接 网址 邮箱",
    "cleanup_remove_external_uri_and_text_black": "删除 外部 uri 链接 去色 黑色",
    "cleanup_remove_invalid_links": "清理 失效 超链接 空链接",
    "cleanup_remove_invalid_links_and_text_black": "清理 失效 链接 去色 黑色",
    "cleanup_remove_unknown_action_links": "清理 非标准 动作 链接",
    "cleanup_remove_dynamic_content": "清除 动态内容 javascript js 3d 脚本 安全",
    "cleanup_remove_attachments": "移除 附件 内嵌 zip xml",
    "cleanup_remove_tags": "结构化 标签 structtreeroot markinfo",
    "cleanup_remove_annotations": "清理 高亮 批注 注释 画笔",
    "cleanup_remove_metadata": "清空 元数据 标题 作者 pieceinfo metadata",
    "cleanup_remove_all_links_bookmarks": "移除 全部 链接 书签 高风险",
    "convert_pdf_version": "pdf 版本 1.7 转换",
    "fast_web_view": "线性化 快速网页浏览 web 流式",
    "generate_object_streams": "对象流 object streams 压缩 无损 瘦身",
    "realloc_flate": "flate deflate 重压缩 无损 瘦身",
    "compress_standard": "标准 压缩 智能无损 瘦身 体积",
    "compress_aggressive": "深度 极限无损 压缩 garbage clean 体积",
    "compress_images": "图像 图片 压缩 dpi jpeg 降采样 扫描件 有损",
    "filename_ectd_format": "ectd 文件名 格式化 小写 合规",
}

MODULE_TOOLS: tuple[dict, ...] = (
    {
        "id": "embed_missing_fonts",
        "module": "页面与字体标准化",
        "title": "嵌入缺失字体",
        "desc": "打开选中的 PDF，并在 Acrobat 中手动执行印前检查的嵌入缺失字体",
        "keywords": "字体 嵌入 acrobat 印前 缺失字体 font embed",
    },
    {
        "id": "bookmark_io_wizard",
        "module": "书签管理",
        "title": "书签数据导入/导出",
        "desc": "批量导出或导入 PDF 书签结构（CSV / Excel）",
        "keywords": "书签 导出 导入 csv bookmark export import",
    },
    {
        "id": "link_io_wizard",
        "module": "超链接处理",
        "title": "链接数据导入/导出",
        "desc": "批量导出或导入 PDF 超链接数据",
        "keywords": "超链接 链接 导出 导入 csv link hyperlink export import",
    },
    {
        "id": "remove_watermark",
        "module": "内容合规与安全性",
        "title": "去除水印",
        "desc": "对单个 PDF 进行交互式对象分析与前后对比预览，确认后另存为新文件",
        "keywords": "水印 去水印 去除水印 删除水印 watermark",
    },
    {
        "id": "detect_annotations",
        "module": "文档检测",
        "title": "检测批注 (便签/高亮等)",
        "desc": "扫描队列中所有 PDF，报告便签、高亮等批注情况，仅检测不修改文件",
        "keywords": "批注 注释 便签 高亮 检测 扫描 annotation comment",
    },
    {
        "id": "detect_broken_refs",
        "module": "文档检测",
        "title": "检测失效引用/链接文本",
        "desc": "扫描队列中所有 PDF，检测 Word 转 PDF 后残留的“错误！未找到引用源”等失效占位文本",
        "keywords": "失效 引用 链接文本 错误 未找到引用源 word 检测 broken reference",
    },
    {
        "id": "security_center",
        "module": None,
        "title": "加密与解密",
        "desc": "PDF 加密、解密、解除权限限制与安全属性管理",
        "keywords": "安全 安全中心 加密 解密 密码 权限 解除限制 decrypt encrypt password security",
    },
    {
        "id": "global_settings",
        "module": None,
        "title": "全局设置",
        "desc": "配置外观主题、界面字号、默认保存位置、并行处理与每文件超时",
        "keywords": "设置 主题 深色 浅色 字体 字号 输出目录 并行 超时 settings theme font",
    },
)


@dataclass(frozen=True)
class SearchItem:
    id: str
    title: str
    desc: str
    kind: str  # "rule" | "lossless_rule" | "tool"
    module_index: int
    module_title: str
    score: int


@dataclass(frozen=True)
class SearchGroup:
    module_index: int
    module_title: str
    module_matched: bool
    items: tuple[SearchItem, ...]


@dataclass(frozen=True)
class SearchResult:
    query: str
    tokens: tuple[str, ...]
    groups: tuple[SearchGroup, ...]
    matched_module_indices: frozenset[int]
    total_matches: int

    @property
    def is_active(self) -> bool:
        return bool(self.tokens)


def tokenize_query(query: str) -> list[str]:
    """将用户输入拆分为去重后的小写关键词列表，保留原始顺序。"""
    if not query:
        return []
    tokens: list[str] = []
    seen: set[str] = set()
    for part in str(query).strip().split():
        low = part.lower()
        if low and low not in seen:
            seen.add(low)
            tokens.append(low)
    return tokens


def _matches_all(haystack: str, tokens: list[str]) -> bool:
    low = haystack.lower()
    return all(tok in low for tok in tokens)


def _score_item(
    tokens: list[str],
    title: str,
    desc: str,
    item_id: str,
    extra_keywords: str,
    module_title: str,
) -> int:
    title_l = title.lower()
    desc_l = desc.lower()
    id_l = item_id.lower()
    kw_l = extra_keywords.lower()
    mod_l = module_title.lower()

    score = 0
    for tok in tokens:
        if tok in title_l:
            score += 40
        elif tok in kw_l:
            score += 24
        elif tok in desc_l:
            score += 18
        elif tok in id_l:
            score += 14
        elif tok in mod_l:
            score += 8
    if title_l.startswith(tokens[0]):
        score += 15
    return score


def search_modules(query: str, modules=None) -> SearchResult:
    """跨模块检索规则与专用工具；全部关键词均匹配（AND 语义）。"""
    tokens = tokenize_query(query)
    if not tokens:
        return SearchResult(
            query=query or "",
            tokens=(),
            groups=(),
            matched_module_indices=frozenset(),
            total_matches=0,
        )

    catalog = rules_catalog.MODULES if modules is None else modules
    module_groups: list[SearchGroup] = []
    matched_indices: set[int] = set()

    for idx, mod in enumerate(catalog):
        mod_title = mod["title"]
        mod_kw = MODULE_KEYWORDS.get(mod_title, "")
        mod_haystack = f"{mod_title} {mod_kw}"
        module_matched = _matches_all(mod_haystack, tokens)

        candidates: list[tuple[int, int, SearchItem]] = []
        order = 0

        for opt in mod.get("options", []):
            opt_id = opt["id"]
            display_title = opt["title"].replace(" ⚠️", "") if opt_id == "compress_images" else opt["title"]
            opt_desc = opt["desc"]
            opt_kw = OPTION_KEYWORDS.get(opt_id, "")
            haystack = f"{mod_title} {display_title} {opt_desc} {opt_id} {opt_kw}"
            if _matches_all(haystack, tokens):
                kind = "lossless_rule" if opt_id in LOSSLESS_OPTION_IDS else "rule"
                score = _score_item(tokens, display_title, opt_desc, opt_id, opt_kw, mod_title)
                candidates.append((
                    -score,
                    order,
                    SearchItem(
                        id=opt_id,
                        title=display_title,
                        desc=opt_desc,
                        kind=kind,
                        module_index=idx,
                        module_title=mod_title,
                        score=score,
                    ),
                ))
            order += 1

        for tool in MODULE_TOOLS:
            if tool["module"] != mod_title:
                continue
            tool_id = tool["id"]
            tool_title = tool["title"]
            tool_desc = tool["desc"]
            tool_kw = tool["keywords"]
            haystack = f"{mod_title} {tool_title} {tool_desc} {tool_id} {tool_kw}"
            if _matches_all(haystack, tokens):
                score = _score_item(tokens, tool_title, tool_desc, tool_id, tool_kw, mod_title)
                candidates.append((
                    -score,
                    order,
                    SearchItem(
                        id=tool_id,
                        title=tool_title,
                        desc=tool_desc,
                        kind="tool",
                        module_index=idx,
                        module_title=mod_title,
                        score=score,
                    ),
                ))
            order += 1

        if candidates or module_matched:
            matched_indices.add(idx)
            candidates.sort(key=lambda entry: (entry[0], entry[1]))
            module_groups.append(
                SearchGroup(
                    module_index=idx,
                    module_title=mod_title,
                    module_matched=module_matched,
                    items=tuple(entry[2] for entry in candidates),
                )
            )

    # 模块标题命中的分组优先排在前面，其余保持目录顺序
    ordered_groups = [g for g in module_groups if g.module_matched]
    ordered_groups.extend(g for g in module_groups if not g.module_matched)

    # 全局工具分组（安全中心 / 全局设置）
    global_candidates: list[tuple[int, int, SearchItem]] = []
    for order, tool in enumerate(t for t in MODULE_TOOLS if t["module"] is None):
        haystack = f"快捷工具 {tool['title']} {tool['desc']} {tool['id']} {tool['keywords']}"
        if _matches_all(haystack, tokens):
            score = _score_item(tokens, tool["title"], tool["desc"], tool["id"], tool["keywords"], "快捷工具")
            global_candidates.append((
                -score,
                order,
                SearchItem(
                    id=tool["id"],
                    title=tool["title"],
                    desc=tool["desc"],
                    kind="tool",
                    module_index=-1,
                    module_title="快捷工具",
                    score=score,
                ),
            ))
    if global_candidates:
        global_candidates.sort(key=lambda entry: (entry[0], entry[1]))
        ordered_groups.append(
            SearchGroup(
                module_index=-1,
                module_title="快捷工具",
                module_matched=False,
                items=tuple(entry[2] for entry in global_candidates),
            )
        )

    total_matches = sum(len(g.items) for g in ordered_groups)
    return SearchResult(
        query=query or "",
        tokens=tuple(tokens),
        groups=tuple(ordered_groups),
        matched_module_indices=frozenset(matched_indices),
        total_matches=total_matches,
    )


def highlight_matches(text: str, tokens: list[str] | tuple[str, ...], bg_hex: str, fg_hex: str) -> str:
    """安全转义文本并对命中关键词片段包裹高亮 span。"""
    source = text or ""
    clean_tokens = [t for t in tokens if t]
    if not source or not clean_tokens:
        return html.escape(source)

    lower = source.lower()
    spans: list[tuple[int, int]] = []
    for tok in clean_tokens:
        tok_low = tok.lower()
        t_len = len(tok_low)
        start = 0
        while True:
            pos = lower.find(tok_low, start)
            if pos < 0:
                break
            spans.append((pos, pos + t_len))
            start = pos + t_len

    if not spans:
        return html.escape(source)

    spans.sort()
    merged: list[list[int]] = []
    for s, e in spans:
        if not merged or s > merged[-1][1]:
            merged.append([s, e])
        else:
            merged[-1][1] = max(merged[-1][1], e)

    parts: list[str] = []
    cursor = 0
    for s, e in merged:
        if s > cursor:
            parts.append(html.escape(source[cursor:s]))
        parts.append(
            f'<span style="background-color:{bg_hex}; color:{fg_hex}; font-weight:600;">'
            f"{html.escape(source[s:e])}</span>"
        )
        cursor = e
    if cursor < len(source):
        parts.append(html.escape(source[cursor:]))
    return "".join(parts)
