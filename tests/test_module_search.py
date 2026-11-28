"""纯 Python 功能模块与规则搜索引擎单元测试（不依赖 PySide6 / fitz）。"""

import ast
import pathlib
import unittest

from ratools_pdf.config import rules_catalog
from ratools_pdf.ui import module_search


class ModuleSearchTests(unittest.TestCase):
    def test_module_search_does_not_import_qt_or_fitz_at_top_level(self):
        source_path = pathlib.Path(module_search.__file__)
        tree = ast.parse(source_path.read_text(encoding="utf-8"))
        top_level_modules = []
        for node in tree.body:
            if isinstance(node, ast.Import):
                top_level_modules.extend(alias.name for alias in node.names)
            elif isinstance(node, ast.ImportFrom) and node.module:
                top_level_modules.append(node.module)
        forbidden = ("PySide6", "PyQt5", "PyQt6", "fitz", "pymupdf")
        self.assertFalse(
            any(mod.startswith(forbidden) for mod in top_level_modules),
            f"Unexpected GUI/PDF import in module_search: {top_level_modules}",
        )

    def test_empty_or_whitespace_query_is_inactive(self):
        for query in ("", "   ", "\t\n"):
            res = module_search.search_modules(query)
            self.assertFalse(res.is_active)
            self.assertEqual((), res.groups)
            self.assertEqual(frozenset(), res.matched_module_indices)
            self.assertEqual(0, res.total_matches)

    def test_search_by_module_title_prioritizes_matched_module(self):
        res = module_search.search_modules("书签管理")
        self.assertTrue(res.is_active)
        self.assertGreaterEqual(len(res.groups), 1)
        first = res.groups[0]
        self.assertEqual("书签管理", first.module_title)
        self.assertTrue(first.module_matched)
        self.assertIn(first.module_index, res.matched_module_indices)
        ids = {item.id for item in first.items}
        self.assertIn("bookmark_inherit_zoom", ids)
        self.assertIn("bookmark_io_wizard", ids)

    def test_search_by_rule_title_and_description_and_case_insensitive_terms(self):
        res_a4 = module_search.search_modules("A4")
        ids_a4 = {item.id for g in res_a4.groups for item in g.items}
        self.assertIn("page_size_a4", ids_a4)

        res_js = module_search.search_modules("javascript")
        ids_js = {item.id for g in res_js.groups for item in g.items}
        self.assertIn("cleanup_remove_dynamic_content", ids_js)

        res_ectd = module_search.search_modules("ectd")
        ids_ectd = {item.id for g in res_ectd.groups for item in g.items}
        self.assertIn("filename_ectd_format", ids_ectd)

        res_dpi = module_search.search_modules("dpi")
        ids_dpi = {item.id for g in res_dpi.groups for item in g.items}
        self.assertIn("compress_images", ids_dpi)

    def test_multi_token_and_query_narrows_results(self):
        res = module_search.search_modules("书签 外部")
        ids = [item.id for g in res.groups for item in g.items]
        self.assertIn("bookmark_remove_external_links", ids)
        self.assertNotIn("bookmark_inherit_zoom", ids)

        res_blue = module_search.search_modules("链接 蓝")
        ids_blue = [item.id for g in res_blue.groups for item in g.items]
        self.assertIn("link_text_blue", ids_blue)
        self.assertNotIn("link_black_border", ids_blue)

    def test_lossless_rules_and_normal_rules_have_distinct_kinds(self):
        res_lossless = module_search.search_modules("对象流")
        items = [item for g in res_lossless.groups for item in g.items]
        obj_stream = next(item for item in items if item.id == "generate_object_streams")
        self.assertEqual("lossless_rule", obj_stream.kind)

        res_normal = module_search.search_modules("线性化")
        norm_items = [item for g in res_normal.groups for item in g.items]
        fast_web = next(item for item in norm_items if item.id == "fast_web_view")
        self.assertEqual("rule", fast_web.kind)

    def test_tools_and_global_tools_are_searchable_and_module_refs_valid(self):
        known_modules = {mod["title"] for mod in rules_catalog.MODULES}
        for tool in module_search.MODULE_TOOLS:
            if tool["module"] is not None:
                self.assertIn(tool["module"], known_modules)

        res_wm = module_search.search_modules("水印")
        wm_items = [item for g in res_wm.groups for item in g.items]
        self.assertTrue(any(item.id == "remove_watermark" and item.kind == "tool" for item in wm_items))

        res_sec = module_search.search_modules("加密")
        sec_items = [item for g in res_sec.groups for item in g.items]
        self.assertTrue(
            any(item.id == "security_center" and item.module_index == -1 and item.kind == "tool" for item in sec_items)
        )

    def test_highlight_matches_escapes_html_and_merges_overlapping_spans(self):
        highlighted = module_search.highlight_matches(
            "清除 <JS & 3D> 动态内容",
            ["js", "s &", "动态"],
            "#DCEBFA",
            "#0E4F8A",
        )
        self.assertIn("&lt;", highlighted)
        self.assertIn("&amp;", highlighted)
        self.assertIn("&gt;", highlighted)
        self.assertIn(
            '<span style="background-color:#DCEBFA; color:#0E4F8A; font-weight:600;">JS &amp;</span>',
            highlighted,
        )
        self.assertIn(
            '<span style="background-color:#DCEBFA; color:#0E4F8A; font-weight:600;">动态</span>',
            highlighted,
        )


if __name__ == "__main__":
    unittest.main()
