"""线性 SVG 图标库纯 Python 校验（不依赖 PySide6）。"""

import ast
import inspect
import pathlib
import textwrap
import unittest
import xml.etree.ElementTree as ET

from ratools_pdf.config import rules_catalog
from ratools_pdf.ui import icons


class SvgIconCatalogTests(unittest.TestCase):
    def test_importing_icons_does_not_require_pyside6(self):
        icons_path = pathlib.Path(icons.__file__)
        tree = ast.parse(icons_path.read_text(encoding="utf-8"))
        top_level_modules = []
        for node in tree.body:
            if isinstance(node, ast.Import):
                top_level_modules.extend(alias.name for alias in node.names)
            elif isinstance(node, ast.ImportFrom) and node.module:
                top_level_modules.append(node.module)
        self.assertFalse(
            any(mod.startswith("PySide6") for mod in top_level_modules),
            f"Unexpected top-level PySide6 import: {top_level_modules}",
        )

    def test_every_svg_definition_renders_well_formed_xml(self):
        self.assertTrue(icons.SVG_PATHS)
        for name in icons.SVG_PATHS:
            with self.subTest(icon=name):
                xml_text = icons.build_svg_xml(name, "#0F6CBD")
                root = ET.fromstring(xml_text)
                self.assertTrue(root.tag.endswith("svg"))
                self.assertEqual("0 0 24 24", root.attrib.get("viewBox"))

    def test_module_icon_map_covers_all_rules_catalog_modules(self):
        for module in rules_catalog.MODULES:
            title = module["title"]
            with self.subTest(module=title):
                self.assertIn(title, icons.MODULE_ICON_MAP)
                icon_name = icons.MODULE_ICON_MAP[title]
                self.assertIn(icon_name, icons.SVG_PATHS)

    def test_render_svg_pixmap_uses_explicit_target_rect_and_sets_dpr_after_end(self):
        src = textwrap.dedent(inspect.getsource(icons.render_svg_pixmap))
        tree = ast.parse(src)
        func = tree.body[0]

        render_calls = []
        end_idx = None
        dpr_idx = None
        for idx, stmt in enumerate(func.body):
            if isinstance(stmt, ast.Expr) and isinstance(stmt.value, ast.Call):
                call = stmt.value
                if isinstance(call.func, ast.Attribute):
                    if call.func.attr == "render":
                        render_calls.append(call)
                    elif call.func.attr == "end":
                        end_idx = idx
                    elif call.func.attr == "setDevicePixelRatio":
                        dpr_idx = idx

        self.assertEqual(1, len(render_calls))
        self.assertEqual(2, len(render_calls[0].args), "renderer.render must pass an explicit QRectF target")
        rect_arg = render_calls[0].args[1]
        self.assertIsInstance(rect_arg, ast.Call)
        self.assertIsInstance(rect_arg.func, ast.Name)
        self.assertEqual("QRectF", rect_arg.func.id)
        self.assertIsNotNone(end_idx)
        self.assertIsNotNone(dpr_idx)
        self.assertLess(end_idx, dpr_idx, "setDevicePixelRatio must run after painter.end()")


if __name__ == "__main__":
    unittest.main()
