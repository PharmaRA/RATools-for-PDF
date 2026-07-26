"""rules_catalog 一致性守卫：预设/预检引用的规则 id 必须存在于目录中。"""

import unittest

from ratools_pdf.config import rules_catalog
from ratools_pdf.pdf import precheck


class RulesCatalogConsistencyTests(unittest.TestCase):
    def test_option_ids_unique(self):
        ids = [
            opt["id"]
            for module in rules_catalog.MODULES
            for opt in module["options"]
        ]
        self.assertEqual(len(ids), len(set(ids)), "规则 id 出现重复")

    def test_preset_options_exist_in_catalog(self):
        known = set(rules_catalog.OPTION_TITLES)
        for key, preset in rules_catalog.PRESETS.items():
            unknown = preset["options"] - known
            self.assertFalse(unknown, f"预设 {key} 引用了目录外的规则: {unknown}")

    def test_precheck_detectable_options_exist_in_catalog(self):
        known = set(rules_catalog.OPTION_TITLES)
        unknown = precheck.PRECHECK_DETECTABLE_OPTIONS - known
        self.assertFalse(unknown, f"预检声明了目录外的规则: {unknown}")

    def test_precheck_titles_are_catalog_titles(self):
        # 防回归：precheck 的标题表必须与目录同源，不允许再出现双份维护
        for option_id, title in precheck.PRECHECK_OPTION_TITLES.items():
            self.assertEqual(title, rules_catalog.option_title(option_id))

    def test_every_option_has_title_and_desc(self):
        for module in rules_catalog.MODULES:
            for opt in module["options"]:
                self.assertTrue(opt["title"].strip(), f"{opt['id']} 缺标题")
                self.assertTrue(opt["desc"].strip(), f"{opt['id']} 缺描述")


if __name__ == "__main__":
    unittest.main()
