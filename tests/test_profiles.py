"""PD04 configuration contracts; expectations are independent literal fixtures."""

import copy
import inspect
import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
from types import SimpleNamespace

from PySide6.QtWidgets import QApplication
from ratools_pdf.ui.main_window import MainWindow

try:
    from ratools_pdf.services import profiles
except ImportError:
    profiles = None


def example_profile():
    return {
        "schema_version": 1,
        "profile_id": "custom.qa",
        "profile_version": "1.0.0",
        "agency_or_region": "Internal QA",
        "enabled_rules": ["compress_images", "open_page_first"],
        "rule_parameters": {"compress_images": {"dpi": 240, "quality": 80}},
        "validation_level": "structural",
        "source": "User configuration; no official certification",
        "applicability": "Internal workflow",
    }


class ProfileServiceTests(unittest.TestCase):
    def api(self):
        self.assertIsNotNone(profiles, "PD04 versioned profile service is missing")
        return profiles

    def test_roundtrip_keeps_exact_configuration_and_stable_digest(self):
        api = self.api()
        original = example_profile()
        exported = api.dumps_profile(original)
        restored = api.loads_profile(exported)
        self.assertEqual(restored, original)
        reordered = dict(reversed(list(original.items())))
        self.assertEqual(api.profile_digest(restored), api.profile_digest(reordered))
        altered = copy.deepcopy(original)
        altered["rule_parameters"]["compress_images"]["dpi"] = 300
        self.assertNotEqual(api.profile_digest(restored), api.profile_digest(altered))
        self.assertNotIn("output", exported)

    def test_builtin_china_and_us_keep_existing_rule_selections(self):
        api = self.api()
        common = {
            "convert_pdf_version", "fast_web_view", "initial_view_bookmarks_and_page",
            "page_layout_default", "open_page_first", "bookmark_inherit_zoom",
            "cleanup_remove_annotations", "cleanup_remove_metadata", "cleanup_remove_attachments",
            "cleanup_remove_dynamic_content", "link_inherit_zoom", "link_open_new_window",
            "bookmark_open_new_window",
        }
        self.assertEqual(set(api.builtin_profile("us")["enabled_rules"]), common)
        china = api.builtin_profile("china")
        self.assertEqual(set(china["enabled_rules"]), common | {
            "cleanup_remove_external_uri", "collapse_all_bookmarks", "filename_ectd_format",
        })
        self.assertIn("not official", china["source"].lower())
        self.assertEqual(china["profile_id"], "builtin.china")
        self.assertEqual(china["profile_version"], "1.0.0")

    def test_invalid_profiles_are_rejected_without_normalization(self):
        api = self.api()
        cases = [
            ("schema_version", 2), ("schema_version", True),
            ("profile_id", "../outside"), ("profile_version", "next"),
            ("enabled_rules", ["unknown_rule"]), ("enabled_rules", ["open_page_first"] * 2),
            ("enabled_rules", ["page_size_a4", "page_size_letter"]),
            ("enabled_rules", ["compress_standard", "compress_aggressive"]),
            ("enabled_rules", ["link_black_border", "link_remove_border"]),
            ("enabled_rules", ["flatten_annotations", "cleanup_remove_annotations"]),
            ("enabled_rules", ["cleanup_remove_all_links_bookmarks", "link_inherit_zoom"]),
            ("validation_level", "official_certified"), ("output_directory", "private"),
            ("required_fields", ["future_feature"]),
        ]
        for key, value in cases:
            with self.subTest(key=key, value=value):
                payload = example_profile()
                payload[key] = value
                with self.assertRaises(ValueError):
                    api.validate_profile(payload)
        for params in ({"dpi": 71, "quality": 80}, {"dpi": 601, "quality": 80},
                       {"dpi": "240", "quality": 80}, {"dpi": True, "quality": 80},
                       {"dpi": 240, "quality": 96}, {"dpi": 240},
                       {"dpi": 240, "quality": 80, "future": True}):
            with self.subTest(params=params):
                payload = example_profile()
                payload["rule_parameters"] = {"compress_images": params}
                with self.assertRaises(ValueError):
                    api.validate_profile(payload)

    def test_duplicate_json_keys_and_unknown_parameter_rules_are_rejected(self):
        api = self.api()
        text = json.dumps(example_profile())
        with self.assertRaises(ValueError):
            api.loads_profile(text[:-1] + ', "schema_version": 1}')
        payload = example_profile()
        payload["rule_parameters"]["open_page_first"] = {"page": 7}
        with self.assertRaises(ValueError):
            api.validate_profile(payload)

    def test_legacy_list_and_comma_favorites_migrate_without_losing_unknowns(self):
        api = self.api()
        for legacy in (["open_page_first", "fast_web_view"], "open_page_first,fast_web_view"):
            profile = api.migrate_favorite(legacy)
            self.assertEqual(profile["enabled_rules"], ["fast_web_view", "open_page_first"])
            self.assertEqual(profile["profile_id"], "favorite.legacy")
        with self.assertRaises(ValueError):
            api.migrate_favorite("open_page_first,future_rule")

    def test_diff_exposes_added_removed_rules_and_parameter_changes(self):
        api = self.api()
        before = example_profile()
        after = example_profile()
        after["enabled_rules"] = ["compress_images", "fast_web_view"]
        after["rule_parameters"]["compress_images"]["quality"] = 75
        changes = api.diff_profiles(before, after)
        self.assertEqual(changes["added_rules"], ["fast_web_view"])
        self.assertEqual(changes["removed_rules"], ["open_page_first"])
        self.assertEqual(changes["parameter_changes"], {
            "compress_images.quality": {"before": 80, "after": 75},
        })


class ProfileWindowTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.root = Path(self.directory.name)
        with patch("ratools_pdf.ui.main_window.get_app_dir", return_value=str(self.root)):
            self.window = MainWindow()
        self.addCleanup(self.window.deleteLater)

    def require_api(self):
        self.assertTrue(hasattr(self.window, "get_processing_profile"), "Desktop profiles are missing")

    def snapshot_settings(self):
        return {key: self.window.app_settings.value(key) for key in self.window.app_settings.allKeys()}

    def test_import_cancel_previews_diff_without_widget_or_settings_changes(self):
        self.require_api()
        path = self.root / "incoming.json"
        path.write_text(json.dumps(example_profile()), encoding="utf-8")
        before = self.snapshot_settings()
        rules = self.window.get_processing_options()
        previews = []
        def cancel(title, message):
            previews.append(message)
            return False
        with patch("ratools_pdf.ui.main_window.QFileDialog.getOpenFileName", return_value=(str(path), "")), \
                patch.object(self.window, "show_confirm_message", side_effect=cancel):
            self.window.btn_import_profile.click()
        self.assertEqual(self.snapshot_settings(), before)
        self.assertEqual(self.window.get_processing_options(), rules)
        self.assertEqual(len(previews), 1)
        self.assertIn("compress_images", previews[0])
        self.assertIn("240", previews[0])

    def test_accepted_import_roundtrip_exact_parameters_and_identity_invalidation(self):
        self.require_api()
        incoming = example_profile()
        self.assertTrue(self.window.apply_processing_profile(incoming))
        self.assertEqual(self.window.get_processing_profile(), incoming)
        self.assertEqual(self.window.get_compression_settings(), {"dpi": 240, "quality": 80})
        self.window.spin_image_dpi.setValue(250)
        derived = self.window.get_processing_profile()
        self.assertNotEqual(derived["profile_id"], "custom.qa")
        self.assertEqual(derived["rule_parameters"]["compress_images"]["dpi"], 250)
        self.window.apply_processing_profile(incoming)
        self.window.all_checkboxes["fast_web_view"].setChecked(True)
        self.assertNotEqual(self.window.get_processing_profile()["profile_id"], "custom.qa")

    def test_unrepresentable_compression_import_rejects_without_any_persistence(self):
        self.require_api()
        incoming = example_profile()
        incoming["enabled_rules"] = ["compress_standard"]
        incoming["rule_parameters"] = {}
        before = self.snapshot_settings()
        before_rules = self.window.get_processing_options()
        with self.assertRaises(ValueError):
            self.window.apply_processing_profile(incoming)
        self.assertEqual(self.snapshot_settings(), before)
        self.assertEqual(self.window.get_processing_options(), before_rules)

    def test_export_button_and_favorite_preserve_only_portable_profile(self):
        self.require_api()
        self.window.apply_processing_profile(example_profile())
        self.window.settings_dialog.default_output_edit.setText("private-output-location")
        self.window.all_checkboxes["覆盖原始文件 (不推荐)"].setChecked(True)
        path = self.root / "export.json"
        with patch("ratools_pdf.ui.main_window.QFileDialog.getSaveFileName", return_value=(str(path), "")), \
                patch.object(self.window, "show_success_message"):
            self.window.btn_export_profile.click()
            self.window.save_favorite_preset()
        text = path.read_text(encoding="utf-8")
        self.assertNotIn("private-output", text)
        self.assertEqual(json.loads(text)["enabled_rules"], ["compress_images", "open_page_first"])
        self.window.clear_selected_options()
        self.window.apply_preset("favorite")
        self.assertEqual(self.window.get_compression_settings(), {"dpi": 240, "quality": 80})
        self.assertEqual(self.window.get_processing_profile()["profile_id"], "custom.qa")

    def test_actual_builtin_selection_has_exact_versioned_identity(self):
        self.require_api()
        for key, count in (("china", 16), ("us", 13)):
            self.window.apply_preset(key)
            self.assertEqual(len(self.window.get_processing_options()), count)
            self.assertEqual(self.window.get_processing_profile()["profile_id"], "builtin." + key)


class ProfileEvidenceTests(unittest.TestCase):
    def test_plan_and_audit_freeze_exact_profile_version_digest_and_parameters(self):
        from ratools_pdf.services.processing_plan import build_processing_plan, configuration_matches_plan
        from ratools_pdf.services.audit_evidence import export_evidence_package, verify_evidence_package
        from ratools_pdf.controllers.main_controller import MainController
        self.assertIn("profile", inspect.signature(build_processing_plan).parameters,
                      "PD02 profile snapshot binding is missing")
        incoming = example_profile()
        plan = build_processing_plan([], incoming["enabled_rules"], compression_settings={"dpi": 240, "quality": 80},
                                     profile=incoming)
        self.assertEqual(plan["configuration"]["profile"], incoming)
        digest = plan["configuration"]["profile_digest"]
        self.assertEqual(digest, profiles.profile_digest(example_profile()))
        incoming["profile_version"] = "2.0.0"
        self.assertEqual(plan["configuration"]["profile"]["profile_version"], "1.0.0")
        self.assertFalse(configuration_matches_plan(plan, incoming["enabled_rules"], "smart",
                                                    {"dpi": 240, "quality": 80}, profile=incoming))
        self.assertTrue(configuration_matches_plan(plan, incoming["enabled_rules"], "smart",
                                                   {"dpi": 240, "quality": 80}))
        host = SimpleNamespace()
        MainController._retain_processing_plan(host, plan, "cancelled")
        MainController._finish_processing_evidence(host, "cancelled")
        with tempfile.TemporaryDirectory() as root:
            target = Path(root) / "audit"
            export_evidence_package(host.processing_evidence, target)
            exported = json.loads((target / "evidence.json").read_text(encoding="utf-8"))
            self.assertEqual(exported["configuration"]["profile"], example_profile())
            self.assertEqual(exported["configuration"]["profile_digest"], digest)
            self.assertTrue(verify_evidence_package(target)["valid"])
            host.processing_evidence["plan"]["configuration"]["profile_digest"] = "0" * 64
            with self.assertRaises(ValueError):
                export_evidence_package(host.processing_evidence, Path(root) / "tampered")

    def test_profile_snapshot_must_match_actual_engine_rules_and_parameters(self):
        from ratools_pdf.services.processing_plan import build_processing_plan
        self.assertIn("profile", inspect.signature(build_processing_plan).parameters,
                      "PD02 profile snapshot binding is missing")
        for rules, params in ((["open_page_first"], {"dpi": 240, "quality": 80}),
                             (["compress_images", "open_page_first"], {"dpi": 300, "quality": 80})):
            with self.subTest(rules=rules, params=params), self.assertRaises(ValueError):
                build_processing_plan([], rules, compression_settings=params, profile=example_profile())


if __name__ == "__main__":
    unittest.main()
