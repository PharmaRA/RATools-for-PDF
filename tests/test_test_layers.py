"""Behavioral checks for the safe runner, not assertions on manifest text."""

import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest


ROOT = Path(__file__).resolve().parents[1]


class TestLayerRunnerTests(unittest.TestCase):
    def fixture(self, root):
        tests = root / "tests"
        tests.mkdir()
        (tests / "test_pure.py").write_text(
            "import unittest\nclass Pure(unittest.TestCase):\n"
            " def test_runs(self): print('PURE_ACTUALLY_RAN')\n", encoding="utf-8")
        (tests / "test_native.py").write_text(
            "raise RuntimeError('UNSELECTED_MODULE_IMPORTED')\n", encoding="utf-8")
        manifest = root / "layers.json"
        manifest.write_text(json.dumps({"schema_version": "1.0", "inventory_version": "fixture-1",
            "layers": {"unit": ["test_pure.py"], "integration": ["test_native.py"]}}), encoding="utf-8")
        return tests, manifest

    def run_cli(self, *args):
        return subprocess.run([sys.executable, "-X", "utf8", "-B", str(ROOT / "scripts/run_tests_safely.py"),
            *map(str, args)], capture_output=True, text=True, encoding="utf-8", timeout=90)

    def summary(self, result):
        lines = [line for line in result.stdout.splitlines() if line.startswith("LAYER_SUMMARY ")]
        self.assertEqual(len(lines), 1, result.stdout + result.stderr)
        return json.loads(lines[0].removeprefix("LAYER_SUMMARY "))

    def test_unit_selection_runs_real_test_without_importing_other_layer(self):
        # Breaking discovery to import all modules would execute the poison module.
        with tempfile.TemporaryDirectory() as directory:
            tests, manifest = self.fixture(Path(directory))
            result = self.run_cli("--layer", "unit", "--tests-dir", tests, "--classification", manifest)
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
            self.assertIn("PURE_ACTUALLY_RAN", result.stdout)
            self.assertNotIn("UNSELECTED_MODULE_IMPORTED", result.stderr)
            summary = self.summary(result)
            self.assertEqual(summary["layers"]["unit"]["passed"], 1)
            self.assertEqual(summary["layers"]["integration"]["status"], "not_selected")
            self.assertEqual(summary["totals"], {"passed": 1, "failed": 0, "skipped": 0, "pending": 0})

    def test_default_full_and_positional_pattern_preserve_failure_exit(self):
        with tempfile.TemporaryDirectory() as directory:
            tests, manifest = self.fixture(Path(directory))
            (tests / "test_native.py").write_text(
                "import unittest\nclass Native(unittest.TestCase):\n"
                " def test_failure(self): self.fail('REAL_FAILURE')\n", encoding="utf-8")
            common = ("--tests-dir", tests, "--classification", manifest)
            full = self.run_cli(*common)
            self.assertEqual(full.returncode, 1, full.stdout + full.stderr)
            self.assertIn("REAL_FAILURE", full.stderr)
            self.assertEqual(self.summary(full)["totals"], {"passed": 1, "failed": 1, "skipped": 0, "pending": 0})
            positional = self.run_cli("test_pure.py", *common)
            self.assertEqual(positional.returncode, 0, positional.stdout + positional.stderr)
            self.assertEqual(self.summary(positional)["totals"]["passed"], 1)

    def test_missing_and_duplicate_classification_rejected_before_any_test_runs(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            tests, manifest = self.fixture(root)
            data = json.loads(manifest.read_text(encoding="utf-8"))
            for integration in ([], ["test_native.py", "test_pure.py"]):
                with self.subTest(integration=integration):
                    data["layers"]["integration"] = integration
                    manifest.write_text(json.dumps(data), encoding="utf-8")
                    result = self.run_cli("--layer", "unit", "--tests-dir", tests, "--classification", manifest)
                    self.assertEqual(result.returncode, 2, result.stdout + result.stderr)
                    self.assertIn("classification", result.stderr.lower())
                    self.assertNotIn("PURE_ACTUALLY_RAN", result.stdout)

    def test_unknown_version_duplicate_json_keys_and_empty_selection_rejected(self):
        with tempfile.TemporaryDirectory() as directory:
            tests, manifest = self.fixture(Path(directory))
            for content in (
                '{"schema_version":"9","inventory_version":"x","layers":{}}',
                '{"schema_version":"1.0","schema_version":"1.0"}',
            ):
                manifest.write_text(content, encoding="utf-8")
                result = self.run_cli("--tests-dir", tests, "--classification", manifest)
                self.assertEqual(result.returncode, 2, result.stdout + result.stderr)
            tests2 = Path(directory) / "another"
            tests2.mkdir()
            _, good_manifest = self.fixture(tests2)
            result = self.run_cli("test_absent.py", "--tests-dir", tests2 / "tests", "--classification", good_manifest)
            self.assertEqual(result.returncode, 2, result.stdout + result.stderr)

    def test_subtest_failures_and_skips_are_counted_once_per_executed_test(self):
        with tempfile.TemporaryDirectory() as directory:
            tests, manifest = self.fixture(Path(directory))
            (tests / "test_pure.py").write_text(
                "import unittest\nclass Pure(unittest.TestCase):\n"
                " def test_subtests(self):\n"
                "  for n in (1,2):\n"
                "   with self.subTest(n=n): self.assertEqual(n,0)\n"
                " @unittest.skip('material absent')\n"
                " def test_skip(self): pass\n", encoding="utf-8")
            result = self.run_cli("--layer", "unit", "--tests-dir", tests, "--classification", manifest)
            self.assertEqual(result.returncode, 1, result.stdout + result.stderr)
            self.assertEqual(self.summary(result)["totals"], {"passed": 0, "failed": 1, "skipped": 1, "pending": 0})

    def test_unit_policy_rejects_external_process_and_native_engine_import(self):
        with tempfile.TemporaryDirectory() as directory:
            tests, manifest = self.fixture(Path(directory))
            for operation in ("import fitz", "import PySide6", "import subprocess; subprocess.run(['not-a-real-command'])"):
                with self.subTest(operation=operation):
                    (tests / "test_pure.py").write_text("import unittest\nclass Pure(unittest.TestCase):\n"
                        " def test_forbidden(self): " + operation + "\n", encoding="utf-8")
                    result = self.run_cli("--layer", "unit", "--tests-dir", tests, "--classification", manifest)
                    self.assertEqual(result.returncode, 1, result.stdout + result.stderr)
                    self.assertIn("Unit layer cannot", result.stderr)
                    self.assertEqual(self.summary(result)["totals"]["passed"], 0)

    def test_setup_class_failure_reports_unrun_tests_pending_instead_of_crashing_summary(self):
        with tempfile.TemporaryDirectory() as directory:
            tests, manifest = self.fixture(Path(directory))
            (tests / "test_pure.py").write_text("import unittest\nclass Pure(unittest.TestCase):\n"
                " @classmethod\n def setUpClass(cls): raise RuntimeError('SETUP_FAILED')\n"
                " def test_unrun(self): print('MUST_NOT_RUN')\n", encoding="utf-8")
            result = self.run_cli("--layer", "unit", "--tests-dir", tests, "--classification", manifest)
            self.assertEqual(result.returncode, 1, result.stdout + result.stderr)
            self.assertNotIn("MUST_NOT_RUN", result.stdout)
            summary = self.summary(result)
            self.assertEqual(summary["totals"], {"passed": 0, "failed": 0, "skipped": 0, "pending": 1})
            self.assertEqual(summary["layers"]["unit"]["status"], "failed")
            self.assertEqual(summary["layers"]["unit"]["suite_errors"], 1)

    def test_qualification_selector_consumes_pd01_manifest_and_reports_missing_material_nonzero(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            manifest = root / "samples.json"
            manifest.write_text(json.dumps({"schema_version": "1.0", "matrix_version": "fixture-1",
                "samples_root": "samples", "samples": [{"id": "external", "path": None,
                    "source": {"type": "real", "tool": "Acrobat", "version": "pending"},
                    "authorization": "pending authorized non-sensitive sample", "features": [],
                    "processing_options": {}, "expected": {}, "sha256": None}],
                "matrix": [{"id": "reader", "evidence_level": "external_reader",
                    "sample_ids": ["external"], "status": "passed"}]}), encoding="utf-8")
            result = self.run_cli("--layer", "qualification", "--qualification-manifest", manifest,
                "--qualification-output", root / "evidence", "--skip-qpdf")
            self.assertEqual(result.returncode, 2, result.stdout + result.stderr)
            report = json.loads((root / "evidence/qualification-report.json").read_text(encoding="utf-8"))
            self.assertEqual(report["matrix"][0]["status"], "pending")
            summary = self.summary(result)
            self.assertEqual(summary["layers"]["qualification"]["pending"], 1)
            self.assertEqual(summary["layers"]["unit"]["status"], "not_selected")
            self.assertEqual(summary["totals"]["passed"], 0)
