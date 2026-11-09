import hashlib
import json
import tempfile
import unittest
import subprocess
import sys
from pathlib import Path
from unittest.mock import patch

import fitz

try:
    from ratools_pdf.services.qualification import MANIFEST_SCHEMA_VERSION, run_qualification
except ModuleNotFoundError:
    MANIFEST_SCHEMA_VERSION = "1.0"
    run_qualification = None


class PdfQualificationTests(unittest.TestCase):
    def setUp(self):
        self.assertIsNotNone(run_qualification, "PD-01 verifier is not implemented")

    def _make_pdf(self, path: Path) -> str:
        with fitz.open() as document:
            page = document.new_page(width=300, height=400)
            page.insert_text((30, 40), "PD-01 synthetic CJK 样本", fontname="china-s")
            document.save(path)
        return hashlib.sha256(path.read_bytes()).hexdigest()

    def _write_manifest(self, root: Path, sample: dict) -> Path:
        manifest = {
            "schema_version": MANIFEST_SCHEMA_VERSION,
            "matrix_version": "2026-11-09",
            "samples_root": "samples",
            "samples": [sample],
            "matrix": [
                {
                    "id": "auto-synthetic",
                    "evidence_level": "automated_regression",
                    "sample_ids": [sample["id"]],
                    "status": "pending",
                }
            ],
        }
        path = root / "manifest.json"
        path.write_text(json.dumps(manifest, ensure_ascii=False), encoding="utf-8")
        return path

    def _sample(self, sample_path: str, sha256: str = "") -> dict:
        return {
            "id": "synthetic-one-page",
            "path": sample_path,
            "source": {
                "type": "synthetic",
                "tool": "PyMuPDF",
                "version": fitz.VersionBind,
            },
            "authorization": "synthetic fixture; no private document content",
            "sha256": sha256,
            "features": ["CJK", "render"],
            "processing_options": {"mode": "read-only"},
            "expected": {"page_count": 1, "render_pages": [1]},
        }

    def test_verifies_hash_structure_render_and_qpdf_for_synthetic_sample(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            samples = root / "samples"
            samples.mkdir()
            pdf = samples / "minimal.pdf"
            digest = self._make_pdf(pdf)
            manifest = self._write_manifest(root, self._sample("minimal.pdf", digest))

            report = run_qualification(manifest, root / "evidence")

            result = report["results"][0]
            self.assertEqual(result["status"], "passed")
            self.assertEqual(result["checks"]["sha256"]["status"], "passed")
            self.assertEqual(result["checks"]["structure"]["page_count"], 1)
            self.assertEqual(result["checks"]["render"]["status"], "passed")
            self.assertEqual(result["checks"]["qpdf"]["status"], "passed")
            self.assertTrue((root / "evidence" / "qualification-report.json").exists())
            self.assertEqual(digest, hashlib.sha256(pdf.read_bytes()).hexdigest())

    def test_missing_sample_is_pending_and_does_not_pass(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            samples = root / "samples"
            samples.mkdir()
            manifest = self._write_manifest(root, self._sample("missing.pdf", "0" * 64))

            report = run_qualification(manifest, root / "evidence")

            self.assertEqual(report["results"][0]["status"], "pending")
            self.assertEqual(report["summary"]["passed"], 0)
            self.assertEqual(report["summary"]["pending"], 1)

    def test_hash_mismatch_fails_before_pdf_checks(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            samples = root / "samples"
            samples.mkdir()
            pdf = samples / "minimal.pdf"
            self._make_pdf(pdf)
            manifest = self._write_manifest(root, self._sample("minimal.pdf", "0" * 64))

            report = run_qualification(manifest, root / "evidence")

            result = report["results"][0]
            self.assertEqual(result["status"], "failed")
            self.assertEqual(result["checks"]["sha256"]["status"], "failed")
            self.assertNotIn("structure", result["checks"])

    def test_path_escape_fails_and_is_not_opened(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            samples = root / "samples"
            samples.mkdir()
            outside = root / "outside.pdf"
            digest = self._make_pdf(outside)
            manifest = self._write_manifest(root, self._sample("../outside.pdf", digest))

            report = run_qualification(manifest, root / "evidence")

            result = report["results"][0]
            self.assertEqual(result["status"], "failed")
            self.assertIn("containment", result["failures"][0])

    def test_existing_output_directory_is_rejected(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            samples = root / "samples"
            samples.mkdir()
            pdf = samples / "minimal.pdf"
            digest = self._make_pdf(pdf)
            manifest = self._write_manifest(root, self._sample("minimal.pdf", digest))
            output = root / "evidence"
            output.mkdir()

            with self.assertRaises(FileExistsError):
                run_qualification(manifest, output)

    def test_qpdf_warning_remains_pending_with_warning_detail(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            samples = root / "samples"
            samples.mkdir()
            pdf = samples / "minimal.pdf"
            digest = self._make_pdf(pdf)
            manifest = self._write_manifest(root, self._sample("minimal.pdf", digest))

            with patch(
                "ratools_pdf.services.qualification.check_pdf_syntax",
                return_value={
                    "is_healthy": False,
                    "has_warnings": True,
                    "returncode": 3,
                    "output": "qpdf warning: damaged stream recovered",
                },
            ):
                report = run_qualification(manifest, root / "evidence")

            result = report["results"][0]
            self.assertEqual(result["status"], "pending")
            self.assertEqual(result["checks"]["qpdf"]["status"], "pending")
            self.assertEqual(result["checks"]["qpdf"]["outcome"], "warning")
            self.assertTrue(result["warnings"])

    def test_damaged_pdf_fails_with_unchanged_input_hash(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / "samples").mkdir()
            pdf = root / "samples" / "broken.pdf"
            pdf.write_bytes(b"synthetic damaged PDF, not document content")
            digest = hashlib.sha256(pdf.read_bytes()).hexdigest()
            manifest = self._write_manifest(root, self._sample("broken.pdf", digest))
            report = run_qualification(manifest, root / "evidence")
            self.assertEqual(report["results"][0]["status"], "failed")
            self.assertEqual(report["results"][0]["checks"]["unchanged"]["status"], "passed")

    def test_disabled_qpdf_keeps_qualification_pending(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / "samples").mkdir()
            digest = self._make_pdf(root / "samples" / "minimal.pdf")
            manifest = self._write_manifest(root, self._sample("minimal.pdf", digest))
            report = run_qualification(manifest, root / "evidence", qpdf_enabled=False)
            self.assertEqual(report["results"][0]["checks"]["qpdf"]["status"], "pending")
            self.assertEqual(report["status"], "pending")

    def test_external_evidence_rows_are_not_passed_by_automated_checks(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / "samples").mkdir()
            digest = self._make_pdf(root / "samples" / "minimal.pdf")
            manifest = self._write_manifest(root, self._sample("minimal.pdf", digest))
            content = json.loads(manifest.read_text(encoding="utf-8"))
            for level in ("windows_release", "real_business_sample", "external_reader"):
                content["matrix"].append({"id": level, "evidence_level": level,
                    "sample_ids": ["synthetic-one-page"], "status": "passed"})
            manifest.write_text(json.dumps(content), encoding="utf-8")
            report = run_qualification(manifest, root / "evidence")
            self.assertEqual(report["matrix"][0]["status"], "passed")
            self.assertEqual([row["status"] for row in report["matrix"][1:]], ["pending"] * 3)

    def test_cli_returns_failure_code_for_hash_mismatch_and_writes_report(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / "samples").mkdir()
            self._make_pdf(root / "samples" / "minimal.pdf")
            manifest = self._write_manifest(root, self._sample("minimal.pdf", "0" * 64))
            result = subprocess.run([sys.executable, "-B", "scripts/run_qualification.py",
                "--manifest", str(manifest), "--output", str(root / "evidence")],
                capture_output=True, text=True, encoding="utf-8", timeout=60)
            self.assertEqual(result.returncode, 1, result.stdout + result.stderr)
            self.assertEqual(json.loads((root / "evidence" / "qualification-report.json").read_text(
                encoding="utf-8"))["status"], "failed")

    def test_unsupported_manifest_version_is_rejected_without_output(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            manifest = self._write_manifest(root, self._sample("missing.pdf"))
            content = json.loads(manifest.read_text(encoding="utf-8"))
            content["schema_version"] = "2.0"
            manifest.write_text(json.dumps(content), encoding="utf-8")
            with self.assertRaises(ValueError):
                run_qualification(manifest, root / "evidence")
            self.assertFalse((root / "evidence").exists())

    def test_engine_version_unavailable_cannot_report_passed(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / "samples").mkdir()
            digest = self._make_pdf(root / "samples" / "minimal.pdf")
            manifest = self._write_manifest(root, self._sample("minimal.pdf", digest))
            with patch("ratools_pdf.services.qualification.run_qpdf_command", side_effect=OSError):
                report = run_qualification(manifest, root / "evidence")
            self.assertEqual(report["engines"]["qpdf"]["status"], "blocked")
            self.assertEqual(report["status"], "blocked")

    def test_reparse_or_hardlink_sample_is_rejected_before_hashing(self):
        # Real link creation would violate recycle-only cleanup. Simulate only the
        # OS lstat boundary; production check_unlinked_path and verifier stay real.
        for link_kind in ("reparse", "hardlink"):
            with self.subTest(link_kind=link_kind), tempfile.TemporaryDirectory() as directory:
                root = Path(directory)
                (root / "samples").mkdir()
                pdf = root / "samples" / "minimal.pdf"
                digest = self._make_pdf(pdf)
                manifest = self._write_manifest(root, self._sample("minimal.pdf", digest))
                real_lstat = Path.lstat
                from types import SimpleNamespace
                def linked_lstat(candidate, *args, **kwargs):
                    info = real_lstat(candidate, *args, **kwargs)
                    if candidate == pdf:
                        attributes = {name: getattr(info, name) for name in dir(info) if name.startswith("st_")}
                        attributes.update(st_file_attributes=0x400 if link_kind == "reparse" else 0,
                            st_nlink=2 if link_kind == "hardlink" else 1)
                        return SimpleNamespace(**attributes)
                    return info
                with patch.object(Path, "lstat", linked_lstat):
                    report = run_qualification(manifest, root / "evidence")
                self.assertEqual(report["results"][0]["status"], "failed")
                self.assertNotIn("sha256", report["results"][0]["checks"])

    def test_unknown_expected_metric_blocks_instead_of_being_ignored(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / "samples").mkdir()
            digest = self._make_pdf(root / "samples" / "minimal.pdf")
            sample = self._sample("minimal.pdf", digest)
            sample["expected"]["reader_visual_match"] = True
            manifest = self._write_manifest(root, sample)
            report = run_qualification(manifest, root / "evidence")
            self.assertEqual(report["results"][0]["status"], "blocked")

    def test_empty_expected_metrics_cannot_qualify_by_render_alone(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / "samples").mkdir()
            digest = self._make_pdf(root / "samples" / "minimal.pdf")
            sample = self._sample("minimal.pdf", digest)
            sample["expected"] = {"render_pages": [1]}
            manifest = self._write_manifest(root, sample)
            report = run_qualification(manifest, root / "evidence")
            self.assertEqual(report["results"][0]["status"], "pending")

    def test_absolute_and_drive_relative_paths_fail_containment(self):
        for relative in ("C:/not-a-sample.pdf", "C:not-a-sample.pdf", "/not-a-sample.pdf", "x.pdf:alternate"):
            with self.subTest(path=relative), tempfile.TemporaryDirectory() as directory:
                root = Path(directory)
                manifest = self._write_manifest(root, self._sample(relative, "0" * 64))
                report = run_qualification(manifest, root / "evidence")
                self.assertEqual(report["results"][0]["status"], "failed")

    def test_original_hash_change_during_check_is_reported_failed(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / "samples").mkdir()
            pdf = root / "samples" / "minimal.pdf"
            digest = self._make_pdf(pdf)
            manifest = self._write_manifest(root, self._sample("minimal.pdf", digest))
            def changed_input(*args, **kwargs):
                with pdf.open("ab") as stream:
                    stream.write(b"\n% simulated concurrent external change\n")
                return {"returncode": 0, "output": "clean"}
            with patch("ratools_pdf.services.qualification.check_pdf_syntax", side_effect=changed_input):
                report = run_qualification(manifest, root / "evidence")
            self.assertEqual(report["results"][0]["status"], "failed")
            self.assertEqual(report["results"][0]["checks"]["unchanged"]["status"], "failed")


if __name__ == "__main__":
    unittest.main()
