"""PD05 behavior: only a fresh real font inspection can verify a repair."""

import importlib
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

import fitz


class FontProviderTests(unittest.TestCase):
    def api(self):
        try:
            return importlib.import_module("ratools_pdf.services.font_providers")
        except ImportError:
            self.fail("PD05 unified font capability/inspect/repair/recheck API is missing")

    def make_pdf(self, directory, *, text=False):
        path = Path(directory) / "sample.pdf"
        with fitz.open() as doc:
            page = doc.new_page()
            if text:
                page.insert_text((72, 72), "Font sample")
            doc.save(path)
        return path

    def provider(self, *, platform="Windows", acrobat=""):
        api = self.api()
        with patch.object(api.system_shell, "find_acrobat_executable", return_value=acrobat):
            return api.AcrobatManualProvider(platform_name=platform)

    def test_installed_acrobat_never_unlocks_retired_automatic_repair(self):
        api = self.api()
        provider = self.provider(acrobat="installed-acrobat.exe")
        workflow = api.FontWorkflow(provider)
        capability = workflow.capability()
        result = workflow.repair("untouched.pdf")
        self.assertEqual(capability.provider, "acrobat_manual")
        self.assertEqual(capability.platform, "Windows")
        self.assertTrue(capability.acrobat_detected)
        self.assertIsNone(capability.external_version)
        self.assertNotIn("repair", capability.supported_operations)
        self.assertIn("repair", capability.unavailable_reasons)
        self.assertEqual(result.status, "blocked")
        self.assertFalse(result.verified)

    def test_missing_acrobat_exposes_only_unqualified_default_app_handoff(self):
        capability = self.provider().capability()
        self.assertEqual(capability.availability, "default_app_only")
        self.assertFalse(capability.acrobat_detected)
        self.assertIn("open_manual", capability.supported_operations)
        self.assertEqual(capability.external_qualification, "pending")

    def test_acrobat_discovery_error_keeps_inspection_and_reports_degraded_handoff(self):
        api = self.api()
        with patch.object(api.system_shell, "find_acrobat_executable", side_effect=OSError):
            try:
                provider = api.AcrobatManualProvider(platform_name="Windows")
            except OSError:
                self.fail("Acrobat discovery failure must not disable local font inspection")
        capability = provider.capability()
        self.assertEqual(capability.availability, "default_app_only")
        self.assertIn("OSError", capability.unavailable_reasons["acrobat"])
        with tempfile.TemporaryDirectory() as directory:
            result = api.FontWorkflow(provider).recheck(self.make_pdf(directory))
        self.assertTrue(result.verified)

    def test_unsupported_platform_keeps_inspection_but_blocks_handoff(self):
        api = self.api()
        workflow = api.FontWorkflow(self.provider(platform="UnsupportedOS"))
        capability = workflow.capability()
        self.assertEqual(capability.availability, "unsupported")
        self.assertIn("inspect", capability.supported_operations)
        self.assertNotIn("open_manual", capability.supported_operations)
        self.assertEqual(workflow.open_manual("sample.pdf").status, "unavailable")

    def test_manual_open_is_pending_and_preserves_original_bytes(self):
        api = self.api()
        with tempfile.TemporaryDirectory() as directory:
            path = self.make_pdf(directory, text=True)
            before = path.read_bytes()
            workflow = api.FontWorkflow(self.provider())
            # Only the OS process launch is external; the adapter/workflow stay real.
            with patch.object(api.system_shell, "open_pdf_in_acrobat_or_default"):
                result = workflow.open_manual(path)
            self.assertEqual(result.status, "manual_required")
            self.assertTrue(result.opened)
            self.assertFalse(result.verified)
            self.assertEqual(path.read_bytes(), before)

    def test_manual_launcher_error_is_failed_without_exposing_path(self):
        api = self.api()
        workflow = api.FontWorkflow(self.provider())
        with patch.object(api.system_shell, "open_pdf_in_acrobat_or_default",
                          side_effect=OSError("private-path.pdf")):
            result = workflow.open_manual("sample.pdf")
        self.assertEqual(result.status, "failed")
        self.assertFalse(result.opened)
        self.assertNotIn("private-path", result.reason)

    def test_launcher_timeout_is_distinct_from_failed_launch(self):
        api = self.api()
        workflow = api.FontWorkflow(self.provider())
        with patch.object(api.system_shell, "open_pdf_in_acrobat_or_default",
                          side_effect=TimeoutError):
            result = workflow.open_manual("sample.pdf")
        self.assertEqual(result.status, "timeout")
        self.assertFalse(result.verified)

    def test_real_generated_pdf_recheck_reports_unembedded_and_preserves_bytes(self):
        api = self.api()
        with tempfile.TemporaryDirectory() as directory:
            path = self.make_pdf(directory, text=True)
            before = path.read_bytes()
            result = api.FontWorkflow(self.provider()).recheck(path)
            self.assertEqual(result.status, "risks_remain")
            self.assertEqual(result.unembedded_count, 1)
            self.assertEqual(result.substitution_count, 0)
            self.assertEqual(result.unknown_count, 0)
            self.assertFalse(result.verified)
            self.assertEqual(path.read_bytes(), before)

    def test_real_blank_pdf_can_be_verified_without_claiming_repair(self):
        api = self.api()
        with tempfile.TemporaryDirectory() as directory:
            path = self.make_pdf(directory)
            result = api.FontWorkflow(self.provider()).inspect(path)
            self.assertTrue(result.verified)
            self.assertEqual(result.status, "verified")
            self.assertEqual((result.unembedded_count, result.substitution_count,
                              result.unknown_count), (0, 0, 0))

    def test_recheck_reads_saved_bytes_instead_of_reusing_previous_result(self):
        api = self.api()
        with tempfile.TemporaryDirectory() as directory:
            path = self.make_pdf(directory)
            workflow = api.FontWorkflow(self.provider())
            self.assertTrue(workflow.recheck(path).verified)
            with fitz.open(path) as doc:
                doc[0].insert_text((72, 72), "Newly saved unembedded font")
                doc.saveIncr()
            saved_bytes = path.read_bytes()
            result = workflow.recheck(path)
            self.assertFalse(result.verified)
            self.assertEqual(result.unembedded_count, 1)
            self.assertEqual(path.read_bytes(), saved_bytes)

    def test_unreadable_pdf_is_unavailable_and_never_looks_risk_free(self):
        api = self.api()
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "invalid.pdf"
            path.write_bytes(b"not a PDF")
            result = api.FontWorkflow(self.provider()).recheck(path)
            self.assertEqual(result.status, "unavailable")
            self.assertIsNone(result.unembedded_count)
            self.assertIsNone(result.substitution_count)
            self.assertIsNone(result.unknown_count)
            self.assertFalse(result.verified)
            # Closing the real inspector releases the input even on failure.
            moved = path.with_name("released.pdf")
            path.rename(moved)
            self.assertEqual(moved.read_bytes(), b"not a PDF")

    def test_page_font_enumeration_failure_cannot_verify_a_real_pdf(self):
        api = self.api()
        with tempfile.TemporaryDirectory() as directory:
            path = self.make_pdf(directory, text=True)
            before = path.read_bytes()
            # Keep the real open/iteration/collector/close path; only the failing
            # PyMuPDF font-enumeration boundary is substituted.
            with patch.object(fitz.Page, "get_fonts", side_effect=RuntimeError("enumeration failed")):
                result = api.FontWorkflow(self.provider()).recheck(path)
            self.assertFalse(result.verified, "partial font enumeration must never certify zero risk")
            self.assertEqual(result.status, "unavailable")
            self.assertIsNone(result.unembedded_count)
            self.assertIsNone(result.substitution_count)
            self.assertIsNone(result.unknown_count)
            self.assertEqual(path.read_bytes(), before)
            moved = path.with_name("released.pdf")
            path.rename(moved)
            self.assertEqual(moved.read_bytes(), before)

    def fake_provider(self, *, failure=None, result_status="reported_success"):
        """Automatic success/failure exists exclusively in this test utility."""
        api = self.api()

        class TestProvider:
            def capability(self):
                return api.FontCapability(
                    platform="test", provider="test_only", version="1",
                    external_version=None, availability="available",
                    supported_operations=("inspect", "repair", "recheck"),
                    unavailable_reasons={}, acrobat_detected=False,
                    external_qualification="test_only",
                )

            def repair(self, path):
                if failure is not None:
                    raise failure
                return api.FontOperationResult(status=result_status, reason="test attempt")

        return TestProvider()

    def test_provider_false_success_with_real_remaining_risk_is_recheck_failed(self):
        api = self.api()
        with tempfile.TemporaryDirectory() as directory:
            path = self.make_pdf(directory, text=True)
            before = path.read_bytes()
            result = api.FontWorkflow(self.fake_provider()).repair(path)
            self.assertEqual(result.status, "recheck_failed")
            self.assertTrue(result.provider_reported_success)
            self.assertFalse(result.verified)
            self.assertEqual(result.inspection.unembedded_count, 1)
            self.assertEqual(path.read_bytes(), before)

    def test_substitution_and_unknown_risks_independently_prevent_verified_result(self):
        api = self.api()
        for flag in ("substitution_risk", "has_unknown_embedding_status"):
            with self.subTest(flag=flag):
                finding = {
                    "font_name": "sample", "original_names": ["sample"],
                    "normalized_name": "sample", "pages": [1], "embedded": False,
                    "base14": False, "substitution_risk": False,
                    "embedding_status_known": flag != "has_unknown_embedding_status",
                    "has_unembedded": False, "has_unknown_embedding_status": False,
                }
                finding[flag] = True
                # Rare font classifications use the existing public inspector seam.
                raw = {"available": True, "error": "", "font_summary": "",
                       "font_details": "", "font_findings": [finding]}
                with patch.object(api.pdf_inspect, "collect_font_precheck_for_path", return_value=raw):
                    result = api.FontWorkflow(self.fake_provider()).repair("sample.pdf")
                self.assertEqual(result.status, "recheck_failed")
                self.assertFalse(result.verified)

    def test_provider_success_requires_an_available_recheck(self):
        api = self.api()
        result = api.FontWorkflow(self.fake_provider()).repair("nonexistent.pdf")
        self.assertEqual(result.status, "recheck_failed")
        self.assertFalse(result.verified)
        self.assertEqual(result.inspection.status, "unavailable")

    def test_provider_timeout_and_failure_have_distinct_statuses(self):
        api = self.api()
        for failure, status in ((TimeoutError(), "timeout"), (RuntimeError(), "failed")):
            with self.subTest(status=status):
                result = api.FontWorkflow(self.fake_provider(failure=failure)).repair("sample.pdf")
                self.assertEqual(result.status, status)
                self.assertFalse(result.verified)

    def test_provider_failure_is_not_promoted_by_a_risk_free_input(self):
        api = self.api()
        with tempfile.TemporaryDirectory() as directory:
            path = self.make_pdf(directory)
            result = api.FontWorkflow(self.fake_provider(result_status="failed")).repair(path)
            self.assertEqual(result.status, "failed")
            self.assertFalse(result.verified)

    def test_verified_result_requires_provider_success_and_fresh_zero_risk_inspection(self):
        api = self.api()
        with tempfile.TemporaryDirectory() as directory:
            path = self.make_pdf(directory)
            result = api.FontWorkflow(self.fake_provider()).repair(path)
            self.assertEqual(result.status, "verified")
            self.assertTrue(result.verified)
            self.assertTrue(result.inspection.verified)


if __name__ == "__main__":
    unittest.main()
