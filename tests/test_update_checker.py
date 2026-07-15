import unittest
from unittest import mock

from ratools_pdf.services import update_checker


class UpdateCheckerTests(unittest.TestCase):
    def test_parse_version_tag_accepts_optional_v_prefix(self):
        self.assertEqual(update_checker.parse_version_tag("v1.2.3"), (1, 2, 3))
        self.assertEqual(update_checker.parse_version_tag("1.2.3"), (1, 2, 3))

    def test_release_from_github_payload_builds_release_info(self):
        payload = {
            "draft": False,
            "prerelease": False,
            "tag_name": "v1.2.3",
            "name": "Release 1.2.3",
            "body": "Notes",
            "html_url": "https://example.com/release",
            "published_at": "2026-07-04T00:00:00Z",
        }

        release = update_checker.release_from_github_payload(payload)

        self.assertEqual(release.version, (1, 2, 3))
        self.assertEqual(release.version_text, "1.2.3")
        self.assertEqual(release.title, "Release 1.2.3")
        self.assertEqual(release.html_url, "https://example.com/release")

    def test_release_from_github_payload_rejects_prerelease(self):
        payload = {
            "draft": False,
            "prerelease": True,
            "tag_name": "v1.2.3",
        }

        with self.assertRaises(ValueError):
            update_checker.release_from_github_payload(payload)

    def test_is_major_update_uses_minor_version_when_current_major_is_zero(self):
        latest = update_checker.ReleaseInfo(
            version=(0, 7, 0),
            version_text="0.7.0",
            title="Release 0.7.0",
            body="",
            html_url="",
            published_at="",
        )

        self.assertTrue(update_checker.is_major_update((0, 6, 2), latest))

    def test_select_latest_release_picks_highest_semver(self):
        payloads = [
            {"draft": False, "prerelease": False, "tag_name": "v0.7.0", "name": "0.7.0"},
            {"draft": False, "prerelease": False, "tag_name": "v0.7.1", "name": "0.7.1"},
            {"draft": False, "prerelease": False, "tag_name": "v0.6.9", "name": "0.6.9"},
        ]

        release = update_checker.select_latest_release(payloads)

        self.assertEqual(release.version, (0, 7, 1))

    def test_select_latest_release_ignores_prereleases_and_drafts(self):
        payloads = [
            {"draft": False, "prerelease": True, "tag_name": "v0.8.0", "name": "0.8.0"},
            {"draft": True, "prerelease": False, "tag_name": "v0.7.2", "name": "0.7.2"},
            {"draft": False, "prerelease": False, "tag_name": "v0.7.1", "name": "0.7.1"},
        ]

        release = update_checker.select_latest_release(payloads)

        self.assertEqual(release.version, (0, 7, 1))

    def test_select_latest_release_skips_unparsable_tags(self):
        payloads = [
            {"draft": False, "prerelease": False, "tag_name": "nightly", "name": "nightly"},
            {"draft": False, "prerelease": False, "tag_name": "v0.7.1", "name": "0.7.1"},
        ]

        release = update_checker.select_latest_release(payloads)

        self.assertEqual(release.version, (0, 7, 1))

    def test_select_latest_release_raises_when_no_stable_release(self):
        payloads = [
            {"draft": False, "prerelease": True, "tag_name": "v0.8.0", "name": "0.8.0"},
        ]

        with self.assertRaises(ValueError):
            update_checker.select_latest_release(payloads)

    def test_check_for_updates_marks_available_update(self):
        latest = update_checker.ReleaseInfo(
            version=(0, 6, 3),
            version_text="0.6.3",
            title="Release 0.6.3",
            body="",
            html_url="",
            published_at="",
        )

        with mock.patch("ratools_pdf.services.update_checker.fetch_latest_release", return_value=latest):
            result = update_checker.check_for_updates(current_version=(0, 6, 2))

        self.assertTrue(result.ok)
        self.assertTrue(result.has_update)
        self.assertEqual(result.current_version, "0.6.2")
        self.assertEqual(result.latest_release.version, (0, 6, 3))

    def test_check_for_updates_returns_error_result_on_fetch_failure(self):
        with mock.patch("ratools_pdf.services.update_checker.fetch_latest_release", side_effect=OSError("boom")):
            result = update_checker.check_for_updates(current_version=(0, 6, 2))

        self.assertFalse(result.ok)
        self.assertFalse(result.has_update)
        self.assertEqual(result.current_version, "0.6.2")
        self.assertIn("boom", result.error)


if __name__ == "__main__":
    unittest.main()
