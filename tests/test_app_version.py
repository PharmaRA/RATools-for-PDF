import unittest

import app_version


class AppVersionTests(unittest.TestCase):
    def test_get_display_version_uses_version_string(self):
        self.assertEqual(
            app_version.get_display_version(),
            f"Version {app_version.APP_VERSION_STR}",
        )


if __name__ == "__main__":
    unittest.main()
