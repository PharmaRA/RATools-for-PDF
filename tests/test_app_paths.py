import os
import unittest
from unittest import mock

from ratools_pdf.config import paths as app_paths


class AppPathsTests(unittest.TestCase):
    def test_get_app_dir_uses_module_directory_when_not_frozen(self):
        expected = str(app_paths.Path(app_paths.__file__).resolve().parents[2])
        with mock.patch.object(app_paths.sys, "frozen", False, create=True):
            self.assertEqual(app_paths.get_app_dir(), expected)

    def test_get_app_dir_uses_executable_parent_when_frozen(self):
        with mock.patch.object(app_paths.sys, "frozen", True, create=True), \
             mock.patch.object(app_paths.sys, "executable", r"C:\\Apps\\RATools\\RATools.exe", create=True):
            expected = str(app_paths.Path(r"C:\\Apps\\RATools\\RATools.exe").resolve().parent)
            self.assertEqual(app_paths.get_app_dir(), expected)

    def test_get_resource_dir_uses_meipass_when_available(self):
        with mock.patch.object(app_paths.sys, "frozen", True, create=True), \
             mock.patch.object(app_paths.sys, "_MEIPASS", r"C:\\Bundle", create=True):
            expected = str(app_paths.Path(r"C:\\Bundle").resolve())
            self.assertEqual(app_paths.get_resource_dir(), expected)

    def test_get_resource_dir_falls_back_to_app_dir_when_frozen_without_meipass(self):
        with mock.patch.object(app_paths.sys, "frozen", True, create=True), \
             mock.patch.object(app_paths.sys, "_MEIPASS", None, create=True), \
             mock.patch.object(app_paths.sys, "executable", r"C:\\Apps\\RATools\\RATools.exe", create=True):
            expected = str(app_paths.Path(r"C:\\Apps\\RATools\\RATools.exe").resolve().parent)
            self.assertEqual(app_paths.get_resource_dir(), expected)

    def test_get_resource_path_joins_requested_parts(self):
        with mock.patch("ratools_pdf.config.paths.get_resource_dir", return_value=r"C:\\Bundle"):
            self.assertEqual(
                app_paths.get_resource_path("plugins", "qpdf", "qpdf.exe"),
                os.path.join(r"C:\\Bundle", "plugins", "qpdf", "qpdf.exe"),
            )


if __name__ == "__main__":
    unittest.main()
