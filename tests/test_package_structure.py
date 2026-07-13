import unittest
from pathlib import Path


class PackageStructureTests(unittest.TestCase):
    def test_app_run_is_importable_from_package_and_root_entrypoint(self):
        import main
        from ratools_pdf import app

        self.assertIs(main.run, app.run)

    def test_config_modules_are_available_from_package(self):
        from ratools_pdf.config import features, paths, version

        self.assertIsInstance(features.ENABLE_UPDATE_CHECK, bool)
        self.assertTrue(callable(paths.get_app_dir))
        self.assertTrue(callable(paths.get_resource_path))
        self.assertTrue(version.APP_VERSION_STR)

    def test_service_modules_are_available_from_package(self):
        from ratools_pdf.services import update_checker as package_update_checker

        self.assertTrue(callable(package_update_checker.check_for_updates))
        self.assertTrue(callable(package_update_checker.release_from_github_payload))

    def test_ui_helper_modules_are_available_from_package(self):
        from ratools_pdf.ui import log_view_model as package_log_view_model
        from ratools_pdf.ui import theme as package_theme

        self.assertTrue(callable(package_log_view_model.build_log_summary_items))
        self.assertTrue(callable(package_log_view_model.filter_log_summary_items))
        self.assertTrue(callable(package_theme.active_palette))
        self.assertTrue(package_theme.ThemeManager)

    def test_pdf_helper_modules_are_available_from_package(self):
        from ratools_pdf.pdf import font_embedding_providers as package_font_embedding_providers

        self.assertTrue(callable(package_font_embedding_providers.get_font_embedding_provider))

    def test_pdf_processor_is_available_from_package(self):
        from ratools_pdf.pdf.processor import PDFProcessor

        self.assertEqual(PDFProcessor.__name__, "PDFProcessor")

    def test_pdf_helper_modules_are_split_by_responsibility(self):
        from ratools_pdf.pdf import bookmarks_links, hyperlink_styles, page_layout, precheck, qpdf

        self.assertTrue(callable(qpdf.rewrite_with_qpdf))
        self.assertTrue(callable(precheck.build_precheck_report))
        self.assertTrue(callable(bookmarks_links.export_bookmarks))
        self.assertTrue(callable(page_layout.resize_pages_with_padding))
        self.assertTrue(callable(hyperlink_styles.apply_hyperlink_styles))

    def test_main_controller_is_available_from_package(self):
        from ratools_pdf.controllers.main_controller import MainController

        self.assertEqual(MainController.__name__, "MainController")

    def test_ui_classes_are_available_from_package(self):
        from ratools_pdf.ui.dialogs import IODataWizardDialog, LogDialog
        from ratools_pdf.ui.main_window import MainWindow

        self.assertEqual(IODataWizardDialog.__name__, "IODataWizardDialog")
        self.assertEqual(LogDialog.__name__, "LogDialog")
        self.assertEqual(MainWindow.__name__, "MainWindow")

    def test_root_compatibility_shim_files_are_removed(self):
        repo_root = Path(__file__).resolve().parents[1]
        shim_names = [
            "app_features.py",
            "app_paths.py",
            "app_version.py",
            "controller.py",
            "font_embedding_providers.py",
            "log_view_model.py",
            "pdf_processor.py",
            "theme.py",
            "update_checker.py",
            "view.py",
        ]

        remaining_shims = [name for name in shim_names if (repo_root / name).exists()]

        self.assertEqual([], remaining_shims)

    def test_build_scripts_use_package_module_paths(self):
        repo_root = Path(__file__).resolve().parents[1]
        pyinstaller_script = (repo_root / "build_pyinstaller.bat").read_text(encoding="utf-8")
        nuitka_script = (repo_root / "build_nuitka.bat").read_text(encoding="utf-8")
        release_workflow = (repo_root / ".github" / "workflows" / "build.yml").read_text(encoding="utf-8")
        package_version_import = "from ratools_pdf.config.version import APP_VERSION_STR"

        self.assertIn(package_version_import, pyinstaller_script)
        self.assertIn(package_version_import, nuitka_script)
        self.assertIn(package_version_import, release_workflow)
        self.assertIn("--include-module=ratools_pdf.config.paths", nuitka_script)
        self.assertIn("--exclude-module ratools_pdf.services.update_checker", pyinstaller_script)
        self.assertNotIn("from app_version import", pyinstaller_script)
        self.assertNotIn("from app_version import", nuitka_script)
        self.assertNotIn("from app_version import", release_workflow)
        self.assertNotIn("--include-module=app_paths", nuitka_script)
        self.assertNotIn("--exclude-module update_checker", pyinstaller_script)


if __name__ == "__main__":
    unittest.main()
