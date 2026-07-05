import unittest


class PackageStructureTests(unittest.TestCase):
    def test_app_run_is_importable_from_package_and_root_entrypoint(self):
        import main
        from ratools_pdf import app

        self.assertIs(main.run, app.run)

    def test_config_modules_are_available_from_package_and_root(self):
        import app_features
        import app_paths
        import app_version
        from ratools_pdf.config import features, paths, version

        self.assertEqual(app_features.ENABLE_UPDATE_CHECK, features.ENABLE_UPDATE_CHECK)
        self.assertIs(app_paths.get_app_dir, paths.get_app_dir)
        self.assertIs(app_paths.get_resource_path, paths.get_resource_path)
        self.assertEqual(app_version.APP_VERSION_STR, version.APP_VERSION_STR)

    def test_service_modules_are_available_from_package_and_root(self):
        import update_checker
        from ratools_pdf.services import update_checker as package_update_checker

        self.assertIs(update_checker.check_for_updates, package_update_checker.check_for_updates)
        self.assertIs(update_checker.release_from_github_payload, package_update_checker.release_from_github_payload)

    def test_ui_helper_modules_are_available_from_package_and_root(self):
        import log_view_model
        import theme
        from ratools_pdf.ui import log_view_model as package_log_view_model
        from ratools_pdf.ui import theme as package_theme

        self.assertIs(log_view_model.build_log_summary_items, package_log_view_model.build_log_summary_items)
        self.assertIs(log_view_model.filter_log_summary_items, package_log_view_model.filter_log_summary_items)
        self.assertIs(theme.ThemeManager, package_theme.ThemeManager)
        self.assertIs(theme.active_palette, package_theme.active_palette)

    def test_pdf_helper_modules_are_available_from_package_and_root(self):
        import font_embedding_providers
        from ratools_pdf.pdf import font_embedding_providers as package_font_embedding_providers

        self.assertIs(
            font_embedding_providers.get_font_embedding_provider,
            package_font_embedding_providers.get_font_embedding_provider,
        )

    def test_pdf_processor_is_available_from_package_and_root(self):
        import pdf_processor
        from ratools_pdf.pdf.processor import PDFProcessor

        self.assertIs(pdf_processor.PDFProcessor, PDFProcessor)

    def test_pdf_helper_modules_are_split_by_responsibility(self):
        from ratools_pdf.pdf import bookmarks_links, hyperlink_styles, page_layout, precheck, qpdf

        self.assertTrue(callable(qpdf.rewrite_with_qpdf))
        self.assertTrue(callable(precheck.build_precheck_report))
        self.assertTrue(callable(bookmarks_links.export_bookmarks))
        self.assertTrue(callable(page_layout.resize_pages_with_padding))
        self.assertTrue(callable(hyperlink_styles.apply_hyperlink_styles))

    def test_main_controller_is_available_from_package_and_root(self):
        import controller
        from ratools_pdf.controllers.main_controller import MainController

        self.assertIs(controller.MainController, MainController)

    def test_ui_classes_are_available_from_package_and_root(self):
        import view
        from ratools_pdf.ui.dialogs import IODataWizardDialog, LogDialog
        from ratools_pdf.ui.main_window import MainWindow

        self.assertIs(view.IODataWizardDialog, IODataWizardDialog)
        self.assertIs(view.LogDialog, LogDialog)
        self.assertIs(view.MainWindow, MainWindow)


if __name__ == "__main__":
    unittest.main()
