import gc
import os
import tempfile
import unittest
from unittest.mock import patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import QCoreApplication, QEvent, QSettings
from PySide6.QtWidgets import QAbstractSpinBox, QApplication

from ratools_pdf.ui import theme
from ratools_pdf.ui import win32
from ratools_pdf.ui.main_window import MainWindow
from ratools_pdf.ui.theme import DARK, active_palette


class ThemeStartupTests(unittest.TestCase):
    def setUp(self):
        self.app = QApplication.instance() or QApplication([])
        QCoreApplication.sendPostedEvents(None, QEvent.DeferredDelete)
        gc.collect()
        self._previous_app_stylesheet = self.app.styleSheet()
        self._previous_app_font = self.app.font()
        self.app.setStyleSheet("")
        self._windows = []

    def tearDown(self):
        for window in self._windows:
            window.close()
            window.deleteLater()
        self._windows.clear()
        QCoreApplication.sendPostedEvents(None, QEvent.DeferredDelete)
        self.app.processEvents()
        gc.collect()
        self.app.setStyleSheet(self._previous_app_stylesheet)
        self.app.setFont(self._previous_app_font)

    def _create_window_with_theme(self, theme_mode):
        with tempfile.TemporaryDirectory() as tmp_dir:
            settings = QSettings(os.path.join(tmp_dir, "settings.ini"), QSettings.IniFormat)
            settings.setValue("Settings/ThemeMode", theme_mode)
            settings.sync()

            with patch("ratools_pdf.ui.main_window.get_app_dir", return_value=tmp_dir):
                window = MainWindow()
                self._windows.append(window)
                return window

    def test_saved_dark_theme_uses_application_stylesheet_without_local_light_override(self):
        window = self._create_window_with_theme("dark")

        self.assertEqual(window.theme_manager.mode, "dark")
        self.assertEqual(active_palette().name, "dark")
        self.assertIn("#0F1115", self.app.styleSheet())
        self.assertEqual("", window.styleSheet())

    def test_windows_title_bar_helper_calls_dwm_dark_attribute(self):
        setter = getattr(theme, "set_windows_title_bar_dark_mode", None)
        self.assertTrue(callable(setter))
        if not callable(setter):
            return

        dwm_calls = []
        frame_calls = []
        redraw_calls = []

        class FakeDwmApi:
            @staticmethod
            def DwmSetWindowAttribute(hwnd, attribute, enabled_ptr, size):
                dwm_calls.append({
                    "hwnd": hwnd,
                    "attribute": attribute,
                    "enabled": enabled_ptr._obj.value,
                    "size": size,
                })
                return 0

        class FakeUser32:
            @staticmethod
            def SetWindowPos(hwnd, hwnd_after, x, y, cx, cy, flags):
                frame_calls.append({
                    "hwnd": hwnd,
                    "flags": flags,
                })
                return 1

            @staticmethod
            def RedrawWindow(hwnd, rect, region, flags):
                redraw_calls.append({
                    "hwnd": hwnd,
                    "flags": flags,
                })
                return 1

        class FakeWindll:
            dwmapi = FakeDwmApi()
            user32 = FakeUser32()

        with patch("ratools_pdf.ui.win32.sys.platform", "win32"), patch.object(win32.ctypes, "windll", FakeWindll(), create=True):
            self.assertTrue(setter(12345, True))

        self.assertEqual(1, len(dwm_calls))
        self.assertEqual(20, dwm_calls[0]["attribute"])
        self.assertEqual(1, dwm_calls[0]["enabled"])
        self.assertEqual(1, len(frame_calls))
        self.assertEqual(0x0037, frame_calls[0]["flags"])
        self.assertEqual(1, len(redraw_calls))
        self.assertEqual(0x0501, redraw_calls[0]["flags"])

    def test_windows_title_bar_theme_applies_palette_caption_colors_when_supported(self):
        dwm_calls = []

        class FakeWidget:
            def winId(self):
                return 12345

        class FakeDwmApi:
            @staticmethod
            def DwmSetWindowAttribute(hwnd, attribute, value_ptr, size):
                dwm_calls.append({
                    "attribute": attribute,
                    "value": value_ptr._obj.value,
                })
                return 0

        class FakeUser32:
            @staticmethod
            def SetWindowPos(hwnd, hwnd_after, x, y, cx, cy, flags):
                return 1

            @staticmethod
            def RedrawWindow(hwnd, rect, region, flags):
                return 1

        class FakeWindll:
            dwmapi = FakeDwmApi()
            user32 = FakeUser32()

        with patch("ratools_pdf.ui.win32.sys.platform", "win32"), patch.object(win32.ctypes, "windll", FakeWindll(), create=True):
            self.assertTrue(theme.apply_windows_title_bar_theme(FakeWidget(), DARK))

        values_by_attribute = {call["attribute"]: call["value"] for call in dwm_calls}
        self.assertEqual(1, values_by_attribute[20])
        self.assertEqual(0x00221C19, values_by_attribute[35])
        self.assertEqual(0x00F6F3F1, values_by_attribute[36])
        self.assertEqual(0x003D332E, values_by_attribute[34])

    def test_main_window_reapplies_native_title_bar_when_theme_changes(self):
        with patch("ratools_pdf.ui.main_window.apply_windows_title_bar_theme", create=True) as apply_title_bar:
            window = self._create_window_with_theme("dark")

            apply_title_bar.assert_not_called()
            window.show()
            self.app.processEvents()

            self.assertTrue(any(
                call.args[0] is window and call.args[1].name == "dark"
                for call in apply_title_bar.call_args_list
            ))

            apply_title_bar.reset_mock()
            window.on_theme_mode_changed("light")

            self.assertTrue(any(
                call.args[0] is window and call.args[1].name == "light"
                for call in apply_title_bar.call_args_list
            ))

    def test_theme_change_schedules_native_title_bar_refresh_after_click_event(self):
        with patch("ratools_pdf.ui.main_window.apply_windows_title_bar_theme", create=True), \
                patch("ratools_pdf.ui.main_window.QTimer.singleShot", create=True) as single_shot:
            window = self._create_window_with_theme("dark")

            single_shot.reset_mock()
            window.on_theme_mode_changed("light")

            self.assertTrue(any(
                call.args[0] == 0
                and getattr(call.args[1], "__self__", None) is window
                and getattr(call.args[1], "__name__", "") == "apply_native_title_bar_theme"
                for call in single_shot.call_args_list
            ))

    def test_switching_from_forced_dark_to_system_resolves_current_system_scheme(self):
        class FakeStyleHints:
            def __init__(self):
                self.scheme = theme.Qt.ColorScheme.Light

            def colorScheme(self):
                return self.scheme

            def setColorScheme(self, scheme):
                self.scheme = theme.Qt.ColorScheme.Light if scheme == theme.Qt.ColorScheme.Unknown else scheme

        class FakeApp:
            def __init__(self):
                self.hints = FakeStyleHints()
                self.stylesheet = ""

            def styleHints(self):
                return self.hints

            def setStyleSheet(self, qss):
                self.stylesheet = qss

        manager = theme.ThemeManager(self.app, "dark")
        manager._app = FakeApp()

        manager.apply()
        self.assertEqual("dark", manager.current_palette().name)

        manager.set_mode("system")

        self.assertEqual("light", manager.current_palette().name)
        self.assertIn("#F7F8FA", manager._app.stylesheet)

    def test_theme_change_recolors_tree_status_column(self):
        from PySide6.QtWidgets import QTreeWidgetItem

        from ratools_pdf.controllers.main_controller import MainController
        from ratools_pdf.ui.theme import LIGHT

        window = self._create_window_with_theme("light")
        controller = MainController(window)

        done_node = QTreeWidgetItem(window.tree)
        done_node.setText(2, "处理完成")
        waiting_node = QTreeWidgetItem(window.tree)
        waiting_node.setText(2, "等待处理")
        controller.file_nodes = {"a.pdf": done_node, "b.pdf": waiting_node}

        window.on_theme_mode_changed("dark")

        self.assertEqual(DARK.success, done_node.foreground(2).color().name().upper())
        self.assertEqual(DARK.text_muted, waiting_node.foreground(2).color().name().upper())

        window.on_theme_mode_changed("light")

        self.assertEqual(LIGHT.success, done_node.foreground(2).color().name().upper())
        self.assertEqual(LIGHT.text_muted, waiting_node.foreground(2).color().name().upper())

    def test_worker_spinbox_uses_custom_stepper_buttons(self):
        window = self._create_window_with_theme("light")
        dialog = window.settings_dialog
        dialog.show()
        self.app.processEvents()
        try:
            self.assertEqual(QAbstractSpinBox.NoButtons, dialog.spin_parallel_workers.buttonSymbols())
            self.assertEqual(theme.Qt.UpArrow, dialog.btn_parallel_worker_up.arrowType())
            self.assertEqual(theme.Qt.DownArrow, dialog.btn_parallel_worker_down.arrowType())
            self.assertGreaterEqual(dialog.spin_parallel_workers.lineEdit().width(), 40)
            self.assertEqual(QAbstractSpinBox.NoButtons, dialog.spin_file_timeout.buttonSymbols())
            self.assertGreaterEqual(dialog.spin_file_timeout.lineEdit().width(), 40)
        finally:
            dialog.close()

    def test_worker_spinbox_styles_custom_stepper_buttons(self):
        qss = theme.build_app_qss(theme.LIGHT)

        self.assertIn("#settingsSpinStepper", qss)
        self.assertIn("#settingsSpinStepBtn", qss)
        self.assertIn('#settingsSpinStepBtn[direction="up"]', qss)
        self.assertIn('#settingsSpinStepBtn[direction="down"]', qss)

    def test_fluent_qss_includes_modern_controls_and_indicators(self):
        qss = theme.build_app_qss(theme.LIGHT)

        for selector in (
            "#ruleCard",
            "#thinProgressBar",
            "#statusBadge",
            "QRadioButton::indicator",
            "QTreeWidget::indicator",
            "QComboBox::down-arrow",
            "#queueCard",
            "#modeSegContainer",
            "#segmentedContainer",
            "#watermarkPanelCard",
            "#watermarkFilterBar",
            "#watermarkPreviewScroll",
            "#watermarkConfirmCard",
        ):
            self.assertIn(selector, qss)

        dark_qss = theme.build_app_qss(theme.DARK)
        for selector in (
            "QTreeWidget::indicator",
            "#watermarkPanelCard",
            "#watermarkFilterBar",
            "#watermarkPreviewScroll",
            "#watermarkConfirmCard",
        ):
            self.assertIn(selector, dark_qss)

        self.assertTrue(theme.LIGHT.sidebar.startswith("#"))
        self.assertTrue(theme.DARK.sidebar.startswith("#"))
        self.assertTrue(theme.LIGHT.accent_bar.startswith("#"))
        self.assertTrue(theme.DARK.accent_bar.startswith("#"))

    def test_typography_uses_microsoft_yahei_ui_and_avoids_overbold_qss(self):
        typo = theme.resolve_typography(
            "standard",
            installed={"Segoe UI", "Microsoft YaHei UI", "Noto Sans SC", "Cascadia Mono"},
        )
        self.assertEqual("Microsoft YaHei UI", typo.primary_family)
        self.assertEqual(13, typo.fs_base)
        self.assertIn('"Microsoft YaHei UI"', typo.ui_font_css)
        self.assertIn('"Cascadia Mono"', typo.mono_font_css)

        typo_compact = theme.resolve_typography("compact", installed={"Microsoft YaHei UI"})
        self.assertEqual(12, typo_compact.fs_base)
        self.assertEqual(11, typo_compact.fs_sm)

        typo_large = theme.resolve_typography("comfortable", installed={"Microsoft YaHei UI"})
        self.assertEqual(14, typo_large.fs_base)
        self.assertEqual(13, typo_large.fs_sm)

        qss = theme.build_app_qss(theme.LIGHT, typo)
        for selector in (
            "#moduleSearchInput",
            "#searchActionCard",
            "#searchEmptyCard",
        ):
            self.assertIn(selector, qss)
        self.assertNotIn("font-weight: 600", qss)
        self.assertIn("font-weight: 500", qss)

    def test_font_scale_change_applies_and_persists_across_restart(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            with patch("ratools_pdf.ui.main_window.get_app_dir", return_value=tmp_dir):
                window = MainWindow()
                self._windows.append(window)
                window.on_font_scale_changed("comfortable")
                self.app.processEvents()

                self.assertEqual("comfortable", window.theme_manager.font_scale)
                self.assertEqual(14, window.theme_manager.current_typography().fs_base)
                self.assertIn("font-size: 14px", self.app.styleSheet())
                self.assertEqual("Microsoft YaHei UI", self.app.font().families()[0])
                self.assertEqual("comfortable", window.settings_dialog.get_font_scale())
                window.persist_all_settings()

            with patch("ratools_pdf.ui.main_window.get_app_dir", return_value=tmp_dir):
                reloaded = MainWindow()
                self._windows.append(reloaded)
                self.assertEqual("comfortable", reloaded.theme_manager.font_scale)
                self.assertEqual(14, reloaded.theme_manager.current_typography().fs_base)
                self.assertEqual("comfortable", reloaded.settings_dialog.get_font_scale())


if __name__ == "__main__":
    unittest.main()
