import os
import tempfile
import unittest
from unittest.mock import patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import QSettings
from PySide6.QtWidgets import QAbstractSpinBox, QApplication

from ratools_pdf.ui import theme
from ratools_pdf.ui.main_window import MainWindow
from ratools_pdf.ui.theme import DARK, active_palette


class ThemeStartupTests(unittest.TestCase):
    def setUp(self):
        self.app = QApplication.instance() or QApplication([])
        self._previous_app_stylesheet = self.app.styleSheet()
        self.app.setStyleSheet("")
        self._windows = []

    def tearDown(self):
        for window in self._windows:
            window.close()
            window.deleteLater()
        self.app.processEvents()
        self.app.setStyleSheet(self._previous_app_stylesheet)

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

        with patch("ratools_pdf.ui.theme.sys.platform", "win32"), patch.object(theme.ctypes, "windll", FakeWindll(), create=True):
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

        with patch("ratools_pdf.ui.theme.sys.platform", "win32"), patch.object(theme.ctypes, "windll", FakeWindll(), create=True):
            self.assertTrue(theme.apply_windows_title_bar_theme(FakeWidget(), DARK))

        values_by_attribute = {call["attribute"]: call["value"] for call in dwm_calls}
        self.assertEqual(1, values_by_attribute[20])
        self.assertEqual(0x00221C19, values_by_attribute[35])
        self.assertEqual(0x00F6F3F1, values_by_attribute[36])
        self.assertEqual(0x003D332E, values_by_attribute[34])

    def test_main_window_reapplies_native_title_bar_when_theme_changes(self):
        with patch("ratools_pdf.ui.main_window.apply_windows_title_bar_theme", create=True) as apply_title_bar:
            window = self._create_window_with_theme("dark")

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

    def test_worker_spinbox_uses_custom_stepper_buttons(self):
        window = self._create_window_with_theme("light")
        dialog = window.settings_dialog

        self.assertEqual(QAbstractSpinBox.NoButtons, dialog.spin_parallel_workers.buttonSymbols())
        self.assertEqual(theme.Qt.UpArrow, dialog.btn_parallel_worker_up.arrowType())
        self.assertEqual(theme.Qt.DownArrow, dialog.btn_parallel_worker_down.arrowType())

    def test_worker_spinbox_styles_custom_stepper_buttons(self):
        qss = theme.build_app_qss(theme.LIGHT)

        self.assertIn("#settingsSpinStepper", qss)
        self.assertIn("#settingsSpinStepBtn", qss)
        self.assertIn('#settingsSpinStepBtn[direction="up"]', qss)
        self.assertIn('#settingsSpinStepBtn[direction="down"]', qss)


if __name__ == "__main__":
    unittest.main()
