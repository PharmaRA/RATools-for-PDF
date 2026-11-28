"""Startup placement keeps the native frame inside the screen work area."""

import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from PySide6.QtCore import QCoreApplication, QEvent, QPoint, QRect, Qt
from PySide6.QtWidgets import QApplication, QScrollArea
from shiboken6 import isValid

from ratools_pdf.ui.main_window import MainWindow


class WindowPlacementTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.window = MainWindow(settings_path=str(Path(self.temp.name) / "settings.ini"))

    def tearDown(self):
        if isValid(self.window):
            self.window.close()
            self.window.deleteLater()
            QCoreApplication.sendPostedEvents(self.window, QEvent.DeferredDelete)
        self.temp.cleanup()

    def show_in_work_area(self, available):
        screen = self.app.primaryScreen()
        with patch.object(QApplication, "screenAt", return_value=screen), \
                patch.object(screen, "availableGeometry", return_value=available), \
                patch.object(self.window, "move", wraps=self.window.move) as move:
            self.window.show()
            self.app.processEvents()
            self.app.processEvents()
            # Keep the simulated screen active while Qt resolves frame margins.
            self.assert_frame_inside(available)
            # Offscreen has no real negative-coordinate monitor and can shift
            # the reported frame by its synthetic border. Check the requested
            # center separately from the actual full-frame containment above.
            requested = QRect(self.window.frameGeometry())
            requested.moveTopLeft(QPoint(*move.call_args.args))
            self.assertLessEqual(abs(requested.center().x() - available.center().x()), 1)
            self.assertLessEqual(abs(requested.center().y() - available.center().y()), 1)

    def assert_frame_inside(self, available):
        frame = self.window.frameGeometry()
        self.assertTrue(available.contains(frame), f"frame={frame}, available={available}")
        flags = self.window.windowFlags()
        self.assertTrue(flags & Qt.WindowTitleHint)
        self.assertTrue(flags & Qt.WindowMinimizeButtonHint)
        self.assertTrue(flags & Qt.WindowMaximizeButtonHint)
        self.assertTrue(flags & Qt.WindowCloseButtonHint)
        self.assertFalse(flags & Qt.FramelessWindowHint)

    def test_startup_recovers_negative_position_and_keeps_native_title_bar(self):
        available = QRect(0, 0, 1920, 1040)
        self.window.move(-480, -320)
        self.show_in_work_area(available)

    def test_startup_reserves_top_taskbar_and_native_frame(self):
        available = QRect(0, 40, 1366, 688)
        self.show_in_work_area(available)

    def test_startup_fits_125_percent_logical_work_area(self):
        available = QRect(0, 0, 1229, 819)
        self.show_in_work_area(available)

    def test_startup_fits_150_percent_logical_work_area(self):
        available = QRect(0, 0, 1024, 640)
        self.show_in_work_area(available)

    def test_startup_centers_on_secondary_screen_with_negative_coordinates(self):
        available = QRect(-1600, -180, 1600, 860)
        self.show_in_work_area(available)

    def test_small_screen_keeps_full_content_accessible_by_scrolling(self):
        available = QRect(0, 0, 768, 480)
        self.show_in_work_area(available)
        scroll = self.window.centralWidget()
        self.assertIsInstance(scroll, QScrollArea)
        self.assertGreater(scroll.horizontalScrollBar().maximum(), 0)
        self.assertGreater(scroll.verticalScrollBar().maximum(), 0)
        self.assertTrue(scroll.widget().isAncestorOf(self.window.btn_start))

    def test_startup_placement_runs_deferred_correction_for_live_window(self):
        with patch.object(self.window, "_finish_startup_placement",
                          wraps=self.window._finish_startup_placement) as finish:
            self.window.show()
            self.assertEqual(finish.call_count, 1)
            self.app.processEvents()
            self.assertEqual(finish.call_count, 2)

    def test_destroyed_window_cancels_pending_startup_placement(self):
        with patch.object(self.window, "_finish_startup_placement",
                          wraps=self.window._finish_startup_placement) as finish, \
                patch("sys.excepthook") as exception_hook:
            self.window.show()
            self.assertEqual(finish.call_count, 1)
            self.window.close()
            self.window.deleteLater()
            QCoreApplication.sendPostedEvents(self.window, QEvent.DeferredDelete)
            self.assertFalse(isValid(self.window))
            self.app.processEvents()
            self.app.processEvents()
            exception_hook.assert_not_called()
            self.assertEqual(finish.call_count, 1)

    def test_reshow_and_theme_change_preserve_user_position(self):
        self.show_in_work_area(QRect(0, 0, 1920, 1040))
        self.window.move(QPoint(80, 90))
        self.window.resize(1000, 700)
        self.app.processEvents()
        before = self.window.frameGeometry()
        self.window.hide()
        self.window.show()
        self.window.on_theme_mode_changed("dark")
        self.app.processEvents()
        self.assertEqual(before, self.window.frameGeometry())


if __name__ == "__main__":
    unittest.main()
