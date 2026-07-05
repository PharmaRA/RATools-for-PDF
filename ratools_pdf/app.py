import ctypes
import multiprocessing as mp
import sys

from PySide6.QtCore import Qt, QTimer
from PySide6.QtWidgets import QApplication

from ratools_pdf.config.features import ENABLE_UPDATE_CHECK
from ratools_pdf.controllers.main_controller import MainController
from ratools_pdf.ui.main_window import MainWindow


def detach_console_if_needed():
    if sys.platform != "win32" or not getattr(sys, "frozen", False):
        return

    try:
        kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
        if kernel32.GetConsoleWindow():
            kernel32.FreeConsole()
    except Exception:
        pass


def configure_runtime():
    # PyInstaller 冻结后，multiprocessing 子进程需要先经过 freeze_support，
    # 否则点击处理时会再次拉起整个 GUI 程序。
    mp.freeze_support()
    detach_console_if_needed()
    QApplication.setHighDpiScaleFactorRoundingPolicy(
        Qt.HighDpiScaleFactorRoundingPolicy.PassThrough
    )


def run():
    configure_runtime()

    app = QApplication(sys.argv)
    view = MainWindow()
    controller = MainController(view)

    view.show()
    if ENABLE_UPDATE_CHECK:
        QTimer.singleShot(0, controller.check_updates_on_startup)
    sys.exit(app.exec())
