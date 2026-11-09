import ctypes
import multiprocessing as mp
import sys

from PySide6.QtCore import Qt, QTimer
from PySide6.QtWidgets import QApplication

from ratools_pdf.config.features import ENABLE_UPDATE_CHECK


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

    if len(sys.argv) > 1 and sys.argv[1] == '--verify-release':
        if len(sys.argv) != 3:
            raise SystemExit('Usage: --verify-release NEW_OUTPUT_DIRECTORY')
        from ratools_pdf.services.release_check import verify_release
        raise SystemExit(verify_release(sys.argv[2]))

    if len(sys.argv) > 1 and sys.argv[1] == '--verify-image-compression':
        if len(sys.argv) != 3:
            raise SystemExit('Usage: --verify-image-compression NEW_OUTPUT_DIRECTORY')
        from ratools_pdf.services.compression_check import verify_image_compression
        raise SystemExit(verify_image_compression(sys.argv[2]))

    from ratools_pdf.controllers.main_controller import MainController
    from ratools_pdf.ui.main_window import MainWindow
    app = QApplication(sys.argv)
    view = MainWindow()
    controller = MainController(view)

    view.show()
    if ENABLE_UPDATE_CHECK:
        QTimer.singleShot(0, controller.check_updates_on_startup)
    sys.exit(app.exec())
