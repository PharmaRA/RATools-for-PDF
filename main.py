import sys
import multiprocessing as mp
from PySide6.QtWidgets import QApplication
from PySide6.QtCore import Qt

# 导入分离出来的 View (界面) 和 Controller (逻辑) 模块
from view import MainWindow
from controller import MainController


def configure_runtime():
    # PyInstaller 冻结后，multiprocessing 子进程需要先经过 freeze_support，
    # 否则点击处理时会再次拉起整个 GUI 程序。
    mp.freeze_support()
    QApplication.setHighDpiScaleFactorRoundingPolicy(Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)

if __name__ == '__main__':
    configure_runtime()

    app = QApplication(sys.argv)

    # 1. 初始化视图 (View)
    view = MainWindow()

    # 2. 初始化逻辑控制器 (Controller)，并将视图注入进去
    controller = MainController(view)

    # 3. 显示主窗口并进入应用循环
    view.show()
    sys.exit(app.exec())
