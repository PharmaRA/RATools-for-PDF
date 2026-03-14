import sys
from PySide6.QtWidgets import QApplication
from PySide6.QtCore import Qt

# 导入分离出来的 View (界面) 和 Controller (逻辑) 模块
from view import MainWindow
from controller import MainController

if __name__ == '__main__':
    # 支持高分辨率屏幕缩放
    QApplication.setHighDpiScaleFactorRoundingPolicy(Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)

    app = QApplication(sys.argv)

    # 1. 初始化视图 (View)
    view = MainWindow()

    # 2. 初始化逻辑控制器 (Controller)，并将视图注入进去
    controller = MainController(view)

    # 3. 显示主窗口并进入应用循环
    view.show()
    sys.exit(app.exec())