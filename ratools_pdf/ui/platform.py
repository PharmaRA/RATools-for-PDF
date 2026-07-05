import sys


def is_win11():
    """准确判断是否为 Windows 11 (基于内部 build 版本号)"""
    if sys.platform != "win32":
        return False
    # Win11 的版本号是从 22000 开始的
    return sys.getwindowsversion().build >= 22000


def should_use_manual_dialog_shadow():
    """Win11 自带窗口阴影明显，关闭手工阴影以避免双层叠加。"""
    return not is_win11()


# ================== 自定义无边框拖拽对话框基类 ==================
