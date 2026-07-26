"""Windows DWM / Win32 集成工具。

此前 DwmSetWindowAttribute 相关 ctypes 逻辑散落三处（dialogs 的去边框、
theme 的标题栏配色、platform 的版本判定），DWM 属性 ID 全是裸数字。
这里集中定义常量与调用封装；非 Windows 平台所有函数静默返回 False。
"""

import ctypes
import sys

# --- DWMWINDOWATTRIBUTE（Windows SDK dwmapi.h）---
DWMWA_USE_IMMERSIVE_DARK_MODE_LEGACY = 19  # Win10 1809~1903 的旧属性号
DWMWA_USE_IMMERSIVE_DARK_MODE = 20
DWMWA_WINDOW_CORNER_PREFERENCE = 33
DWMWA_BORDER_COLOR = 34
DWMWA_CAPTION_COLOR = 35
DWMWA_TEXT_COLOR = 36

# DWM_WINDOW_CORNER_PREFERENCE 取值
DWMWCP_DONOTROUND = 1

# DWMWA_BORDER_COLOR 特殊值：完全不画边框
DWMWA_COLOR_NONE = 0xFFFFFFFE


def hex_to_colorref(hex_value: str) -> int:
    """把 #RRGGBB 转为 Windows COLORREF (0x00BBGGRR)。"""
    normalized = hex_value.lstrip("#")
    red = int(normalized[0:2], 16)
    green = int(normalized[2:4], 16)
    blue = int(normalized[4:6], 16)
    return (blue << 16) | (green << 8) | red


def set_dwm_int_attribute(hwnd: int, attribute: int, value: int) -> bool:
    """DwmSetWindowAttribute 的 int 版封装；失败/非 Windows 返回 False。"""
    if sys.platform != "win32" or not hwnd:
        return False
    data = ctypes.c_int(value)
    try:
        result = ctypes.windll.dwmapi.DwmSetWindowAttribute(
            ctypes.c_void_p(int(hwnd)),
            attribute,
            ctypes.byref(data),
            ctypes.sizeof(data),
        )
    except Exception:
        return False
    return result == 0


def _set_dwm_uint_attribute(hwnd: int, attribute: int, value: int) -> bool:
    if sys.platform != "win32" or not hwnd:
        return False
    data = ctypes.c_uint(value)
    try:
        result = ctypes.windll.dwmapi.DwmSetWindowAttribute(
            ctypes.c_void_p(int(hwnd)),
            attribute,
            ctypes.byref(data),
            ctypes.sizeof(data),
        )
    except Exception:
        return False
    return result == 0


def refresh_window_frame(hwnd: int):
    """强制窗口重算非客户区，让 DWM 属性变更立即生效。"""
    if sys.platform != "win32" or not hwnd:
        return
    try:
        # SWP_NOSIZE | SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED
        flags = 0x0001 | 0x0002 | 0x0004 | 0x0010 | 0x0020
        ctypes.windll.user32.SetWindowPos(
            ctypes.c_void_p(int(hwnd)), None, 0, 0, 0, 0, flags,
        )
        # RDW_INVALIDATE | RDW_UPDATENOW | RDW_FRAME
        ctypes.windll.user32.RedrawWindow(ctypes.c_void_p(int(hwnd)), None, None, 0x0501)
    except Exception:
        pass


def remove_win11_window_decorations(hwnd: int) -> bool:
    """去掉 Win11 强制附加的系统圆角与 1px 边框（无边框自绘窗口用）。"""
    if sys.platform != "win32" or not hwnd:
        return False
    ok = set_dwm_int_attribute(hwnd, DWMWA_WINDOW_CORNER_PREFERENCE, DWMWCP_DONOTROUND)
    ok = _set_dwm_uint_attribute(hwnd, DWMWA_BORDER_COLOR, DWMWA_COLOR_NONE) and ok
    return ok


def set_title_bar_dark_mode(
    hwnd: int,
    enabled: bool,
    caption_color: str | None = None,
    text_color: str | None = None,
    border_color: str | None = None,
) -> bool:
    """在 Windows 原生标题栏上应用深色模式与可选配色。"""
    if sys.platform != "win32" or not hwnd:
        return False

    mode_value = 1 if enabled else 0
    ok = set_dwm_int_attribute(hwnd, DWMWA_USE_IMMERSIVE_DARK_MODE, mode_value)
    if not ok:
        ok = set_dwm_int_attribute(hwnd, DWMWA_USE_IMMERSIVE_DARK_MODE_LEGACY, mode_value)

    # Windows 11 支持显式标题栏配色；Windows 10 会返回无效参数，忽略即可。
    for attribute, color in (
        (DWMWA_CAPTION_COLOR, caption_color),
        (DWMWA_TEXT_COLOR, text_color),
        (DWMWA_BORDER_COLOR, border_color),
    ):
        if color:
            set_dwm_int_attribute(hwnd, attribute, hex_to_colorref(color))

    refresh_window_frame(hwnd)
    return ok
