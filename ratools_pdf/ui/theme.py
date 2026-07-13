"""集中式主题系统。

本模块是整个应用视觉层的唯一颜色来源：

* ``Palette``   —— 语义化设计令牌 (token) 的数据结构。
* ``LIGHT`` / ``DARK`` —— 两套配色 (现代精致亮色 + 匹配暗色)。
* ``build_app_qss``   —— 依据当前 Palette 生成整份应用级 QSS。
* ``ThemeManager``    —— 解析「跟随系统 / 手动亮色 / 手动暗色」，
  在 :class:`~PySide6.QtWidgets.QApplication` 上应用样式，并在系统主题
  变化时实时刷新。

设计原则：所有颜色都以令牌形式定义，``view.py`` 不再硬编码任何十六进制色值；
组件通过 objectName / 动态属性挂接到中央 QSS，因此切换主题只需要重新调用
``app.setStyleSheet(...)`` 即可级联到所有已打开与将要打开的窗口。
"""

from __future__ import annotations

import ctypes
import sys
from dataclasses import dataclass, asdict
from string import Template

from PySide6.QtCore import QObject, Qt, Signal
from PySide6.QtGui import QColor
from PySide6.QtWidgets import QApplication


# ============================================================================
# 设计令牌
# ============================================================================
@dataclass(frozen=True)
class Palette:
    """一套完整的语义化配色令牌。"""

    name: str          # "light" | "dark"
    is_dark: bool

    # --- 基础表面 ---
    window: str        # 应用主背景
    surface: str       # 卡片 / 面板 (旧: white)
    surface_alt: str   # 次级填充 (旧: #F8FAFC)
    surface_alt2: str  # 交替行背景 (旧: #FBFDFF)
    surface_hover: str # 悬停填充

    # --- 边框 ---
    border: str        # 默认边框
    border_strong: str # 强调边框 / 输入框
    border_subtle: str # 行分隔线

    # --- 文字 ---
    text: str          # 主标题
    text_body: str     # 正文
    text_muted: str    # 次要说明
    text_faint: str    # 最弱 (占位 / 分区标题)
    text_on_primary: str

    # --- 主色 (靛蓝) ---
    primary: str
    primary_hover: str
    primary_pressed: str
    primary_soft: str        # 选中态浅底
    primary_soft_border: str
    primary_text: str        # 浅底之上的文字
    primary_disabled: str

    # --- 选区 / 焦点 ---
    selection_bg: str
    selection_text: str

    # --- 危险 ---
    danger: str
    danger_hover: str
    danger_pressed: str
    danger_text: str
    danger_soft: str
    danger_soft_border: str

    # --- 成功 ---
    success: str
    success_text: str
    success_soft: str

    # --- 警告 ---
    warning: str
    warning_text: str
    warning_soft: str

    # --- 信息 / 预检 ---
    info: str
    info_text: str
    info_soft: str

    # --- 滚动条 ---
    scrollbar_handle: str
    scrollbar_handle_hover: str
    scrollbar_track: str

    # --- 拖拽热区 ---
    dropzone_border: str
    dropzone_bg: str

    # --- 阴影 (rgba 字符串, 供 QColor 使用) ---
    shadow_rgba: tuple

    def as_dict(self) -> dict:
        return asdict(self)

    def qcolor(self, hex_value: str) -> QColor:
        return QColor(hex_value)


# ---------------------------------------------------------------------------
# 现代精致亮色 —— 靛蓝 #4F46E5
# ---------------------------------------------------------------------------
LIGHT = Palette(
    name="light",
    is_dark=False,

    window="#F7F8FA",
    surface="#FFFFFF",
    surface_alt="#F5F6F9",
    surface_alt2="#FBFCFE",
    surface_hover="#F1F3F7",

    border="#E6E8EC",
    border_strong="#D3D7DE",
    border_subtle="#EEF1F4",

    text="#111827",
    text_body="#374151",
    text_muted="#6B7280",
    text_faint="#9AA2AF",
    text_on_primary="#FFFFFF",

    primary="#4F46E5",
    primary_hover="#4338CA",
    primary_pressed="#3730A3",
    primary_soft="#EEF2FF",
    primary_soft_border="#C7D2FE",
    primary_text="#4338CA",
    primary_disabled="#C7D2FE",

    selection_bg="#E0E7FF",
    selection_text="#3730A3",

    danger="#EF4444",
    danger_hover="#DC2626",
    danger_pressed="#B91C1C",
    danger_text="#B42318",
    danger_soft="#FEF2F2",
    danger_soft_border="#FECACA",

    success="#10B981",
    success_text="#047857",
    success_soft="#ECFDF5",

    warning="#F59E0B",
    warning_text="#92400E",
    warning_soft="#FFF7ED",

    info="#3B82F6",
    info_text="#1D4ED8",
    info_soft="#EFF6FF",

    scrollbar_handle="#CBD5E1",
    scrollbar_handle_hover="#94A3B8",
    scrollbar_track="#EAF0F6",

    dropzone_border="#C2CBD9",
    dropzone_bg="#F8FAFF",

    shadow_rgba=(15, 23, 42, 38),
)


# ---------------------------------------------------------------------------
# 匹配暗色
# ---------------------------------------------------------------------------
DARK = Palette(
    name="dark",
    is_dark=True,

    window="#0F1115",
    surface="#191C22",
    surface_alt="#21252E",
    surface_alt2="#1D212A",
    surface_hover="#262B36",

    border="#2E333D",
    border_strong="#3A404C",
    border_subtle="#252A33",

    text="#F1F3F6",
    text_body="#D4D8DF",
    text_muted="#9AA2AF",
    text_faint="#6B7280",
    text_on_primary="#FFFFFF",

    primary="#6366F1",
    primary_hover="#7C7EF6",
    primary_pressed="#4F46E5",
    primary_soft="#262A45",
    primary_soft_border="#4F46E5",
    primary_text="#C7D2FE",
    primary_disabled="#3A3F63",

    selection_bg="#312E5C",
    selection_text="#E0E7FF",

    danger="#F87171",
    danger_hover="#EF4444",
    danger_pressed="#DC2626",
    danger_text="#FCA5A5",
    danger_soft="#3A2427",
    danger_soft_border="#7F1D1D",

    success="#34D399",
    success_text="#6EE7B7",
    success_soft="#14342B",

    warning="#FBBF24",
    warning_text="#FCD34D",
    warning_soft="#3A2E1A",

    info="#60A5FA",
    info_text="#93C5FD",
    info_soft="#1E2A3D",

    scrollbar_handle="#3A404C",
    scrollbar_handle_hover="#4B5563",
    scrollbar_track="#21252E",

    dropzone_border="#3A404C",
    dropzone_bg="#1D212A",

    shadow_rgba=(0, 0, 0, 130),
)


PALETTES = {"light": LIGHT, "dark": DARK}

# 当前生效的 Palette。ThemeManager.apply() 会更新它，
# 供无法走 QSS 的代码 (阴影颜色、逐单元格 QColor、HTML 详情) 取用。
_ACTIVE_PALETTE: Palette = LIGHT


def active_palette() -> Palette:
    """返回当前生效的配色。未初始化 ThemeManager 时回退到亮色。"""
    return _ACTIVE_PALETTE


def _hex_to_colorref(hex_value: str) -> int:
    """把 #RRGGBB 转为 Windows COLORREF (0x00BBGGRR)。"""
    normalized = hex_value.lstrip("#")
    red = int(normalized[0:2], 16)
    green = int(normalized[2:4], 16)
    blue = int(normalized[4:6], 16)
    return (blue << 16) | (green << 8) | red


def _set_dwm_int_attribute(hwnd: int, attribute: int, value: int) -> bool:
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


def _refresh_windows_frame(hwnd: int):
    try:
        # SWP_NOSIZE | SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED
        flags = 0x0001 | 0x0002 | 0x0004 | 0x0010 | 0x0020
        ctypes.windll.user32.SetWindowPos(
            ctypes.c_void_p(int(hwnd)),
            None,
            0,
            0,
            0,
            0,
            flags,
        )
        # RDW_INVALIDATE | RDW_UPDATENOW | RDW_FRAME
        ctypes.windll.user32.RedrawWindow(ctypes.c_void_p(int(hwnd)), None, None, 0x0501)
    except Exception:
        pass


def set_windows_title_bar_dark_mode(
    hwnd: int,
    enabled: bool,
    caption_color: str | None = None,
    text_color: str | None = None,
    border_color: str | None = None,
) -> bool:
    """在 Windows 原生标题栏上应用深色模式。非 Windows 或失败时静默跳过。"""
    if sys.platform != "win32" or not hwnd:
        return False

    mode_value = 1 if enabled else 0
    ok = _set_dwm_int_attribute(hwnd, 20, mode_value)
    if not ok:
        ok = _set_dwm_int_attribute(hwnd, 19, mode_value)

    # Windows 11 支持显式标题栏配色；Windows 10 会返回无效参数，忽略即可。
    for attribute, color in ((35, caption_color), (36, text_color), (34, border_color)):
        if color:
            _set_dwm_int_attribute(hwnd, attribute, _hex_to_colorref(color))

    _refresh_windows_frame(hwnd)

    return ok


def apply_windows_title_bar_theme(widget, palette: Palette) -> bool:
    """让 QWidget 对应的 Windows 原生标题栏跟随当前主题。"""
    if sys.platform != "win32":
        return False
    try:
        hwnd = int(widget.winId())
    except Exception:
        return False
    return set_windows_title_bar_dark_mode(
        hwnd,
        palette.is_dark,
        caption_color=palette.surface,
        text_color=palette.text,
        border_color=palette.border,
    )


# ============================================================================
# QSS 模板  (使用 string.Template 的 $token 占位符, 避开 QSS 的花括号冲突)
# ============================================================================
_APP_QSS_TEMPLATE = Template(r"""
/* ============================ 全局基础 ============================ */
QWidget {
    font-family: "Segoe UI", "Microsoft YaHei", sans-serif;
    font-size: 13px;
    color: $text_body;
}
QMainWindow { background-color: $window; }
QToolTip {
    background-color: $surface;
    color: $text_body;
    border: 1px solid $border;
    border-radius: 6px;
    padding: 5px 8px;
}

/* ============================ 主窗口骨架 ============================ */
#header, #leftSidebar, #rightSidebar, #footer, #presetBar { background-color: $surface; }
#header { border-bottom: 1px solid $border; }
#leftSidebar { border-right: 1px solid $border; }
#rightSidebar { border-left: 1px solid $border; }
#footer { border-top: 1px solid $border; }
#presetBar { border-top: 1px solid $border; border-bottom: 1px solid $border; }
#presetLabel { color: $text_muted; font-size: 12px; font-weight: 700; }
#presetSummary { color: $text_muted; font-size: 12px; }
#mainView { background-color: $window; }

#topBtn { background: transparent; border: none; font-weight: 600; color: $text_body; padding: 7px 14px; border-radius: 8px; }
#topBtn:hover { background-color: $surface_hover; color: $text; }

/* ---- 左侧导航 ---- */
#navTitle { color: $text_faint; font-size: 11px; font-weight: bold; letter-spacing: 1px; margin-bottom: 6px; }
#navBtn { text-align: left; padding: 12px 14px; border: 1px solid transparent; border-radius: 10px; color: $text_body; font-weight: 600; background-color: transparent; }
#navBtn:hover { background-color: $surface_hover; border-color: $border; }
#navBtn:checked { background-color: $primary_soft; color: $primary_text; font-weight: 700; border-color: $primary_soft_border; }

/* ---- 卡片 ---- */
#importCard, #listContainer { background-color: $surface; border: 1px solid $border; border-radius: 16px; }
#sectionTitle { font-size: 14px; font-weight: 700; color: $text; }
#sectionHint { color: $text_muted; font-size: 12px; }
#listTitle { font-weight: 700; color: $text; }
#rightPanelTitle { font-weight: bold; font-size: 14px; color: $text; }

/* ---- 拖拽热区 ---- */
#dropZone { border: 2px dashed $dropzone_border; border-radius: 14px; background-color: $dropzone_bg; color: $text_muted; font-weight: 600; }
#dropZone:hover { border-color: $primary; background-color: $primary_soft; color: $primary_text; }
#dropZone[dragActive="true"] { border: 2px dashed $primary; background-color: $primary_soft; color: $primary_text; }

/* ---- 通用按钮 ---- */
#secondaryBtn { padding: 8px 14px; border: 1px solid $border_strong; border-radius: 9px; background-color: $surface; color: $text_body; font-weight: 600; }
#secondaryBtn:hover { background-color: $surface_hover; border-color: $primary_soft_border; }
#mutedLabel { color: $text_muted; font-size: 12px; }
#mutedSmall { color: $text_muted; font-size: 11px; }
#listHeader { border-bottom: 1px solid $border; background-color: $surface_alt; border-top-left-radius: 16px; border-top-right-radius: 16px; }
#sectionCaption { color: $text_faint; font-size: 12px; font-weight: bold; margin-bottom: 4px; }

/* ---- 规则勾选项 (右栏) ---- */
#optionTitle { font-weight: 500; color: $text_body; }
#optionDesc { color: $text_muted; font-size: 11px; margin-left: 24px; }

/* ============================ 文件树 ============================ */
QTreeWidget { show-decoration-selected: 1; border: none; background-color: $surface; color: $text_body; outline: none; border-bottom-left-radius: 16px; border-bottom-right-radius: 16px; alternate-background-color: $surface_alt2; }
QTreeWidget::item { padding: 7px; border-bottom: 1px solid $border_subtle; }
QTreeWidget::item:hover:!selected { background-color: $surface_hover; color: $text_body; }
QTreeWidget::item:selected { background-color: $primary_soft; color: $primary_text; }
QTreeWidget::branch { background-color: $surface; border: none; }
QTreeWidget::branch:alternate { background-color: $surface_alt2; }
QTreeWidget::branch:hover { background-color: $surface_hover; }
QTreeWidget::branch:has-siblings:!adjoins-item,
QTreeWidget::branch:has-siblings:adjoins-item,
QTreeWidget::branch:!has-children:!has-siblings:adjoins-item { border-image: none; image: none; }
QTreeWidget::branch:selected { background-color: $primary_soft; }
QTreeWidget::branch:selected:hover { background-color: $primary_soft; }
QHeaderView::section { background-color: $surface; border: none; border-bottom: 1px solid $border; padding: 8px; color: $text_muted; font-weight: 600; text-align: left; }

/* ---- 右栏摘要 / 风险提示 ---- */
#rightHeader { border-bottom: 1px solid $border; }
#selectionSummary { color: $text_body; background-color: $surface_alt; border: 1px solid $border; border-radius: 8px; padding: 8px 10px; qproperty-alignment: 'AlignCenter'; }
#dangerHint { color: $danger_text; background-color: $danger_soft; border: 1px solid $danger_soft_border; border-radius: 8px; padding: 8px 10px; }

/* ---- 滚动区 ---- */
#settingsScroll { border: none; background-color: transparent; margin: 0; padding: 0; }
#settingsScroll > QWidget { background-color: transparent; }
#settingsScroll > QWidget > QWidget { background-color: transparent; margin: 0; padding: 0; }

QScrollBar:vertical { border: none; background: transparent; width: 8px; margin: 0px; }
QScrollBar::handle:vertical { background: $scrollbar_handle; min-height: 30px; border-radius: 4px; }
QScrollBar::handle:vertical:hover { background: $scrollbar_handle_hover; }
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0px; background: none; }
QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical { background: transparent; }

QScrollBar:horizontal { border: none; background: $scrollbar_track; height: 14px; margin: 2px 8px 4px 8px; border-radius: 7px; }
QScrollBar::handle:horizontal { background: $scrollbar_handle_hover; min-width: 44px; border-radius: 7px; margin: 1px; }
QScrollBar::handle:horizontal:hover { background: $scrollbar_handle; }
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal { width: 0px; background: none; }
QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal { background: $scrollbar_track; border-radius: 7px; }

/* ---- 复选框 ---- */
QCheckBox { outline: none; spacing: 8px; color: $text_body; }
QCheckBox:focus { outline: none; }
QCheckBox::indicator { width: 16px; height: 16px; border-radius: 4px; border: 1px solid $border_strong; background: $surface; margin-top: 1px; }
QCheckBox::indicator:checked { background: $primary; border-color: $primary; }
#dangerCheck { color: $danger; }

/* ---- 底部操作 ---- */
#actionBtn { padding: 8px 16px; border: 1px solid $border_strong; border-radius: 10px; background-color: $surface; color: $text_body; font-weight: 600; }
#actionBtn:hover { background-color: $surface_hover; border-color: $primary_soft_border; }
#presetBtn { padding: 7px 15px; border: 1px solid $border_strong; border-radius: 8px; background-color: $surface; color: $text_body; font-weight: 600; }
#presetBtn:hover { background-color: $surface_hover; border-color: $primary_soft_border; }
#presetBtn:checked { background-color: $primary_soft; border-color: $primary_soft_border; color: $primary_text; }
#presetBtn:focus { outline: none; }
#footerSummary { color: $text; font-weight: 700; }
#processingHint { color: $primary_text; min-width: 0px; }
#footerHint { color: $text_muted; min-width: 0px; }
#footerHint[danger="true"] { color: $danger_text; font-weight: 600; }
#modeRadio { color: $text_body; }
#startBtn { padding: 10px 24px; background-color: $primary; color: $text_on_primary; border-radius: 10px; font-weight: bold; border: none; }
#startBtn[stopMode="true"] { background-color: $danger_hover; }
#startBtn[stopMode="true"]:hover { background-color: $danger_pressed; }
#startBtn[stopMode="true"]:pressed { background-color: $danger_pressed; }
#startBtn:hover { background-color: $primary_hover; }
#startBtn:pressed { background-color: $primary_pressed; }
#startBtn:disabled { background-color: $primary_disabled; color: $surface; }

/* ============================ 对话框 ============================ */
#dialogBg { background-color: $surface; border: 1px solid $border; border-radius: 12px; }
#dialogTitleBar { background-color: $surface_alt; border: none; border-bottom: 1px solid $border; border-top-left-radius: 12px; border-top-right-radius: 12px; }
#dialogTitle { font-weight: 700; color: $text; font-size: 13px; border: none; }
#dialogCloseBtn { background: transparent; border: none; font-size: 14px; color: $text_faint; border-radius: 6px; }
#dialogCloseBtn:hover { background-color: $surface_hover; color: $danger; }
#dialogContent { border: none; background-color: transparent; }

#dialogPrimaryBtn { background-color: $primary; color: $text_on_primary; border-radius: 8px; padding: 8px 16px; font-weight: 700; border: none; }
#dialogPrimaryBtn:hover { background-color: $primary_hover; }
#dialogPrimaryBtn:pressed { background-color: $primary_pressed; }
#dialogPrimaryBtn:disabled { background-color: $primary_disabled; color: $surface; }
#dialogDangerBtn { background-color: $danger; color: #FFFFFF; border-radius: 8px; padding: 8px 16px; font-weight: 700; border: none; }
#dialogDangerBtn:hover { background-color: $danger_hover; }
#dialogDangerBtn:pressed { background-color: $danger_pressed; }
#dialogSecondaryBtn { background-color: $surface; color: $text_body; border-radius: 8px; padding: 8px 16px; font-weight: 700; border: 1px solid $border_strong; }
#dialogSecondaryBtn:hover { background-color: $surface_hover; border-color: $primary_soft_border; color: $text; }
#dialogSecondaryBtn:pressed { background-color: $surface_alt; }
#dialogSecondaryBtn:checked { background-color: $primary_soft; border-color: $primary_soft_border; color: $primary_text; }
#dialogSectionTitle { color: $text_faint; font-size: 12px; font-weight: 700; border: none; }

/* ---- 对话框内文本 ---- */
#dialogHeading { color: $text; font-size: 17px; font-weight: 700; border: none; }
#dialogSubTitle { color: $text; font-size: 13px; font-weight: 700; border: none; }
#dialogBody { color: $text_body; font-size: 13px; border: none; }
#dialogMuted { color: $text_muted; font-size: 12px; border: none; }
#dialogCaption { color: $text_body; font-size: 12px; font-weight: 700; border: none; }
#dialogCodeBlock { color: $text_body; font-size: 12px; border: 1px solid $border; background-color: $surface_alt; border-radius: 8px; padding: 10px; font-family: Consolas, 'Courier New', monospace; }
#dialogStatus { font-size: 12px; border: none; }
#dialogStatus[state="success"] { color: $success_text; }
#dialogStatus[state="error"] { color: $danger_text; }
#dialogInfoPanel { color: $text_body; background-color: $surface_alt; border: 1px solid $border; border-radius: 8px; padding: 7px 10px; font-size: 12px; }
#dialogEmptyPanel { color: $text_muted; background-color: $surface_alt; border: 1px dashed $border_strong; border-radius: 8px; padding: 18px; font-size: 12px; }

/* ---- 已签名文件提示对话框 ---- */
#signedWarnCard { background-color: $warning_soft; border: 1px solid $warning; border-radius: 10px; }
#signedWarnCard QLabel { border: none; background: transparent; }
#signedWarnIcon { font-size: 22px; border: none; background: transparent; }
#signedWarnTitle { color: $warning_text; font-size: 14px; font-weight: 700; border: none; }
#signedWarnDesc { color: $text_body; font-size: 12px; border: none; }
#signedListCaption { color: $text_muted; font-size: 12px; font-weight: 700; border: none; }
#signedListScroll { background-color: $surface_alt; border: 1px solid $border; border-radius: 8px; }
#signedListScroll > QWidget { background: transparent; }
#signedFileList { color: $text_body; font-size: 12px; background: transparent; border: none; padding: 10px 12px; line-height: 150%; }

/* ---- 消息框 ---- */
#msgIcon { font-size: 36px; border: none; background: transparent; }
#msgText { color: $text_body; font-size: 13px; border: none; }
#msgTextBlock { color: $text_body; font-size: 12px; border: 1px solid $border; background-color: $surface_alt; border-radius: 8px; padding: 10px; font-family: Consolas, 'Courier New', monospace; }

/* ---- 数据向导选择按钮 ---- */
#choiceToggleBtn { color: $text_body; background-color: $surface_alt; border: 1px solid $border_strong; border-radius: 8px; padding: 8px 14px; font-weight: 600; text-align: center; }
#choiceToggleBtn:hover { background-color: $surface_hover; }
#choiceToggleBtn:checked { color: $primary_text; background-color: $primary_soft; border-color: $primary_soft_border; }

/* ---- 数据向导目录卡片 / 预览表 ---- */
#wizardCard { background-color: $surface_alt; border: 1px solid $border; border-radius: 8px; }
#wizardCard QLabel { border: none; background: transparent; }
#previewTable { background-color: $surface; border: 1px solid $border; border-radius: 8px; gridline-color: $border; color: $text_body; selection-background-color: $primary_soft; selection-color: $primary_text; }
#previewTable QHeaderView::section { background-color: $surface_alt; border: none; border-bottom: 1px solid $border; padding: 7px; color: $text_muted; font-weight: 700; }
#previewTable QTableCornerButton::section { background-color: $surface_alt; border: none; border-bottom: 1px solid $border; }

/* ---- 设置对话框 ---- */
#settingsPathCard { background-color: $surface_alt; border: 1px solid $border; border-radius: 10px; }
#settingsPathHint { color: $text_muted; font-size: 12px; border: none; }
#settingsPathEdit { background-color: $surface; border: 1px solid $border_strong; border-radius: 8px; padding: 8px 10px; color: $text_body; selection-background-color: $selection_bg; selection-color: $selection_text; }
#settingsPathEdit:focus { border-color: $primary; }
#settingsPathStatus { font-size: 12px; border: none; padding: 0 2px; }
#settingsPathStatus[state="empty"] { color: $text_faint; }
#settingsPathStatus[state="valid"] { color: $primary_text; }
#settingsPathStatus[state="invalid"] { color: $danger; font-weight: 600; }
#settingsWorkerSpin { background-color: $surface; border: 1px solid $border_strong; border-right: none; border-top-left-radius: 8px; border-bottom-left-radius: 8px; border-top-right-radius: 0; border-bottom-right-radius: 0; padding: 4px 10px; color: $text_body; min-height: 24px; min-width: 78px; }
#settingsWorkerSpin:focus { border-color: $primary; }
#settingsWorkerSpin:disabled { background-color: $surface_alt; border-color: $border; color: $text_muted; }
#settingsSpinStepper { background-color: $surface_alt; border: 1px solid $border_strong; border-left: none; border-top-right-radius: 8px; border-bottom-right-radius: 8px; min-width: 24px; max-width: 24px; }
#settingsSpinStepBtn { background: transparent; border: none; color: $text_muted; font-size: 9px; font-weight: 700; padding: 0; min-width: 23px; max-width: 23px; min-height: 14px; max-height: 14px; }
#settingsSpinStepBtn:hover { background-color: $surface_hover; color: $text; }
#settingsSpinStepBtn:pressed { background-color: $primary_soft; color: $primary_text; }
#settingsSpinStepBtn:disabled { color: $text_faint; background: transparent; }
#settingsSpinStepBtn[direction="up"] { border-top-right-radius: 7px; border-bottom: 1px solid $border_subtle; }
#settingsSpinStepBtn[direction="down"] { border-bottom-right-radius: 7px; }

/* ---- 主题选择分段按钮 ---- */
#themeSegBtn { color: $text_body; background-color: $surface; border: 1px solid $border_strong; padding: 7px 14px; font-weight: 600; border-radius: 8px; }
#themeSegBtn:hover { background-color: $surface_hover; }
#themeSegBtn:checked { color: $primary_text; background-color: $primary_soft; border-color: $primary_soft_border; }

/* ---- 日志对话框 ---- */
#statCard { background-color: $surface_alt; border: 1px solid $border; border-radius: 8px; }
#statCard QLabel { border: none; background: transparent; padding: 0; }
#statCardTitle { color: $text_muted; font-size: 11px; font-weight: 700; }
#statCardValue { color: $text; font-size: 20px; font-weight: 700; }
#logSummaryFrame { background-color: $surface; border: 1px solid $border; border-radius: 8px; }
#logSummaryTable { background-color: $surface; alternate-background-color: $surface_alt2; color: $text_body; border: none; border-radius: 0; gridline-color: $border; selection-background-color: $primary_soft; selection-color: $primary_text; outline: none; }
#logSummaryTable::item { padding: 7px; border-bottom: 1px solid $border_subtle; }
#logSummaryTable::item:selected { background-color: $primary_soft; color: $primary_text; }
#logSummaryTable QHeaderView::section { background-color: $surface_alt; border: none; border-bottom: 1px solid $border; padding: 8px; color: $text_muted; font-weight: 700; }
#logSummaryTable QTableCornerButton::section { background-color: $surface_alt; border: none; border-bottom: 1px solid $border; }
#logSplitter::handle:vertical { background-color: transparent; border: none; height: 10px; margin: 2px 0; }
#logSplitter::handle:vertical:hover { background-color: $primary_soft; border-radius: 5px; }
#logDetailTextEdit { background-color: $surface_alt; border: 1px solid $border; border-radius: 8px; padding: 10px; color: $text_body; font-family: Consolas, 'Courier New', monospace; font-size: 12px; selection-background-color: $selection_bg; selection-color: $selection_text; }

/* ---- 关于对话框 ---- */
#aboutHeroCard { background-color: $primary_soft; border: 1px solid $primary_soft_border; border-radius: 10px; }
#aboutInfoCard { background-color: $surface_alt; border: 1px solid $border; border-radius: 10px; }
#aboutBrandTitle { font-size: 22px; font-weight: 700; color: $primary_text; border: none; }
#aboutBadge { background-color: $surface; color: $primary_text; border: 1px solid $primary_soft_border; border-radius: 999px; padding: 4px 10px; font-weight: 600; }
#aboutTitle { color: $text; font-size: 13px; font-weight: 700; border: none; }
#aboutText { color: $text_body; font-size: 12px; border: none; }
#aboutIntro { color: $text_body; font-size: 13px; border: none; }
""")


def build_app_qss(palette: Palette) -> str:
    """依据 Palette 渲染整份应用级 QSS。"""
    return _APP_QSS_TEMPLATE.substitute(palette.as_dict())


# ============================================================================
# 日志状态色 —— 无法用 QSS 表达 (逐单元格 QColor / HTML), 从当前 Palette 取值
# ============================================================================
def log_status_colors(palette: Palette, tags) -> tuple[str, str, str]:
    """返回 (强调色, 背景色, 文本色), 供日志列表单元格与详情 HTML 使用。"""
    if "failure" in tags:
        return palette.danger, palette.danger_soft, palette.danger_text
    if "skip" in tags:
        return palette.warning, palette.warning_soft, palette.warning_text
    if "precheck" in tags:
        return palette.info, palette.info_soft, palette.info_text
    if "success" in tags:
        return palette.success, palette.success_soft, palette.success_text
    return palette.text_faint, palette.surface, palette.text_body


# ============================================================================
# 主题管理器
# ============================================================================
class ThemeManager(QObject):
    """解析主题模式并应用到 QApplication；系统主题变化时自动刷新。

    模式:
        "system" —— 跟随系统 (随 OS 深/浅色实时切换)
        "light"  —— 强制亮色
        "dark"   —— 强制暗色
    """

    changed = Signal(object)  # 发出当前生效的 Palette

    _MODES = ("system", "light", "dark")

    def __init__(self, app: QApplication, mode: str = "system"):
        super().__init__(app)
        self._app = app
        self._mode = mode if mode in self._MODES else "system"
        self._palette = LIGHT
        self._applying = False

        hints = app.styleHints()
        # 系统主题变化时, 若处于跟随模式则实时刷新
        hints.colorSchemeChanged.connect(self._on_system_scheme_changed)

    # --- 查询 ---
    @property
    def mode(self) -> str:
        return self._mode

    def current_palette(self) -> Palette:
        return self._palette

    def _resolve_palette(self) -> Palette:
        if self._mode == "light":
            return LIGHT
        if self._mode == "dark":
            return DARK
        # system
        scheme = self._app.styleHints().colorScheme()
        return DARK if scheme == Qt.ColorScheme.Dark else LIGHT

    # --- 应用 ---
    def apply(self):
        global _ACTIVE_PALETTE
        self._applying = True
        try:
            # 让 Qt 自身控件 (下拉弹窗、原生绘制) 与选定方案保持一致
            if self._mode == "system":
                self._app.styleHints().setColorScheme(Qt.ColorScheme.Unknown)
            elif self._mode == "dark":
                self._app.styleHints().setColorScheme(Qt.ColorScheme.Dark)
            else:
                self._app.styleHints().setColorScheme(Qt.ColorScheme.Light)
            self._palette = self._resolve_palette()
            _ACTIVE_PALETTE = self._palette
        finally:
            self._applying = False
        self._app.setStyleSheet(build_app_qss(self._palette))
        self.changed.emit(self._palette)

    def set_mode(self, mode: str):
        if mode not in self._MODES:
            return
        self._mode = mode
        self.apply()

    def _on_system_scheme_changed(self, _scheme):
        if self._applying:
            return
        if self._mode == "system":
            self.apply()
