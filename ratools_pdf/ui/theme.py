"""集中式主题系统（Win11 Fluent 清爽专业风）。

本模块是整个应用视觉层的唯一颜色来源：

* ``Palette``   —— 语义化设计令牌 (token) 的数据结构。
* ``LIGHT`` / ``DARK`` —— 两套配色 (Win11 Fluent 亮色 + 匹配暗色)。
* ``build_app_qss``   —— 依据当前 Palette 生成整份应用级 QSS。
* ``ThemeManager``    —— 解析「跟随系统 / 手动亮色 / 手动暗色」，
  在 :class:`~PySide6.QtWidgets.QApplication` 上应用样式，并在系统主题
  变化时实时刷新。

设计原则：所有颜色都以令牌形式定义，UI 组件不再硬编码十六进制色值；
组件通过 objectName / 动态属性挂接到中央 QSS，切换主题只需要重新调用
``app.setStyleSheet(...)`` 即可级联到所有已打开与将要打开的窗口。
"""

from __future__ import annotations

import os
import sys
import tempfile
from dataclasses import dataclass, asdict
from string import Template

from PySide6.QtCore import QObject, Qt, Signal
from PySide6.QtGui import QColor, QFont, QFontDatabase
from PySide6.QtWidgets import QApplication

from ratools_pdf.ui import icons, win32


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
    sidebar: str       # 左侧导航栏底色
    surface: str       # 卡片 / 面板
    surface_alt: str   # 次级填充
    surface_alt2: str  # 交替行背景
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

    # --- 主色 (Win11 Fluent 企业蓝) ---
    primary: str
    primary_hover: str
    primary_pressed: str
    primary_soft: str        # 选中态浅底
    primary_soft_border: str
    primary_text: str        # 浅底之上的文字
    primary_disabled: str
    accent_bar: str          # 左侧导航激活指示竖条

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
# Win11 Fluent 清爽亮色 —— 主色 #0F6CBD
# ---------------------------------------------------------------------------
LIGHT = Palette(
    name="light",
    is_dark=False,

    window="#F7F8FA",
    sidebar="#F3F5F9",
    surface="#FFFFFF",
    surface_alt="#F8FAFC",
    surface_alt2="#FBFCFE",
    surface_hover="#F1F5F9",

    border="#E2E8F0",
    border_strong="#CBD5E1",
    border_subtle="#F1F5F9",

    text="#0F172A",
    text_body="#334155",
    text_muted="#64748B",
    text_faint="#94A3B8",
    text_on_primary="#FFFFFF",

    primary="#0F6CBD",
    primary_hover="#115EA3",
    primary_pressed="#0E4F8A",
    primary_soft="#EBF3FC",
    primary_soft_border="#BFDBFE",
    primary_text="#0F6CBD",
    primary_disabled="#B8D4EE",
    accent_bar="#0F6CBD",

    selection_bg="#DCEBFA",
    selection_text="#0E4F8A",

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

    info="#0EA5E9",
    info_text="#0369A1",
    info_soft="#F0F9FF",

    scrollbar_handle="#CBD5E1",
    scrollbar_handle_hover="#94A3B8",
    scrollbar_track="#EAF0F6",

    dropzone_border="#BFD7F5",
    dropzone_bg="#F8FBFF",

    shadow_rgba=(15, 23, 42, 36),
)


# ---------------------------------------------------------------------------
# Win11 Fluent 匹配暗色
# ---------------------------------------------------------------------------
DARK = Palette(
    name="dark",
    is_dark=True,

    window="#0F1115",
    sidebar="#14171D",
    surface="#191C22",
    surface_alt="#21252E",
    surface_alt2="#1D212A",
    surface_hover="#262B36",

    border="#2E333D",
    border_strong="#3E4654",
    border_subtle="#252A33",

    text="#F1F3F6",
    text_body="#D4D8DF",
    text_muted="#9AA2AF",
    text_faint="#6B7280",
    text_on_primary="#FFFFFF",

    primary="#2B88D8",
    primary_hover="#3B96E8",
    primary_pressed="#1E6FB8",
    primary_soft="#1B2B44",
    primary_soft_border="#2B5B94",
    primary_text="#60B4FF",
    primary_disabled="#2B3B52",
    accent_bar="#60B4FF",

    selection_bg="#1E3A8A",
    selection_text="#DBEAFE",

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

    info="#38BDF8",
    info_text="#7DD3FC",
    info_soft="#16293D",

    scrollbar_handle="#3A404C",
    scrollbar_handle_hover="#4B5563",
    scrollbar_track="#21252E",

    dropzone_border="#334155",
    dropzone_bg="#1A202B",

    shadow_rgba=(0, 0, 0, 130),
)


PALETTES = {"light": LIGHT, "dark": DARK}

# 当前生效的 Palette。ThemeManager.apply() 会更新它，
# 供无法走 QSS 的代码 (阴影颜色、逐单元格 QColor、HTML 详情) 取用。
_ACTIVE_PALETTE: Palette = LIGHT


def active_palette() -> Palette:
    """返回当前生效的配色。未初始化 ThemeManager 时回退到亮色。"""
    return _ACTIVE_PALETTE


# ============================================================================
# 排版与字体系统
# ============================================================================
DEFAULT_UI_FONT_FAMILIES: tuple[str, ...] = ("Microsoft YaHei UI", "Segoe UI")

PREFERRED_MONO_FONT_FAMILIES: tuple[str, ...] = (
    "Cascadia Mono",
    "JetBrains Mono",
    "Consolas",
    "Microsoft YaHei UI",
    "Courier New",
)

FONT_SCALES: dict[str, int] = {
    "compact": -1,
    "standard": 0,
    "comfortable": 1,
}

FONT_SCALE_LABELS: dict[str, str] = {
    "compact": "紧凑",
    "standard": "标准",
    "comfortable": "舒适",
}


@dataclass(frozen=True)
class Typography:
    """应用级排版与字重/字号令牌。"""

    scale: str
    primary_family: str
    ui_families: tuple[str, ...]
    ui_font_css: str
    mono_families: tuple[str, ...]
    mono_font_css: str
    has_medium_weight: bool
    weight_regular: int
    weight_medium: int
    weight_bold: int
    fs_xs: int
    fs_sm: int
    fs_base: int
    fs_md: int
    fs_lg: int
    fs_xl: int
    fs_2xl: int

    def as_qss_tokens(self) -> dict:
        return {
            "font_ui": self.ui_font_css,
            "font_mono": self.mono_font_css,
            "w_regular": self.weight_regular,
            "w_medium": self.weight_medium,
            "w_bold": self.weight_bold,
            "fs_xs": self.fs_xs,
            "fs_sm": self.fs_sm,
            "fs_base": self.fs_base,
            "fs_md": self.fs_md,
            "fs_lg": self.fs_lg,
            "fs_xl": self.fs_xl,
            "fs_2xl": self.fs_2xl,
        }


def installed_font_families() -> set[str]:
    """安全获取当前系统已安装字体族集合。"""
    try:
        if QApplication.instance() is None:
            return set()
        return set(QFontDatabase.families())
    except Exception:
        return set()


def font_has_medium_weight(family: str) -> bool:
    """检查指定字体族是否具备真实的 Medium/Semibold 字重。"""
    if not family or QApplication.instance() is None:
        return False
    try:
        styles = QFontDatabase.styles(family)
    except Exception:
        return False
    return any(
        token in style.lower()
        for style in styles
        for token in ("medium", "semibold", "demibold")
    )


def _format_css_families(families: tuple[str, ...], generic: str) -> str:
    quoted = [f'"{fam}"' for fam in families]
    quoted.append(generic)
    return ", ".join(quoted)


def _dedupe_families(families: list[str]) -> tuple[str, ...]:
    result: list[str] = []
    seen: set[str] = set()
    for fam in families:
        if fam and fam not in seen:
            seen.add(fam)
            result.append(fam)
    return tuple(result)


def resolve_typography(
    scale: str = "standard",
    installed: set[str] | None = None,
) -> Typography:
    """依据字号偏好与系统已安装等宽字体解析应用排版令牌。"""
    norm_scale = scale if scale in FONT_SCALES else "standard"
    delta = FONT_SCALES[norm_scale]
    families = installed_font_families() if installed is None else set(installed)

    ui_list = DEFAULT_UI_FONT_FAMILIES
    primary_family = ui_list[0]
    mono_list = _dedupe_families(
        [f for f in PREFERRED_MONO_FONT_FAMILIES if (not families or f in families)]
        + ["Consolas", "Microsoft YaHei UI", "Courier New"]
    )

    return Typography(
        scale=norm_scale,
        primary_family=primary_family,
        ui_families=ui_list,
        ui_font_css=_format_css_families(ui_list, "sans-serif"),
        mono_families=mono_list,
        mono_font_css=_format_css_families(mono_list, "monospace"),
        has_medium_weight=font_has_medium_weight(primary_family),
        weight_regular=400,
        weight_medium=500,
        weight_bold=700,
        fs_xs=max(10, 11 + delta),
        fs_sm=12 + delta,
        fs_base=13 + delta,
        fs_md=14 + delta,
        fs_lg=15 + delta,
        fs_xl=17 + delta,
        fs_2xl=20 + delta,
    )


_ACTIVE_TYPOGRAPHY: Typography | None = None


def active_typography() -> Typography:
    """返回当前生效的排版令牌。"""
    global _ACTIVE_TYPOGRAPHY
    if _ACTIVE_TYPOGRAPHY is not None:
        return _ACTIVE_TYPOGRAPHY
    if QApplication.instance() is None:
        return resolve_typography(installed=set())
    _ACTIVE_TYPOGRAPHY = resolve_typography()
    return _ACTIVE_TYPOGRAPHY


def build_app_qfont(typography: Typography | None = None) -> QFont:
    """构建应用级抗锯齿 UI 字体。"""
    typo = typography or active_typography()
    font = QFont()
    font.setFamilies(list(typo.ui_families))
    font.setPixelSize(typo.fs_base)
    font.setWeight(QFont.Weight.Normal)
    font.setStyleStrategy(QFont.StyleStrategy.PreferAntialias | QFont.StyleStrategy.PreferQuality)
    font.setHintingPreference(QFont.HintingPreference.PreferDefaultHinting)
    return font


def build_mono_qfont(typography: Typography | None = None, size_delta: int = 0) -> QFont:
    """构建日志与诊断视图使用的等宽字体（附带清晰中文回退）。"""
    typo = typography or active_typography()
    font = QFont()
    font.setFamilies(list(typo.mono_families))
    font.setPixelSize(max(10, typo.fs_sm + size_delta))
    font.setWeight(QFont.Weight.Normal)
    font.setStyleStrategy(QFont.StyleStrategy.PreferAntialias | QFont.StyleStrategy.PreferQuality)
    return font


# ---------------------------------------------------------------------------
# 勾选框对勾图标（兼容旧导出入口）
# ---------------------------------------------------------------------------
_CHECKMARK_SVG = (
    '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 16 16">'
    '<path d="M3.8 8.6 L6.8 11.4 L12.2 4.9" fill="none" stroke="{color}" '
    'stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>'
)

_checkmark_paths: dict = {}


def checkmark_icon_path(color: str = "#FFFFFF") -> str:
    """返回对勾 SVG 的文件路径 (QSS url 格式, 正斜杠)；按颜色缓存生成。"""
    cached = _checkmark_paths.get(color)
    if cached and os.path.exists(cached):
        return cached
    file_path = os.path.join(
        tempfile.gettempdir(),
        f"ratools_pdf_check_{color.lstrip('#').lower()}.svg",
    )
    try:
        with open(file_path, "w", encoding="utf-8") as fh:
            fh.write(_CHECKMARK_SVG.format(color=color))
    except OSError:
        return ""
    normalized = file_path.replace("\\", "/")
    _checkmark_paths[color] = normalized
    return normalized


# DWM/Win32 细节集中在 ui.win32；这里保留同名入口以兼容既有调用与测试。
set_windows_title_bar_dark_mode = win32.set_title_bar_dark_mode


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
    font-family: $font_ui;
    font-size: ${fs_base}px;
    font-weight: $w_regular;
    color: $text_body;
}
QMainWindow { background-color: $window; }
QToolTip {
    background-color: $surface;
    color: $text_body;
    border: 1px solid $border;
    border-radius: 6px;
    padding: 6px 10px;
}

/* ============================ 主窗口骨架 ============================ */
#rightSidebar, #footer, #presetBar { background-color: $surface; }
#middleContainer { border-top: 1px solid $border; }
#leftSidebar { background-color: $sidebar; border-right: 1px solid $border; }
#rightSidebar { border-left: 1px solid $border; }
#footer { border-top: 1px solid $border; }
#presetBar { border-top: 1px solid $border; }
#mainView { background-color: $window; }

/* ---- 左侧导航与底部全局入口 ---- */
#navTitle { color: $text_faint; font-size: ${fs_sm}px; font-weight: $w_medium; letter-spacing: 0.5px; margin: 2px 6px 2px 6px; }
#moduleSearchInput {
    background-color: $surface;
    border: 1px solid $border;
    border-radius: 6px;
    padding: 5px 8px;
    font-size: ${fs_sm}px;
    font-weight: $w_regular;
    color: $text_body;
    selection-background-color: $selection_bg;
    selection-color: $selection_text;
}
#moduleSearchInput:hover { border-color: $border_strong; }
#moduleSearchInput:focus { border-color: $primary; }
#moduleSearchEmpty { color: $text_faint; font-size: ${fs_sm}px; padding: 12px 8px; }
#navBtn {
    text-align: left;
    padding: 10px 36px 10px 12px;
    border: 1px solid transparent;
    border-radius: 6px;
    color: $text_body;
    font-weight: $w_regular;
    background-color: transparent;
}
#navBtn:hover { background-color: $surface_hover; color: $text; }
#navBtn:checked {
    background-color: $primary_soft;
    color: $primary_text;
    font-weight: $w_bold;
    border-color: $primary_soft_border;
}
#sidebarSep { background-color: $border; border: none; margin: 4px 2px; }
#sidebarUtilBtn {
    text-align: left;
    padding: 8px 12px;
    border: 1px solid transparent;
    border-radius: 6px;
    color: $text_body;
    font-weight: $w_regular;
    background-color: transparent;
}
#sidebarUtilBtn:hover { background-color: $surface_hover; color: $text; border-color: $border; }
#sidebarUtilBtn:pressed { background-color: $surface_alt; }
#sidebarFooterBtn, #themeQuickBtn {
    background-color: $surface;
    border: 1px solid $border;
    border-radius: 6px;
    color: $text_body;
    font-size: ${fs_sm}px;
    font-weight: $w_regular;
    padding: 6px 10px;
}
#sidebarFooterBtn:hover, #themeQuickBtn:hover {
    background-color: $surface_hover;
    border-color: $border_strong;
    color: $text;
}
#sidebarFooterBtn:pressed, #themeQuickBtn:pressed { background-color: $surface_alt; }

/* ---- 中央单层卡片 & 队列 ---- */
#queueCard, #importCard, #listContainer {
    background-color: $surface;
    border: 1px solid $border;
    border-radius: 8px;
}
#sectionTitle { font-size: ${fs_md}px; font-weight: $w_bold; color: $text; }
#sectionHint { color: $text_muted; font-size: ${fs_sm}px; }
#listTitle { font-weight: $w_bold; font-size: ${fs_md}px; color: $text; }
#rightPanelTitle { font-weight: $w_bold; font-size: ${fs_md}px; color: $text; }
#listHeader {
    border-bottom: 1px solid $border;
    background-color: $surface_alt;
    border-top-left-radius: 7px;
    border-top-right-radius: 7px;
}
#mutedLabel { color: $text_muted; font-size: ${fs_sm}px; }
#mutedSmall { color: $text_muted; font-size: ${fs_sm}px; }
#sectionCaption { color: $text_muted; font-size: ${fs_sm}px; font-weight: $w_medium; margin-bottom: 2px; }

/* ---- 居中空队列拖拽热区 ---- */
#emptyDropContainer { background-color: transparent; border: none; }
#dropZone {
    border: 2px dashed $dropzone_border;
    border-radius: 8px;
    background-color: $dropzone_bg;
    color: $text_muted;
    font-weight: $w_medium;
}
#dropZone:hover { border-color: $primary; background-color: $primary_soft; color: $primary_text; }
#dropZone[dragActive="true"] { border: 2px dashed $primary; background-color: $primary_soft; color: $primary_text; }
#emptyDropIconBadge {
    background-color: $primary_soft;
    border: 1px solid $primary_soft_border;
    border-radius: 8px;
}
#emptyDropTitle { font-size: ${fs_lg}px; font-weight: $w_bold; color: $text; border: none; background: transparent; }
#emptyDropDesc { font-size: ${fs_sm}px; color: $text_muted; border: none; background: transparent; }

/* ---- 三阶通用按钮 ---- */
#primaryBtn {
    padding: 7px 16px;
    border-radius: 6px;
    background-color: $primary;
    color: $text_on_primary;
    font-weight: $w_medium;
    border: none;
}
#primaryBtn:hover { background-color: $primary_hover; }
#primaryBtn:pressed { background-color: $primary_pressed; }
#primaryBtn:disabled { background-color: $primary_disabled; color: $surface; }

#secondaryBtn, #compactSecondaryBtn {
    border: 1px solid $border_strong;
    border-radius: 6px;
    background-color: $surface;
    color: $text_body;
    font-weight: $w_medium;
}
#secondaryBtn { padding: 7px 14px; }
#compactSecondaryBtn { padding: 5px 10px; }
#secondaryBtn:hover, #compactSecondaryBtn:hover { background-color: $surface_hover; border-color: $primary_soft_border; color: $text; }
#secondaryBtn:pressed, #compactSecondaryBtn:pressed { background-color: $surface_alt; }
#secondaryBtn:disabled, #compactSecondaryBtn:disabled { background-color: $surface_alt; color: $text_faint; border-color: $border; }

#configMenuBtn {
    padding: 7px 26px 7px 14px;
    border: 1px solid $border_strong;
    border-radius: 6px;
    background-color: $surface;
    color: $text_body;
    font-weight: $w_medium;
}
#configMenuBtn:hover { background-color: $surface_hover; border-color: $primary_soft_border; color: $text; }
#configMenuBtn:pressed { background-color: $surface_alt; }
#configMenuBtn:disabled { background-color: $surface_alt; color: $text_faint; border-color: $border; }
QPushButton#configMenuBtn::menu-indicator {
    image: url($chevron_down_url);
    subcontrol-origin: padding;
    subcontrol-position: center right;
    right: 8px;
    width: 13px;
    height: 13px;
}

#ghostBtn {
    padding: 6px 10px;
    border: 1px solid transparent;
    border-radius: 6px;
    background-color: transparent;
    color: $text_muted;
    font-weight: $w_medium;
}
#ghostBtn:hover { background-color: $surface_hover; color: $text; border-color: $border; }
#ghostBtn:pressed { background-color: $surface_alt; }
#ghostBtn:disabled { color: $text_faint; background: transparent; border-color: transparent; }

/* ---- 右栏规则卡片与语义标签 ---- */
#rightHeader { border-bottom: 1px solid $border; background-color: $surface; }
#selectionSummary {
    color: $text_body;
    background-color: $surface_alt;
    border: 1px solid $border;
    border-radius: 6px;
    padding: 3px 10px;
    font-size: ${fs_sm}px;
    font-weight: $w_medium;
}
#moduleStatsLabel { color: $text_muted; font-size: ${fs_sm}px; }
#moduleActionLink {
    background: transparent;
    border: none;
    color: $primary_text;
    font-size: ${fs_sm}px;
    font-weight: $w_medium;
    padding: 2px 6px;
    border-radius: 4px;
}
#moduleActionLink:hover { background-color: $primary_soft; }
#moduleActionLink:disabled { color: $text_faint; background: transparent; }

#ruleCard {
    background-color: $surface;
    border: 1px solid $border;
    border-radius: 8px;
}
#ruleCard:hover {
    background-color: $surface_hover;
    border-color: $border_strong;
}
#ruleCard[checked="true"] {
    background-color: $primary_soft;
    border-color: $primary_soft_border;
}
#ruleCard[searchHighlight="true"], #compressionCard[searchHighlight="true"] {
    border: 2px solid $primary;
    background-color: $primary_soft;
}
#ruleCard QCheckBox { background: transparent; }
#ruleTag {
    font-size: ${fs_xs}px;
    font-weight: $w_medium;
    border-radius: 4px;
    padding: 1px 7px;
    border: none;
}
#ruleTag[tone="info"] { background-color: $info_soft; color: $info_text; }
#ruleTag[tone="warning"] { background-color: $warning_soft; color: $warning_text; }
#ruleTag[tone="danger"] { background-color: $danger_soft; color: $danger_text; }

#optionTitle { font-weight: $w_medium; font-size: ${fs_base}px; color: $text; background: transparent; border: none; }
#optionDesc { color: $text_muted; font-size: ${fs_sm}px; font-weight: $w_regular; background: transparent; border: none; }

#searchGroupTitle { color: $text_muted; font-size: ${fs_sm}px; font-weight: $w_bold; background: transparent; border: none; }
#searchGroupCount { color: $text_faint; font-size: ${fs_xs}px; font-weight: $w_medium; background: transparent; border: none; }
#searchActionCard {
    background-color: $surface;
    border: 1px solid $border;
    border-radius: 8px;
}
#searchActionCard:hover {
    background-color: $surface_hover;
    border-color: $border_strong;
}
#searchActionCard QLabel { background: transparent; border: none; }
#searchEmptyCard {
    background-color: $surface_alt;
    border: 1px dashed $border_strong;
    border-radius: 8px;
}
#searchEmptyCard QLabel { background: transparent; border: none; }

#dangerHint { color: $danger_text; background-color: $danger_soft; border: 1px solid $danger_soft_border; border-radius: 6px; padding: 8px 10px; font-size: ${fs_sm}px; }

/* ============================ 文件树 ============================ */
QTreeWidget {
    show-decoration-selected: 1;
    border: none;
    background-color: $surface;
    color: $text_body;
    outline: none;
    border-bottom-left-radius: 7px;
    border-bottom-right-radius: 7px;
    alternate-background-color: $surface_alt2;
}
QTreeWidget#queueTree[dragActive="true"] {
    border: 2px dashed $primary;
    background-color: $primary_soft;
}
QTreeWidget::item { padding: 7px 6px; min-height: 28px; border-bottom: 1px solid $border_subtle; }
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
QTreeWidget::indicator {
    width: 16px;
    height: 16px;
    border-radius: 4px;
    border: 1px solid $border_strong;
    background: $surface;
}
QTreeWidget::indicator:hover { border-color: $primary; }
QTreeWidget::indicator:checked { background: $primary; border-color: $primary; image: url($checkmark_url); }
QTreeWidget::indicator:disabled { background: $surface_alt; border-color: $border; }
QHeaderView::section {
    background-color: $surface_alt;
    border: none;
    border-bottom: 1px solid $border;
    padding: 8px 10px;
    color: $text_muted;
    font-size: ${fs_sm}px;
    font-weight: $w_medium;
    text-align: left;
}

/* ---- 滚动区 ---- */
#settingsScroll { border: none; background-color: $surface; margin: 0; padding: 0; }
#settingsScroll > QWidget { border: none; background-color: $surface; }
#settingsScroll > QWidget > QWidget { border: none; background-color: $surface; margin: 0; padding: 0; }

QScrollBar:vertical { border: none; background: transparent; width: 8px; margin: 2px 0; }
QScrollBar::handle:vertical { background: $scrollbar_handle; min-height: 30px; border-radius: 4px; }
QScrollBar::handle:vertical:hover { background: $scrollbar_handle_hover; }
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0px; background: none; }
QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical { background: transparent; }

QScrollBar:horizontal { border: none; background: $scrollbar_track; height: 8px; margin: 2px 8px; border-radius: 4px; }
QScrollBar::handle:horizontal { background: $scrollbar_handle; min-width: 44px; border-radius: 4px; }
QScrollBar::handle:horizontal:hover { background: $scrollbar_handle_hover; }
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal { width: 0px; background: none; }
QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal { background: transparent; }

/* ---- 复选框 & 单选框 ---- */
QCheckBox { outline: none; spacing: 8px; color: $text_body; background: transparent; }
QCheckBox:focus { outline: none; }
QCheckBox::indicator { width: 16px; height: 16px; border-radius: 4px; border: 1px solid $border_strong; background: $surface; margin-top: 1px; }
QCheckBox::indicator:hover { border-color: $primary; }
QCheckBox::indicator:checked { background: $primary; border-color: $primary; image: url($checkmark_url); }
QCheckBox::indicator:disabled { background: $surface_alt; border-color: $border; }
#dangerCheck { color: $danger; font-weight: $w_medium; }

QRadioButton { outline: none; spacing: 8px; color: $text_body; background: transparent; }
QRadioButton:focus { outline: none; }
QRadioButton::indicator {
    width: 16px;
    height: 16px;
    border-radius: 8px;
    border: 1px solid $border_strong;
    background-color: $surface;
}
QRadioButton::indicator:hover { border-color: $primary; }
QRadioButton::indicator:checked {
    background-color: $primary;
    border-color: $primary;
    image: url($radio_dot_url);
}
QRadioButton::indicator:disabled { background: $surface_alt; border-color: $border; }

/* ---- 分段控制器 ---- */
#segmentedContainer, #modeSegContainer {
    background-color: $surface_alt;
    border: 1px solid $border;
    border-radius: 6px;
}
#segmentedBtn {
    background: transparent;
    border: 1px solid transparent;
    border-radius: 4px;
    padding: 5px 8px;
    color: $text_muted;
    font-size: ${fs_sm}px;
    font-weight: $w_medium;
}
#segmentedBtn:hover { color: $text; }
#segmentedBtn:checked {
    background-color: $primary_soft;
    border-color: $primary_soft_border;
    color: $primary_text;
    font-weight: $w_bold;
}

QRadioButton#modeRadio {
    color: $text_muted;
    background: transparent;
    border: 1px solid transparent;
    border-radius: 4px;
    padding: 4px 11px;
    font-size: ${fs_sm}px;
    font-weight: $w_medium;
    spacing: 0px;
}
QRadioButton#modeRadio::indicator { width: 0px; height: 0px; border: none; background: transparent; }
QRadioButton#modeRadio:hover { color: $text; }
QRadioButton#modeRadio:checked {
    background-color: $surface;
    color: $text;
    border: 1px solid $border;
}

/* ---- 预设栏 & 细进度条 & 底部操作 ---- */
#presetLabel { color: $text_muted; font-size: ${fs_sm}px; font-weight: $w_medium; }
#presetSummary { color: $text_muted; font-size: ${fs_sm}px; }
#presetBtn {
    padding: 5px 14px;
    border: 1px solid $border_strong;
    border-radius: 6px;
    background-color: $surface;
    color: $text_body;
    font-size: ${fs_sm}px;
    font-weight: $w_medium;
}
#presetBtn:hover { background-color: $surface_hover; border-color: $primary_soft_border; color: $text; }
#presetBtn:checked { background-color: $primary_soft; border-color: $primary_soft_border; color: $primary_text; font-weight: $w_bold; }
#presetBtn:focus { outline: none; }
#presetBtn:disabled { background-color: $surface_alt; color: $text_faint; border-color: $border; }

QProgressBar#thinProgressBar {
    border: none;
    background-color: $border_subtle;
    max-height: 3px;
    min-height: 3px;
    margin: 0;
    padding: 0;
}
QProgressBar#thinProgressBar::chunk {
    background-color: $primary;
    border-radius: 1px;
}

#actionBtn {
    padding: 7px 14px;
    border: 1px solid $border_strong;
    border-radius: 6px;
    background-color: $surface;
    color: $text_body;
    font-weight: $w_medium;
}
#actionBtn:hover { background-color: $surface_hover; border-color: $primary_soft_border; color: $text; }
#actionBtn:pressed { background-color: $surface_alt; }
#actionBtn:disabled { background-color: $surface_alt; color: $text_faint; border-color: $border; }

#footerSummary { color: $text; font-weight: $w_medium; font-size: ${fs_sm}px; }
#processingHint { color: $primary_text; font-weight: $w_medium; font-size: ${fs_sm}px; min-width: 0px; }
#footerHint { color: $text_muted; font-size: ${fs_sm}px; min-width: 0px; }
#footerHint[danger="true"] { color: $danger_text; font-weight: $w_medium; }

#startBtn {
    padding: 8px 20px;
    background-color: $primary;
    color: $text_on_primary;
    border-radius: 6px;
    font-weight: $w_bold;
    border: none;
}
#startBtn[stopMode="true"] {
    background-color: $danger_soft;
    color: $danger_text;
    border: 1px solid $danger;
}
#startBtn[stopMode="true"]:hover { background-color: $danger_hover; color: $text_on_primary; }
#startBtn[stopMode="true"]:pressed { background-color: $danger_pressed; color: $text_on_primary; }
#startBtn:hover { background-color: $primary_hover; }
#startBtn:pressed { background-color: $primary_pressed; }
#startBtn:disabled { background-color: $primary_disabled; color: $surface; border: none; }

/* ============================ 对话框 ============================ */
#dialogBg { background-color: $surface; border: 1px solid $border; border-radius: 8px; }
#dialogTitleBar { background-color: $surface_alt; border: none; border-bottom: 1px solid $border; border-top-left-radius: 7px; border-top-right-radius: 7px; }
#dialogTitle { font-weight: $w_bold; color: $text; font-size: ${fs_base}px; border: none; }
#dialogCloseBtn { background: transparent; border: none; font-size: ${fs_md}px; color: $text_faint; border-radius: 4px; }
#dialogCloseBtn:hover { background-color: $danger_soft; color: $danger_text; }
#dialogContent { border: none; background-color: transparent; }

#dialogPrimaryBtn { background-color: $primary; color: $text_on_primary; border-radius: 6px; padding: 7px 16px; font-weight: $w_bold; border: none; }
#dialogPrimaryBtn:hover { background-color: $primary_hover; }
#dialogPrimaryBtn:pressed { background-color: $primary_pressed; }
#dialogPrimaryBtn:disabled { background-color: $primary_disabled; color: $surface; }
#dialogDangerBtn { background-color: $danger; color: $text_on_primary; border-radius: 6px; padding: 7px 16px; font-weight: $w_bold; border: none; }
#dialogDangerBtn:hover { background-color: $danger_hover; }
#dialogDangerBtn:pressed { background-color: $danger_pressed; }
#dialogSecondaryBtn { background-color: $surface; color: $text_body; border-radius: 6px; padding: 6px 14px; font-weight: $w_medium; border: 1px solid $border_strong; }
#dialogSecondaryBtn:hover { background-color: $surface_hover; border-color: $primary_soft_border; color: $text; }
#dialogSecondaryBtn:pressed { background-color: $surface_alt; }
#dialogSecondaryBtn:checked { background-color: $primary_soft; border-color: $primary_soft_border; color: $primary_text; }
#dialogSecondaryBtn:disabled { background-color: $surface_alt; color: $text_faint; border-color: $border; }
#dialogSectionTitle { color: $text_muted; font-size: ${fs_sm}px; font-weight: $w_medium; border: none; }

/* ---- 语义状态横幅（PDF 深度诊断等对话框使用） ---- */
#statusBadge {
    font-size: ${fs_base}px;
    font-weight: $w_bold;
    padding: 6px 12px;
    border-radius: 6px;
    border: 1px solid $border;
    background-color: $surface_alt;
    color: $text_body;
}
#statusBadge[variant="success"] { background-color: $success_soft; color: $success_text; border-color: $success; }
#statusBadge[variant="warning"] { background-color: $warning_soft; color: $warning_text; border-color: $warning; }
#statusBadge[variant="danger"] { background-color: $danger_soft; color: $danger_text; border-color: $danger_soft_border; }
#statusBadge[variant="info"] { background-color: $info_soft; color: $info_text; border-color: $info; }

/* ---- 选项卡 (QTabWidget / QTabBar) ---- */
QTabWidget { background: transparent; border: none; }
QTabWidget::pane { border: none; background: transparent; margin-top: 6px; }
QTabBar { background: transparent; border: none; qproperty-drawBase: 0; }
QTabBar::tab {
    background-color: $surface_alt;
    color: $text_muted;
    border: 1px solid $border;
    border-radius: 6px;
    padding: 7px 16px;
    margin-right: 8px;
    font-weight: $w_medium;
    font-size: ${fs_base}px;
}
QTabBar::tab:hover { background-color: $surface_hover; color: $text; border-color: $border_strong; }
QTabBar::tab:selected {
    background-color: $primary_soft;
    color: $primary_text;
    border: 1px solid $primary_soft_border;
    font-weight: $w_bold;
}

/* ---- 输入框 / 下拉选择框 / 数字框 / 多行文本 ---- */
QLineEdit {
    background-color: $surface;
    border: 1px solid $border_strong;
    border-radius: 6px;
    padding: 6px 10px;
    color: $text_body;
    selection-background-color: $selection_bg;
    selection-color: $selection_text;
}
QLineEdit:focus { border-color: $primary; }
QLineEdit:disabled { background-color: $surface_alt; color: $text_faint; border-color: $border; }

QComboBox {
    background-color: $surface;
    border: 1px solid $border_strong;
    border-radius: 6px;
    padding: 2px 26px 2px 10px;
    min-height: 28px;
    color: $text_body;
}
QComboBox:hover { border-color: $primary_soft_border; background-color: $surface_hover; }
QComboBox:focus { border-color: $primary; }
QComboBox:disabled { background-color: $surface_alt; color: $text_faint; border-color: $border; }
QComboBox::drop-down {
    subcontrol-origin: padding;
    subcontrol-position: center right;
    width: 24px;
    border-left: none;
}
QComboBox::down-arrow {
    image: url($chevron_down_url);
    width: 14px;
    height: 14px;
}
QComboBox QAbstractItemView {
    background-color: $surface;
    border: 1px solid $border;
    border-radius: 8px;
    color: $text_body;
    selection-background-color: $primary_soft;
    selection-color: $primary_text;
    padding: 4px;
    outline: none;
}

QSpinBox {
    background-color: $surface;
    border: 1px solid $border_strong;
    border-radius: 6px;
    padding: 3px 26px 3px 8px;
    min-height: 20px;
    color: $text_body;
    selection-background-color: $selection_bg;
    selection-color: $selection_text;
}
QSpinBox:focus { border-color: $primary; }
QSpinBox:disabled { background-color: $surface_alt; border-color: $border; color: $text_faint; }
QSpinBox::up-button, QSpinBox::down-button {
    subcontrol-origin: border;
    width: 22px;
    background: transparent;
    border: none;
    border-left: 1px solid $border_subtle;
}
QSpinBox::up-button { border-top-right-radius: 5px; }
QSpinBox::down-button { border-bottom-right-radius: 5px; border-top: 1px solid $border_subtle; }
QSpinBox::up-button:hover, QSpinBox::down-button:hover { background-color: $surface_hover; }
QSpinBox::up-button:pressed, QSpinBox::down-button:pressed { background-color: $primary_soft; }
QSpinBox::up-arrow { image: url($chevron_up_url); width: 10px; height: 10px; }
QSpinBox::down-arrow { image: url($chevron_down_url); width: 10px; height: 10px; }

QPlainTextEdit, QTextEdit {
    background-color: $surface_alt;
    border: 1px solid $border;
    border-radius: 8px;
    padding: 8px;
    color: $text_body;
    selection-background-color: $selection_bg;
    selection-color: $selection_text;
}

/* ---- 对话框内文本 ---- */
#dialogHeading { color: $text; font-size: ${fs_xl}px; font-weight: $w_bold; border: none; }
#dialogSubTitle { color: $text; font-size: ${fs_base}px; font-weight: $w_bold; border: none; }
#dialogBody { color: $text_body; font-size: ${fs_base}px; border: none; }
#dialogMuted { color: $text_muted; font-size: ${fs_sm}px; border: none; }
#dialogCaption { color: $text_body; font-size: ${fs_sm}px; font-weight: $w_medium; border: none; }
#dialogCodeBlock { color: $text_body; font-size: ${fs_sm}px; border: 1px solid $border; background-color: $surface_alt; border-radius: 8px; padding: 10px; font-family: $font_mono; }
#dialogStatus { color: $text_muted; font-size: ${fs_sm}px; border: none; }
#dialogStatus[state="muted"] { color: $text_muted; }
#dialogStatus[state="info"] { color: $primary_text; }
#dialogStatus[state="success"] { color: $success_text; }
#dialogStatus[state="error"] { color: $danger_text; }
#dialogInfoPanel { color: $text_body; background-color: $surface_alt; border: 1px solid $border; border-radius: 8px; padding: 7px 10px; font-size: ${fs_sm}px; }
#dialogEmptyPanel { color: $text_muted; background-color: $surface_alt; border: 1px dashed $border_strong; border-radius: 8px; padding: 18px; font-size: ${fs_sm}px; }

/* ---- 已签名文件提示对话框 ---- */
#signedWarnCard { background-color: $warning_soft; border: 1px solid $warning; border-radius: 8px; }
#signedWarnCard QLabel { border: none; background: transparent; }
#signedWarnIcon { font-size: 22px; border: none; background: transparent; }
#signedWarnTitle { color: $warning_text; font-size: ${fs_md}px; font-weight: $w_bold; border: none; }
#signedWarnDesc { color: $text_body; font-size: ${fs_sm}px; border: none; }
#signedListCaption { color: $text_muted; font-size: ${fs_sm}px; font-weight: $w_medium; border: none; }
#signedListScroll { background-color: $surface_alt; border: 1px solid $border; border-radius: 8px; }
#signedListScroll > QWidget { background: transparent; }
#signedFileList { color: $text_body; font-size: ${fs_sm}px; background: transparent; border: none; padding: 10px 12px; line-height: 150%; }

/* ---- 消息框 ---- */
#msgIcon { border: none; background: transparent; }
#msgIconBadge { border-radius: 8px; border: 1px solid transparent; }
#msgIconBadge[variant="info"] { background-color: $info_soft; border-color: $primary_soft_border; }
#msgIconBadge[variant="success"] { background-color: $success_soft; border-color: $success; }
#msgIconBadge[variant="warning"] { background-color: $warning_soft; border-color: $warning; }
#msgIconBadge[variant="error"] { background-color: $danger_soft; border-color: $danger_soft_border; }
#msgIconBadge[variant="question"] { background-color: $primary_soft; border-color: $primary_soft_border; }
#msgText { color: $text_body; font-size: ${fs_base}px; border: none; background: transparent; }
#msgTextBlock { color: $text_body; font-size: ${fs_sm}px; border: 1px solid $border; background-color: $surface_alt; border-radius: 8px; padding: 10px; font-family: $font_mono; }

/* ---- 数据向导选择按钮 ---- */
#choiceToggleBtn { color: $text_body; background-color: $surface_alt; border: 1px solid $border_strong; border-radius: 6px; padding: 8px 14px; font-weight: $w_medium; text-align: center; }
#choiceToggleBtn:hover { background-color: $surface_hover; }
#choiceToggleBtn:checked { color: $primary_text; background-color: $primary_soft; border-color: $primary_soft_border; font-weight: $w_bold; }

/* ---- 数据向导目录卡片 / 预览表 ---- */
#wizardCard { background-color: $surface_alt; border: 1px solid $border; border-radius: 8px; }
#wizardCard QLabel, #wizardCard QCheckBox, #wizardCard QRadioButton { border: none; background: transparent; }
#compressionCard { background-color: $surface_alt; border: 1px solid $border; border-radius: 8px; }
#compressionCard QLabel { border: none; background: transparent; }
#compressionCardTitle { font-weight: $w_bold; font-size: ${fs_base}px; color: $primary_text; border: none; background: transparent; }
#compressionCardSection { font-weight: $w_medium; font-size: ${fs_sm}px; color: $text; border: none; background: transparent; }
#compressionCardHint { color: $text_muted; font-size: ${fs_sm}px; border: none; background: transparent; }
#compressionCardSep { color: $border; background-color: $border; height: 1px; border: none; }
#previewTable { background-color: $surface; border: 1px solid $border; border-radius: 8px; gridline-color: $border; color: $text_body; selection-background-color: $primary_soft; selection-color: $primary_text; }
#previewTable QHeaderView::section { background-color: $surface_alt; border: none; border-bottom: 1px solid $border; padding: 7px; color: $text_muted; font-weight: $w_medium; }
#previewTable QTableCornerButton::section { background-color: $surface_alt; border: none; border-bottom: 1px solid $border; }

/* ---- 设置对话框 ---- */
#settingsPathCard { background-color: $surface_alt; border: 1px solid $border; border-radius: 8px; }
#settingsPathCard QLabel { background: transparent; border: none; }
#settingsPathHint { color: $text_muted; font-size: ${fs_sm}px; border: none; }
#settingsPathEdit { background-color: $surface; border: 1px solid $border_strong; border-radius: 6px; padding: 8px 10px; color: $text_body; selection-background-color: $selection_bg; selection-color: $selection_text; }
#settingsPathEdit:focus { border-color: $primary; }
#settingsPathStatus { font-size: ${fs_sm}px; border: none; padding: 0 2px; }
#settingsPathStatus[state="empty"] { color: $text_faint; }
#settingsPathStatus[state="valid"] { color: $primary_text; }
#settingsPathStatus[state="invalid"] { color: $danger; font-weight: $w_medium; }
#settingsWorkerSpin { background-color: $surface; border: 1px solid $border_strong; border-right: none; border-top-left-radius: 6px; border-bottom-left-radius: 6px; border-top-right-radius: 0; border-bottom-right-radius: 0; padding: 3px 10px; color: $text_body; min-height: 24px; min-width: 78px; }
#settingsWorkerSpin:focus { border-color: $primary; }
#settingsWorkerSpin:disabled { background-color: $surface_alt; border-color: $border; color: $text_muted; }
#settingsWorkerSpin::up-button, #settingsWorkerSpin::down-button { subcontrol-origin: border; width: 0px; height: 0px; border: none; background: transparent; }
#settingsWorkerSpin::up-arrow, #settingsWorkerSpin::down-arrow { image: none; width: 0px; height: 0px; }
#settingsSpinStepper { background-color: $surface_alt; border: 1px solid $border_strong; border-left: none; border-top-right-radius: 6px; border-bottom-right-radius: 6px; min-width: 24px; max-width: 24px; }
#settingsSpinStepper:disabled { border-color: $border; }
#settingsSpinStepBtn { background: transparent; border: none; color: $text_muted; font-size: 9px; font-weight: $w_bold; padding: 0; min-width: 23px; max-width: 23px; min-height: 15px; max-height: 15px; }
#settingsSpinStepBtn:hover { background-color: $surface_hover; color: $text; }
#settingsSpinStepBtn:pressed { background-color: $primary_soft; color: $primary_text; }
#settingsSpinStepBtn:disabled { color: $text_faint; background: transparent; }
#settingsSpinStepBtn[direction="up"] { border-top-right-radius: 5px; border-bottom: 1px solid $border_subtle; }
#settingsSpinStepBtn[direction="down"] { border-bottom-right-radius: 5px; }

/* ---- 主题选择分段按钮 ---- */
#themeSegBtn { color: $text_body; background-color: $surface; border: 1px solid $border_strong; padding: 7px 14px; font-weight: $w_medium; border-radius: 6px; }
#themeSegBtn:hover { background-color: $surface_hover; }
#themeSegBtn:checked { color: $primary_text; background-color: $primary_soft; border-color: $primary_soft_border; font-weight: $w_bold; }

/* ---- 日志对话框 ---- */
#statCard { background-color: $surface_alt; border: 1px solid $border; border-radius: 8px; }
#statCard:hover { background-color: $surface_hover; border-color: $border_strong; }
#statCard[active="true"] { background-color: $primary_soft; border-color: $primary_soft_border; }
#statCard[accent="success"] { border-left: 3px solid $success; }
#statCard[accent="failure"] { border-left: 3px solid $danger; }
#statCard[accent="precheck"] { border-left: 3px solid $info; }
#statCard[accent="skip"] { border-left: 3px solid $warning; }
#statCard[accent="all"] { border-left: 3px solid $primary; }
#statCard QLabel { border: none; background: transparent; padding: 0; }
#statCardTitle { color: $text_muted; font-size: ${fs_sm}px; font-weight: $w_medium; }
#statCardValue { color: $text; font-size: ${fs_2xl}px; font-weight: $w_bold; }
#logSummaryFrame { background-color: $surface; border: 1px solid $border; border-radius: 8px; }
#logSummaryTable { background-color: $surface; alternate-background-color: $surface_alt2; color: $text_body; border: none; border-radius: 0; gridline-color: $border; selection-background-color: $primary_soft; selection-color: $primary_text; outline: none; }
#logSummaryTable::item { padding: 7px; border-bottom: 1px solid $border_subtle; }
#logSummaryTable::item:selected { background-color: $primary_soft; color: $primary_text; }
#logSummaryTable QHeaderView::section { background-color: $surface_alt; border: none; border-bottom: 1px solid $border; padding: 8px; color: $text_muted; font-weight: $w_medium; }
#logSummaryTable QTableCornerButton::section { background-color: $surface_alt; border: none; border-bottom: 1px solid $border; }
#logSplitter::handle:vertical { background-color: transparent; border: none; height: 10px; margin: 2px 0; }
#logSplitter::handle:vertical:hover { background-color: $primary_soft; border-radius: 4px; }
#logDetailTextEdit { background-color: $surface_alt; border: 1px solid $border; border-radius: 8px; padding: 10px; color: $text_body; font-family: $font_mono; font-size: ${fs_sm}px; selection-background-color: $selection_bg; selection-color: $selection_text; }

/* ---- 右键菜单 ---- */
QMenu { background-color: $surface; border: 1px solid $border_strong; border-radius: 8px; padding: 4px; }
QMenu::item { padding: 6px 24px 6px 12px; border-radius: 4px; color: $text_body; font-size: ${fs_base}px; }
QMenu::item:selected { background-color: $primary_soft; color: $primary_text; }
QMenu::separator { height: 1px; background: $border; margin: 4px 8px; }
QMenu::item:disabled { color: $text_faint; }

/* ---- 关于对话框 ---- */
#aboutHeroCard { background-color: $primary_soft; border: 1px solid $primary_soft_border; border-radius: 8px; }
#aboutInfoCard { background-color: $surface_alt; border: 1px solid $border; border-radius: 8px; }
#aboutBrandTitle { font-size: ${fs_2xl}px; font-weight: $w_bold; color: $primary_text; border: none; }
#aboutBadge { background-color: $surface; color: $primary_text; border: 1px solid $primary_soft_border; border-radius: 6px; padding: 4px 10px; font-weight: $w_medium; font-size: ${fs_sm}px; }
#aboutTitle { color: $text; font-size: ${fs_base}px; font-weight: $w_bold; border: none; }
#aboutText { color: $text_body; font-size: ${fs_sm}px; border: none; }
#aboutIntro { color: $text_body; font-size: ${fs_base}px; border: none; }

/* ---- 去除水印对话框 ---- */
#watermarkSplitter::handle:horizontal { background-color: transparent; border: none; width: 10px; margin: 0 2px; }
#watermarkSplitter::handle:horizontal:hover { background-color: $primary_soft; border-radius: 4px; }
#watermarkPanelCard { background-color: $surface; border: 1px solid $border; border-radius: 8px; }
#watermarkPanelHeader {
    background-color: $surface_alt;
    border: none;
    border-bottom: 1px solid $border;
    border-top-left-radius: 7px;
    border-top-right-radius: 7px;
}
#watermarkPanelHeader QLabel { background: transparent; border: none; }
#watermarkFilterBar {
    background-color: $surface_alt;
    border: none;
    border-bottom: 1px solid $border;
}
#watermarkFilterBar QLineEdit { min-height: 16px; padding: 4px 8px; }
#watermarkFilterBar QComboBox { min-height: 24px; padding: 2px 24px 2px 8px; }
#watermarkPreviewPane { background: transparent; border: none; }
#watermarkPreviewTitle { color: $text; font-size: ${fs_sm}px; font-weight: $w_bold; border: none; background: transparent; }
#watermarkPreviewTitle[accent="after"] { color: $primary_text; }
#watermarkPreviewScroll { background-color: $surface_alt; border: 1px solid $border; border-radius: 8px; }
#watermarkPreviewScroll > QWidget > QWidget { background-color: $surface_alt; }
#watermarkPreviewCanvas { color: $text_muted; font-size: ${fs_sm}px; background-color: transparent; border: none; padding: 12px; }
#watermarkNotesPanel { color: $warning_text; background-color: $warning_soft; border: 1px solid $warning; border-radius: 8px; padding: 8px 12px; font-size: ${fs_sm}px; }
#watermarkConfirmCard { background-color: $surface_alt; border: 1px solid $border; border-radius: 8px; }
#watermarkConfirmCard[checked="true"] { background-color: $primary_soft; border-color: $primary_soft_border; }
#watermarkConfirmCard QCheckBox { background: transparent; border: none; }

/* ---- PDF 加密与解密对话框 ---- */
QTabWidget#securityTabs { background-color: $surface; border: none; }
QTabWidget#securityTabs::pane { border: none; background-color: $surface; margin-top: 10px; }
QTabWidget#securityTabs QTabBar { background-color: $surface; border: none; qproperty-drawBase: 0; outline: none; }
QTabWidget#securityTabs QTabBar::tab {
    background: transparent;
    color: $text_muted;
    border: none;
    border-bottom: 2px solid transparent;
    border-radius: 0px;
    padding: 8px 16px;
    margin: 0px 6px 0px 1px;
    font-weight: $w_medium;
    font-size: ${fs_base}px;
    outline: none;
}
QTabWidget#securityTabs QTabBar::tab:focus { outline: none; }
QTabWidget#securityTabs QTabBar::tab:hover {
    color: $text;
    background-color: $surface_hover;
    border-top-left-radius: 6px;
    border-top-right-radius: 6px;
}
QTabWidget#securityTabs QTabBar::tab:selected {
    color: $primary_text;
    border-bottom: 2px solid $primary;
    font-weight: $w_bold;
    background: transparent;
}
#securitySectionCard {
    background-color: $surface_alt;
    border: none;
    border-radius: 8px;
}
#securitySectionCard QLabel,
#securitySectionCard QCheckBox,
#securitySectionCard QRadioButton {
    border: none;
    background: transparent;
}
#securityTableCard {
    background-color: $surface;
    border: 1px solid $border;
    border-radius: 8px;
}
#securityTableCard[dragActive="true"] {
    border: 2px dashed $primary;
    background-color: $primary_soft;
}
#securityTableHeader {
    background-color: $surface_alt;
    border: none;
    border-bottom: 1px solid $border;
    border-top-left-radius: 7px;
    border-top-right-radius: 7px;
}
#securityTableHeader QLabel { background: transparent; border: none; }
#securityTable {
    background-color: $surface;
    alternate-background-color: $surface_alt2;
    color: $text_body;
    border: none;
    border-bottom-left-radius: 7px;
    border-bottom-right-radius: 7px;
    selection-background-color: $primary_soft;
    selection-color: $primary_text;
    outline: none;
}
#securityTable::item { padding: 6px 8px; border-bottom: 1px solid $border_subtle; }
#securityTable::item:selected { background-color: $primary_soft; color: $primary_text; }
#securityTable QHeaderView::section {
    background-color: $surface_alt;
    border: none;
    border-bottom: 1px solid $border;
    padding: 7px 8px;
    color: $text_muted;
    font-size: ${fs_sm}px;
    font-weight: $w_medium;
    text-align: left;
}
#securityTable QTableCornerButton::section {
    background-color: $surface_alt;
    border: none;
    border-bottom: 1px solid $border;
}
#securityFooter {
    background: transparent;
    border: none;
    border-top: 1px solid $border;
}
""")


def build_app_qss(palette: Palette, typography: Typography | None = None) -> str:
    """依据 Palette 与 Typography 渲染整份应用级 QSS。"""
    tokens = palette.as_dict()
    typo = typography or active_typography()
    tokens.update(typo.as_qss_tokens())
    tokens["checkmark_url"] = checkmark_icon_path(palette.text_on_primary)
    tokens["chevron_down_url"] = icons.svg_temp_file_url(
        "chevron_down", palette.text_muted, stroke_width=2.0
    )
    tokens["chevron_up_url"] = icons.svg_temp_file_url(
        "chevron_up", palette.text_muted, stroke_width=2.0
    )
    tokens["radio_dot_url"] = icons.svg_temp_file_url(
        "radio_dot", palette.text_on_primary, stroke_width=1.0
    )
    return _APP_QSS_TEMPLATE.substitute(tokens)


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


def tree_status_color(palette: Palette, semantic: str) -> str:
    """文件树"当前状态"列的前景色。semantic 取值见 common.status.status_semantic。"""
    if semantic == "positive":
        return palette.success
    if semantic == "negative":
        return palette.danger
    if semantic == "warning":
        return palette.warning
    return palette.info


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

    def __init__(
        self,
        app: QApplication,
        mode: str = "system",
        font_scale: str = "standard",
        parent: QObject | None = None,
    ):
        super().__init__(parent if parent is not None else app)
        self._app = app
        self._mode = mode if mode in self._MODES else "system"
        self._font_scale = font_scale if font_scale in FONT_SCALES else "standard"
        self._palette = LIGHT
        self._typography = resolve_typography(self._font_scale)
        self._applying = False

        hints = app.styleHints()
        # 系统主题变化时, 若处于跟随模式则实时刷新
        hints.colorSchemeChanged.connect(self._on_system_scheme_changed)

    # --- 查询 ---
    @property
    def mode(self) -> str:
        return self._mode

    @property
    def font_scale(self) -> str:
        return self._font_scale

    def current_palette(self) -> Palette:
        return self._palette

    def current_typography(self) -> Typography:
        return self._typography

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
        global _ACTIVE_PALETTE, _ACTIVE_TYPOGRAPHY
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
            self._typography = resolve_typography(self._font_scale)
            _ACTIVE_TYPOGRAPHY = self._typography
        finally:
            self._applying = False
        if hasattr(self._app, "setFont"):
            self._app.setFont(build_app_qfont(self._typography))
        self._app.setStyleSheet(build_app_qss(self._palette, self._typography))
        self.changed.emit(self._palette)

    def set_mode(self, mode: str):
        if mode not in self._MODES:
            return
        self._mode = mode
        self.apply()

    def set_font_scale(self, scale: str):
        if scale not in FONT_SCALES:
            return
        self._font_scale = scale
        self.apply()

    def _on_system_scheme_changed(self, _scheme):
        if self._applying:
            return
        if self._mode == "system":
            self.apply()
