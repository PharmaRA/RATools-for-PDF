"""内联线性 SVG 图标库与主题动态着色渲染器。

将全量 UI 图标以精简 SVG 路径直接内联在代码中：
1. 零外部资源文件清单负担（PyInstaller / Nuitka 无需维护额外图标文件列表）。
2. 纯数据表（``SVG_PATHS`` / ``MODULE_ICON_MAP`` / ``build_svg_xml``）不依赖 Qt，
   可在无头纯 Python 单元测试中直接校验完整性。
3. Qt 渲染方法延迟导入 PySide6，按当前主题 ``Palette`` 动态着色并缓存 ``QPixmap`` / ``QIcon``。
"""

from __future__ import annotations

import os
import tempfile
from typing import Dict

# 统一使用 24x24 viewBox 的 Lucide / Fluent 线性图标片段
SVG_PATHS: Dict[str, str] = {
    "app_logo": (
        '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>'
        '<polyline points="14 2 14 8 20 8"/>'
        '<path d="M9 15h6"/>'
        '<path d="M9 11h3"/>'
    ),
    "eye": (
        '<path d="M2 12s3.6-7 10-7 10 7 10 7-3.6 7-10 7-10-7-10-7Z"/>'
        '<circle cx="12" cy="12" r="3"/>'
    ),
    "file_text": (
        '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>'
        '<polyline points="14 2 14 8 20 8"/>'
        '<line x1="16" y1="13" x2="8" y2="13"/>'
        '<line x1="16" y1="17" x2="8" y2="17"/>'
        '<line x1="10" y1="9" x2="8" y2="9"/>'
    ),
    "bookmark": '<path d="m19 21-7-4-7 4V5a2 2 0 0 1 2-2h10a2 2 0 0 1 2 2v16z"/>',
    "link": (
        '<path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71"/>'
        '<path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71"/>'
    ),
    "shield": (
        '<path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/>'
        '<path d="m9 12 2 2 4-4"/>'
    ),
    "package": (
        '<path d="m16.5 9.4-9-5.19"/>'
        '<path d="M21 16V8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16z"/>'
        '<polyline points="3.29 7 12 12 20.71 7"/>'
        '<line x1="12" y1="22" x2="12" y2="12"/>'
    ),
    "search": (
        '<circle cx="11" cy="11" r="8"/>'
        '<line x1="21" y1="21" x2="16.65" y2="16.65"/>'
    ),
    "settings": (
        '<circle cx="12" cy="12" r="3"/>'
        '<path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1-2.83 2.83l-.06-.06a1.65 1.65 0 0 0-1.82-.33 '
        '1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-4 0v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06'
        'a2 2 0 0 1-2.83-2.83l.06-.06a1.65 1.65 0 0 0 .33-1.82 1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1 0-4h.09'
        'A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 2.83-2.83l.06.06a1.65 1.65 0 0 0 1.82.33'
        'H9a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 4 0v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06'
        'a2 2 0 0 1 2.83 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82V9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 0 4h-.09'
        'a1.65 1.65 0 0 0-1.51 1z"/>'
    ),
    "lock": (
        '<rect x="3" y="11" width="18" height="11" rx="2" ry="2"/>'
        '<path d="M7 11V7a5 5 0 0 1 10 0v4"/>'
    ),
    "info": (
        '<circle cx="12" cy="12" r="10"/>'
        '<line x1="12" y1="16" x2="12" y2="12"/>'
        '<line x1="12" y1="8" x2="12.01" y2="8"/>'
    ),
    "sun": (
        '<circle cx="12" cy="12" r="4"/>'
        '<path d="M12 2v2"/>'
        '<path d="M12 20v2"/>'
        '<path d="m4.93 4.93 1.41 1.41"/>'
        '<path d="m17.66 17.66 1.41 1.41"/>'
        '<path d="M2 12h2"/>'
        '<path d="M20 12h2"/>'
        '<path d="m6.34 17.66-1.41 1.41"/>'
        '<path d="m19.07 4.93-1.41 1.41"/>'
    ),
    "moon": '<path d="M12 3a6 6 0 0 0 9 9 9 9 0 1 1-9-9Z"/>',
    "upload": (
        '<path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>'
        '<polyline points="17 8 12 3 7 8"/>'
        '<line x1="12" y1="3" x2="12" y2="15"/>'
    ),
    "plus": (
        '<line x1="12" y1="5" x2="12" y2="19"/>'
        '<line x1="5" y1="12" x2="19" y2="12"/>'
    ),
    "folder": '<path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"/>',
    "folder_open": (
        '<path d="m6 14 1.5-2.9A2 2 0 0 1 9.24 10H20a2 2 0 0 1 1.94 2.5l-1.54 6a2 2 0 0 1-1.95 1.5'
        'H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h3.9a2 2 0 0 1 1.69.9l.81 1.2a2 2 0 0 0 1.67.9H18a2 2 0 0 1 2 2v2"/>'
    ),
    "drive": (
        '<line x1="22" y1="12" x2="2" y2="12"/>'
        '<path d="M5.45 5.11 2 12v6a2 2 0 0 0 2 2h16a2 2 0 0 0 2-2v-6l-3.45-6.89A2 2 0 0 0 16.76 4H7.24a2 2 0 0 0-1.79 1.11z"/>'
        '<line x1="6" y1="16" x2="6.01" y2="16"/>'
        '<line x1="10" y1="16" x2="10.01" y2="16"/>'
    ),
    "file": (
        '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>'
        '<polyline points="14 2 14 8 20 8"/>'
    ),
    "trash": (
        '<polyline points="3 6 5 6 21 6"/>'
        '<path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/>'
    ),
    "scan": (
        '<path d="M3 7V5a2 2 0 0 1 2-2h2"/>'
        '<path d="M17 3h2a2 2 0 0 1 2 2v2"/>'
        '<path d="M21 17v2a2 2 0 0 1-2 2h-2"/>'
        '<path d="M7 21H5a2 2 0 0 1-2-2v-2"/>'
        '<line x1="7" y1="12" x2="17" y2="12"/>'
    ),
    "list_log": (
        '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>'
        '<polyline points="14 2 14 8 20 8"/>'
        '<line x1="8" y1="13" x2="16" y2="13"/>'
        '<line x1="8" y1="17" x2="13" y2="17"/>'
    ),
    "play": '<polygon points="6 3 20 12 6 21 6 3"/>',
    "square": '<rect x="5" y="5" width="14" height="14" rx="2"/>',
    "skip": (
        '<polygon points="5 4 15 12 5 20 5 4"/>'
        '<line x1="19" y1="5" x2="19" y2="19"/>'
    ),
    "refresh": (
        '<polyline points="23 4 23 10 17 10"/>'
        '<path d="M20.49 15a9 9 0 1 1-2.12-9.36L23 10"/>'
    ),
    "check": '<polyline points="20 6 9 17 4 12"/>',
    "sliders": (
        '<line x1="4" y1="21" x2="4" y2="14"/>'
        '<line x1="4" y1="10" x2="4" y2="3"/>'
        '<line x1="12" y1="21" x2="12" y2="12"/>'
        '<line x1="12" y1="8" x2="12" y2="3"/>'
        '<line x1="20" y1="21" x2="20" y2="16"/>'
        '<line x1="20" y1="12" x2="20" y2="3"/>'
        '<line x1="1" y1="14" x2="7" y2="14"/>'
        '<line x1="9" y1="8" x2="15" y2="8"/>'
        '<line x1="17" y1="16" x2="23" y2="16"/>'
    ),
    "chevron_down": '<polyline points="6 9 12 15 18 9"/>',
    "chevron_up": '<polyline points="18 15 12 9 6 15"/>',
    "wrench": (
        '<path d="M14.7 6.3a1 1 0 0 0 0 1.4l1.6 1.6a1 1 0 0 0 1.4 0l3.77-3.77'
        'a6 6 0 0 1-7.94 7.94l-6.91 6.91a2.12 2.12 0 0 1-3-3l6.91-6.91a6 6 0 0 1 7.94-7.94l-3.76 3.76z"/>'
    ),
    "arrow_up_down": (
        '<path d="m7 15 5 5 5-5"/>'
        '<path d="m7 9 5-5 5 5"/>'
    ),
    "download": (
        '<path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>'
        '<polyline points="7 10 12 15 17 10"/>'
        '<line x1="12" y1="15" x2="12" y2="3"/>'
    ),
    "check_circle": (
        '<circle cx="12" cy="12" r="10"/>'
        '<path d="m9 12 2 2 4-4"/>'
    ),
    "alert_triangle": (
        '<path d="m21.73 18-8-14a2 2 0 0 0-3.48 0l-8 14A2 2 0 0 0 4 21h16a2 2 0 0 0 1.73-3Z"/>'
        '<line x1="12" y1="9" x2="12" y2="13"/>'
        '<line x1="12" y1="17" x2="12.01" y2="17"/>'
    ),
    "x_circle": (
        '<circle cx="12" cy="12" r="10"/>'
        '<line x1="15" y1="9" x2="9" y2="15"/>'
        '<line x1="9" y1="9" x2="15" y2="15"/>'
    ),
    "help_circle": (
        '<circle cx="12" cy="12" r="10"/>'
        '<path d="M9.09 9a3 3 0 0 1 5.83 1c0 2-3 3-3 3"/>'
        '<line x1="12" y1="17" x2="12.01" y2="17"/>'
    ),
    "info_circle": (
        '<circle cx="12" cy="12" r="10"/>'
        '<line x1="12" y1="16" x2="12" y2="12"/>'
        '<line x1="12" y1="8" x2="12.01" y2="8"/>'
    ),
    "close": (
        '<line x1="18" y1="6" x2="6" y2="18"/>'
        '<line x1="6" y1="6" x2="18" y2="18"/>'
    ),
    "radio_dot": '<circle cx="12" cy="12" r="5" fill="{color}" stroke="none"/>',
    "locate": (
        '<circle cx="12" cy="12" r="3"/>'
        '<path d="M12 2v3"/>'
        '<path d="M12 19v3"/>'
        '<path d="M2 12h3"/>'
        '<path d="M19 12h3"/>'
    ),
    "eraser": (
        '<path d="m7 21-4.3-4.3c-1-1-1-2.5 0-3.4l9.6-9.6c1-1 2.5-1 3.4 0l5.6 5.6c1 1 1 2.5 0 3.4L13 21"/>'
        '<path d="M22 21H7"/>'
        '<path d="m5 11 9 9"/>'
    ),
}

# 模块标题 -> 图标键名映射（与 rules_catalog.MODULES 对齐）
MODULE_ICON_MAP: Dict[str, str] = {
    "初始视图与文档属性": "eye",
    "页面与字体标准化": "file_text",
    "书签管理": "bookmark",
    "超链接处理": "link",
    "内容合规与安全性": "shield",
    "文件级优化与输出": "package",
    "文档检测": "search",
}

_pixmap_cache: dict = {}
_icon_cache: dict = {}
_temp_svg_urls: dict = {}


def build_svg_xml(name: str, color: str, *, stroke_width: float = 1.75) -> str:
    """根据图标名与颜色构建完整的 SVG XML 字符串。"""
    body = SVG_PATHS[name]
    if "{color}" in body:
        body = body.format(color=color)
    return (
        f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" '
        f'fill="none" stroke="{color}" stroke-width="{stroke_width}" '
        f'stroke-linecap="round" stroke-linejoin="round">{body}</svg>'
    )


def svg_temp_file_url(name: str, color: str, *, stroke_width: float = 1.8) -> str:
    """将着色后的 SVG 写入临时目录并返回 QSS url 可用的正斜杠路径。"""
    key = (name, color, stroke_width)
    cached = _temp_svg_urls.get(key)
    if cached and os.path.exists(cached):
        return cached
    safe_color = color.lstrip("#").lower()
    file_path = os.path.join(
        tempfile.gettempdir(),
        f"ratools_pdf_icon_{name}_{safe_color}.svg",
    )
    try:
        with open(file_path, "w", encoding="utf-8") as fh:
            fh.write(build_svg_xml(name, color, stroke_width=stroke_width))
    except OSError:
        return ""
    normalized = file_path.replace("\\", "/")
    _temp_svg_urls[key] = normalized
    return normalized


def render_svg_pixmap(
    name: str,
    color: str,
    size: int = 16,
    dpr: float = 2.0,
    *,
    stroke_width: float = 1.75,
):
    """使用 QtSvg 将 SVG 渲染为支持高分屏的透明 QPixmap。"""
    from PySide6.QtCore import QByteArray, QRectF, Qt
    from PySide6.QtGui import QPainter, QPixmap
    from PySide6.QtSvg import QSvgRenderer

    key = (name, color, size, dpr, stroke_width)
    cached = _pixmap_cache.get(key)
    if cached is not None:
        return cached

    pixel_size = max(1, int(round(size * dpr)))
    pixmap = QPixmap(pixel_size, pixel_size)
    pixmap.fill(Qt.GlobalColor.transparent)

    svg_bytes = QByteArray(build_svg_xml(name, color, stroke_width=stroke_width).encode("utf-8"))
    renderer = QSvgRenderer(svg_bytes)
    painter = QPainter(pixmap)
    painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)
    renderer.render(painter, QRectF(0.0, 0.0, float(pixel_size), float(pixel_size)))
    painter.end()

    pixmap.setDevicePixelRatio(dpr)

    _pixmap_cache[key] = pixmap
    return pixmap


def svg_icon(
    name: str,
    color: str,
    size: int = 16,
    *,
    active_color: str | None = None,
    disabled_color: str | None = None,
    stroke_width: float = 1.75,
):
    """生成支持 Normal / Checked / Disabled 状态着色的 QIcon。"""
    from PySide6.QtGui import QIcon

    key = (name, color, size, active_color, disabled_color, stroke_width)
    cached = _icon_cache.get(key)
    if cached is not None:
        return cached

    icon = QIcon()
    normal_pm = render_svg_pixmap(name, color, size=size, stroke_width=stroke_width)
    icon.addPixmap(normal_pm, QIcon.Mode.Normal, QIcon.State.Off)
    if active_color:
        active_pm = render_svg_pixmap(name, active_color, size=size, stroke_width=stroke_width)
        icon.addPixmap(active_pm, QIcon.Mode.Normal, QIcon.State.On)
        icon.addPixmap(active_pm, QIcon.Mode.Active, QIcon.State.On)
    else:
        icon.addPixmap(normal_pm, QIcon.Mode.Normal, QIcon.State.On)

    if disabled_color:
        disabled_pm = render_svg_pixmap(name, disabled_color, size=size, stroke_width=stroke_width)
        icon.addPixmap(disabled_pm, QIcon.Mode.Disabled, QIcon.State.Off)
        icon.addPixmap(disabled_pm, QIcon.Mode.Disabled, QIcon.State.On)

    _icon_cache[key] = icon
    return icon
