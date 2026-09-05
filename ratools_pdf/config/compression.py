# -*- coding: utf-8 -*-
"""图像压缩参数的唯一定义处。

UI 层负责收集参数（settings.ini + DPI 对话框），pdf 层只消费归一化后的
结果；两侧的默认值与取值范围都从这里取，避免双份维护。
"""

DEFAULT_IMAGE_DPI = 300
DEFAULT_JPEG_QUALITY = 85
MIN_IMAGE_DPI = 72
MAX_IMAGE_DPI = 600
MIN_JPEG_QUALITY = 1
MAX_JPEG_QUALITY = 95


def normalize_compression_settings(settings):
    """校验并归一化外部传入的压缩参数，非法/缺失值回落默认。

    返回 {"dpi": int, "quality": int}，保证 pdf 层拿到的值永远可用；
    settings 可能为 None（旧调用方或测试未传参）。
    """
    if not isinstance(settings, dict):
        settings = {}
    return {
        "dpi": _to_int_in_range(settings.get("dpi"), DEFAULT_IMAGE_DPI, MIN_IMAGE_DPI, MAX_IMAGE_DPI),
        "quality": _to_int_in_range(settings.get("quality"), DEFAULT_JPEG_QUALITY, MIN_JPEG_QUALITY, MAX_JPEG_QUALITY),
    }


def _to_int_in_range(value, default, minimum, maximum):
    try:
        value = int(value)
    except (TypeError, ValueError):
        return default
    return max(minimum, min(maximum, value))
