"""Versioned, portable processing profiles. No UI, paths or personal settings."""

from copy import deepcopy
import hashlib
import json
import re

from ratools_pdf.config import rules_catalog
from ratools_pdf.config.compression import (
    DEFAULT_IMAGE_DPI, DEFAULT_JPEG_QUALITY, MAX_IMAGE_DPI, MAX_JPEG_QUALITY,
    MIN_IMAGE_DPI, MIN_JPEG_QUALITY,
)

SCHEMA_VERSION = 1
FIELDS = frozenset({
    "schema_version", "profile_id", "profile_version", "agency_or_region",
    "enabled_rules", "rule_parameters", "validation_level", "source", "applicability",
})
REQUIRED_FIELDS = FIELDS - {"source", "applicability"}
VALIDATION_LEVELS = {"structural"}


def validate_profile(payload):
    """Return a detached canonical profile, rejecting any unrecognized content."""
    if not isinstance(payload, dict) or set(payload) - FIELDS or REQUIRED_FIELDS - set(payload):
        raise ValueError("配置字段缺失或包含未知字段")
    if type(payload["schema_version"]) is not int or payload["schema_version"] != SCHEMA_VERSION:
        raise ValueError("不支持的配置 schema 版本")
    if not isinstance(payload["profile_id"], str) or not re.fullmatch(
        r"[A-Za-z0-9][A-Za-z0-9._-]{0,79}", payload["profile_id"]
    ):
        raise ValueError("无效的配置 ID")
    if not isinstance(payload["profile_version"], str) or not re.fullmatch(
        r"(?:0|[1-9][0-9]*)\.(?:0|[1-9][0-9]*)\.(?:0|[1-9][0-9]*)", payload["profile_version"]
    ):
        raise ValueError("配置版本必须为 major.minor.patch")
    for key in ("agency_or_region", "source", "applicability"):
        if key in payload and (not isinstance(payload[key], str) or not payload[key].strip()
                               or len(payload[key]) > 500 or any(ord(c) < 32 for c in payload[key])):
            raise ValueError("配置说明字段无效")
    if not isinstance(payload["validation_level"], str) or payload["validation_level"] not in VALIDATION_LEVELS:
        raise ValueError("不支持的验证级别（当前仅 structural）")
    rules = payload["enabled_rules"]
    if not isinstance(rules, list) or any(not isinstance(rule, str) for rule in rules):
        raise ValueError("规则必须为 ID 列表")
    selected = set(rules)
    if len(rules) != len(selected) or selected - set(rules_catalog.OPTION_TITLES):
        raise ValueError("规则重复或包含未知 ID")
    conflicts = (
        {"page_size_a4", "page_size_letter"},
        {"compress_standard", "compress_aggressive"},
        {"flatten_annotations", "cleanup_remove_annotations"},
        {"link_remove_border", "link_black_border", "link_bordered_to_blue_border",
         "link_unbordered_blue_to_blue_border"},
        {"cleanup_remove_external_uri", "cleanup_remove_external_uri_and_text_black"},
        {"cleanup_remove_invalid_links", "cleanup_remove_invalid_links_and_text_black"},
        {"cleanup_remove_metadata", "title_from_filename"},
    )
    for group in conflicts:
        # Blue-border conversions work on disjoint link populations and may coexist.
        active = selected & group
        if len(active) > 1 and active != {"link_bordered_to_blue_border", "link_unbordered_blue_to_blue_border"}:
            raise ValueError("存在互斥规则: " + ", ".join(sorted(active)))
    if "cleanup_remove_all_links_bookmarks" in selected and any(
        rule.startswith(("link_", "bookmark_")) or rule == "collapse_all_bookmarks"
        for rule in selected
    ):
        raise ValueError("删除全部链接书签与链接/书签修改规则冲突")
    parameters = payload["rule_parameters"]
    if not isinstance(parameters, dict) or set(parameters) - {"compress_images"}:
        raise ValueError("未知规则参数")
    if "compress_images" in selected:
        params = parameters.get("compress_images")
        if not isinstance(params, dict) or set(params) != {"dpi", "quality"}:
            raise ValueError("图像压缩必须指定 dpi 和 quality")
        for key, low, high in (("dpi", MIN_IMAGE_DPI, MAX_IMAGE_DPI),
                               ("quality", MIN_JPEG_QUALITY, MAX_JPEG_QUALITY)):
            if type(params[key]) is not int or not low <= params[key] <= high:
                raise ValueError("图像压缩参数类型或范围无效: " + key)
    elif parameters:
        raise ValueError("不能为未启用的规则指定参数")
    result = deepcopy(payload)
    result["enabled_rules"] = sorted(selected)
    return result


def _unique_object(pairs):
    result = {}
    for key, value in pairs:
        if key in result:
            raise ValueError("JSON 包含重复字段")
        result[key] = value
    return result


def loads_profile(text):
    try:
        return validate_profile(json.loads(text, object_pairs_hook=_unique_object))
    except (TypeError, json.JSONDecodeError, RecursionError) as exc:
        raise ValueError("无法解析配置 JSON") from exc


def dumps_profile(profile):
    return json.dumps(validate_profile(profile), ensure_ascii=False, sort_keys=True, indent=2) + "\n"


def profile_digest(profile):
    canonical = json.dumps(validate_profile(profile), ensure_ascii=False, sort_keys=True,
                           separators=(",", ":")).encode("utf-8")
    return hashlib.sha256(canonical).hexdigest()


def custom_profile(rules, compression=None, *, profile_id="custom.local", version="1.0.0"):
    selected = sorted(rules)
    return validate_profile({
        "schema_version": SCHEMA_VERSION,
        "profile_id": profile_id,
        "profile_version": version,
        "agency_or_region": "Custom / internal",
        "enabled_rules": selected,
        "rule_parameters": {"compress_images": deepcopy(compression or {
            "dpi": DEFAULT_IMAGE_DPI, "quality": DEFAULT_JPEG_QUALITY,
        })} if "compress_images" in selected else {},
        "validation_level": "structural",
        "source": "User configuration; not official certification",
        "applicability": "User-selected processing rules; validate for the intended workflow",
    })


def builtin_profile(key):
    preset = rules_catalog.PRESETS[key]
    result = custom_profile(preset["options"], profile_id="builtin." + key)
    result.update({
        "agency_or_region": "China" if key == "china" else "United States",
        "source": "RATools existing " + key + " preset; not official certification",
        "applicability": "Existing application convenience preset; independently qualify against current requirements",
    })
    return result


def migrate_favorite(value, compression=None):
    if isinstance(value, str):
        value = [item.strip() for item in value.split(",") if item.strip()]
    if not isinstance(value, list):
        raise ValueError("旧常用配置必须为列表或逗号分隔字符串")
    return custom_profile(value, compression, profile_id="favorite.legacy")


def diff_profiles(before, after):
    before, after = validate_profile(before), validate_profile(after)
    old, new = set(before["enabled_rules"]), set(after["enabled_rules"])
    def flatten(params):
        return {rule + "." + key: value for rule, values in params.items() for key, value in values.items()}
    old_params, new_params = flatten(before["rule_parameters"]), flatten(after["rule_parameters"])
    return {
        "added_rules": sorted(new - old), "removed_rules": sorted(old - new),
        "parameter_changes": {
            key: {"before": old_params.get(key), "after": new_params.get(key)}
            for key in sorted(old_params.keys() | new_params.keys())
            if old_params.get(key) != new_params.get(key)
        },
        "validation_level": {"before": before["validation_level"], "after": after["validation_level"]},
    }
