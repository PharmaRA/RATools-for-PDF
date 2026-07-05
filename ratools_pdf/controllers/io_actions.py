import os
import re


def _safe_relative_subdir(file_path, common_base):
    if not common_base:
        return ''
    try:
        rel_dir = os.path.relpath(os.path.dirname(os.path.abspath(file_path)), common_base)
    except ValueError:
        return ''
    rel_norm = os.path.normpath(rel_dir)
    if rel_norm in ('.', ''):
        return ''
    if rel_norm.startswith('..') or os.path.isabs(rel_norm):
        safe_leaf = os.path.basename(os.path.dirname(os.path.abspath(file_path))) or 'external'
        return os.path.join('_external', safe_leaf)
    return rel_norm


def _normalized_ectd_name(base_name, fallback_index):
    name, ext = os.path.splitext(base_name)
    normalized = name.lower().replace(' ', '-')
    normalized = re.sub(r'[^a-z0-9_-]', '', normalized)
    if not normalized:
        normalized = f'doc_{fallback_index:03d}'
    return f'{normalized}{ext.lower()}'


def _collect_ectd_rename_plan(file_paths):
    rename_pairs = []
    target_names = {}
    for i, file_path in enumerate(file_paths, start=1):
        base_name = os.path.basename(file_path)
        new_name = _normalized_ectd_name(base_name, i)
        if new_name != base_name:
            rename_pairs.append((base_name, new_name))
        target_names.setdefault(new_name, []).append(file_path)
    collisions = {name: paths for name, paths in target_names.items() if len(paths) > 1}
    return rename_pairs, collisions


def _build_io_paths_for_file(file_path, data_kind, target_dir, output_dir=None, common_base=''):
    base_name = os.path.basename(file_path)
    name_no_ext, _ = os.path.splitext(base_name)
    suffix = 'bookmarks.csv' if data_kind == 'bookmarks' else 'links.json'

    rel_dir = _safe_relative_subdir(file_path, common_base)
    data_parent = os.path.join(target_dir, rel_dir) if rel_dir else target_dir
    data_path = os.path.join(data_parent, f'{name_no_ext}_{suffix}')

    output_path = None
    if output_dir:
        output_parent = os.path.join(output_dir, rel_dir) if rel_dir else output_dir
        output_path = os.path.join(output_parent, base_name)

    return data_path, output_path


def _io_action_metadata(action_type):
    is_bookmarks = "bookmarks" in action_type
    is_export = "export" in action_type
    return {
        "data_kind": "bookmarks" if is_bookmarks else "links",
        "data_label": "书签" if is_bookmarks else "链接",
        "data_type": "CSV" if is_bookmarks else "JSON",
        "is_export": is_export,
        "action_name": "导出" if is_export else "导入",
    }


def _normalize_io_action_types(action_type):
    if isinstance(action_type, (list, tuple)):
        return list(action_type)
    return [action_type]


def _build_io_preview_rows(files, action_type, target_dir, common_base=""):
    rows = []
    for current_action_type in _normalize_io_action_types(action_type):
        meta = _io_action_metadata(current_action_type)
        for file_path in files:
            data_path, _ = _build_io_paths_for_file(
                file_path,
                meta["data_kind"],
                target_dir,
                common_base=common_base,
            )
            if meta["is_export"]:
                status = "将生成"
            else:
                status = "已匹配" if os.path.exists(data_path) else "未找到"
            rows.append({
                "action_type": current_action_type,
                "data_kind": meta["data_kind"],
                "data_label": meta["data_label"],
                "file_path": file_path,
                "file_name": os.path.basename(file_path),
                "data_path": data_path,
                "status": status,
            })
    return rows
