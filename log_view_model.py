import os
import re


SUCCESS_STATUSES = {"处理完成", "操作成功", "无需处理"}
FAILURE_STATUSES = {"处理失败", "操作失败", "预检失败"}
SKIP_STATUSES = {"已跳过", "已停止", "未匹配跳过"}
PRECHECK_STATUSES = {"建议处理", "需要复核", "无需处理", "预检失败"}


def split_log_blocks(raw_text):
    blocks = []
    current = []
    for raw_line in (raw_text or "").splitlines():
        line = raw_line.rstrip()
        if not line:
            if current:
                blocks.append("\n".join(current))
                current = []
            continue
        if re.match(r"^\[\d{2}:\d{2}:\d{2}\]", line) and current:
            blocks.append("\n".join(current))
            current = [line]
            continue
        current.append(line)
    if current:
        blocks.append("\n".join(current))
    return blocks


def log_status_tags(status, text=""):
    tags = set()
    if status in SUCCESS_STATUSES:
        tags.add("success")
    if status in FAILURE_STATUSES or "[致命错误]" in text or "[预检错误]" in text:
        tags.add("failure")
    if status in SKIP_STATUSES:
        tags.add("skip")
    if status in PRECHECK_STATUSES or "预检" in text:
        tags.add("precheck")
    return tags


def format_duration(seconds):
    if seconds in ("", None):
        return ""
    try:
        value = int(seconds)
    except (TypeError, ValueError):
        return ""
    if value < 60:
        return f"{value}s"
    minutes, rest = divmod(value, 60)
    if minutes < 60:
        return f"{minutes}m {rest}s"
    hours, minutes = divmod(minutes, 60)
    return f"{hours}h {minutes}m {rest}s"


def _basename(path):
    normalized = str(path or "").replace("\\", "/")
    return os.path.basename(normalized) or normalized


def _find_detail_for_row(blocks, row):
    candidates = [row.get("file_original", ""), row.get("file_output", "")]
    matches = []
    for block in blocks:
        if any(candidate and candidate in block for candidate in candidates):
            matches.append(block)
    if matches:
        return "\n\n".join(matches)
    lines = [
        f"[{row.get('time', '')}] {row.get('file_original', '')}".strip(),
        f"    状态: {row.get('status', '')}",
    ]
    if row.get("file_output"):
        lines.append(f"    输出文件: {row.get('file_output')}")
    if row.get("changes"):
        lines.append(f"    结果: {row.get('changes')}")
    return "\n".join(line for line in lines if line.strip())


def _item_from_row(row, blocks):
    detail = _find_detail_for_row(blocks, row)
    file_original = row.get("file_original", "")
    file_output = row.get("file_output", "")
    status = row.get("status", "")
    result = row.get("changes", "")
    item = {
        "time": row.get("time", ""),
        "file_name": _basename(file_original),
        "file_original": file_original,
        "file_output": file_output,
        "output_name": _basename(file_output),
        "status": status,
        "duration_text": format_duration(row.get("duration_sec", "")),
        "result": result,
        "detail": detail,
        "tags": log_status_tags(status, detail),
    }
    item["search_text"] = " ".join(str(value) for key, value in item.items() if key != "tags")
    return item


def _fallback_rows_from_blocks(blocks):
    rows = []
    for block in blocks:
        status_match = re.search(r"^\s*状态:\s*(.+)$", block, re.MULTILINE)
        if not status_match:
            continue
        first_line = block.splitlines()[0] if block.splitlines() else ""
        time_match = re.match(r"^\[(\d{2}:\d{2}:\d{2})\]\s+(.+)$", first_line)
        output_match = re.search(r"^\s*输出文件:\s*(.+)$", block, re.MULTILINE)
        result_match = re.search(r"^\s*(?:结果|原因):\s*(.+)$", block, re.MULTILINE)
        file_original = time_match.group(2).strip() if time_match else ""
        rows.append({
            "time": time_match.group(1) if time_match else "",
            "file_original": file_original,
            "file_output": output_match.group(1).strip() if output_match else "",
            "status": status_match.group(1).strip(),
            "duration_sec": "",
            "changes": result_match.group(1).strip() if result_match else "",
        })
    return rows


def build_log_summary_items(raw_text, structured_rows=None):
    blocks = split_log_blocks(raw_text)
    rows = list(structured_rows or []) or _fallback_rows_from_blocks(blocks)
    return [_item_from_row(row, blocks) for row in rows]


def filter_log_summary_items(items, mode="all", query=""):
    normalized_query = (query or "").strip().lower()
    filtered = []
    for item in items:
        if mode != "all" and mode not in item.get("tags", set()):
            continue
        if normalized_query and normalized_query not in item.get("search_text", "").lower():
            continue
        filtered.append(item)
    return filtered
