import re


def _render_logs_as_csv_rows(log_text):
    rows = []
    starts_by_file = {}
    pending_result = None
    last_start_file = None
    _time_to_seconds = _log_time_to_seconds

    for raw_line in log_text.splitlines():
        line = raw_line.rstrip()
        if not line:
            continue

        start_match = re.match(r"^\[(\d{2}:\d{2}:\d{2})\]\s+开始处理:\s+(.+)$", line)
        if start_match:
            start_time = start_match.group(1)
            original_file = start_match.group(2)
            starts_by_file[original_file] = {
                "seconds": _time_to_seconds(start_time),
                "output": original_file,
            }
            last_start_file = original_file
            continue

        output_match = re.match(r"^\s+输出文件:\s+(.+)$", line)
        if output_match:
            output_file = output_match.group(1)
            if pending_result:
                pending_result["file_output"] = output_file
            elif last_start_file in starts_by_file:
                starts_by_file[last_start_file]["output"] = output_file
            elif rows:
                rows[-1]["file_output"] = output_file
            continue

        result_match = re.match(r"^\[(\d{2}:\d{2}:\d{2})\]\s+(.+)$", line)
        if result_match and "开始处理:" not in line:
            result_time = result_match.group(1)
            output_file = result_match.group(2)
            if output_file not in starts_by_file:
                pending_result = None
                last_start_file = None
                continue
            start_info = starts_by_file[output_file]
            pending_result = {
                "time": result_time,
                "file_original": output_file,
                "file_output": start_info.get("output") or output_file,
                "start_seconds": start_info.get("seconds"),
            }
            last_start_file = None
            continue

        status_match = re.match(r"^\s+状态:\s+(.+)$", line)
        if status_match and pending_result:
            status_value = status_match.group(1)
            end_seconds = _time_to_seconds(pending_result["time"])
            duration_sec = ""
            if pending_result["start_seconds"] is not None and end_seconds is not None:
                delta = end_seconds - pending_result["start_seconds"]
                if delta < 0:
                    delta += 24 * 3600
                duration_sec = delta

            rows.append({
                "time": pending_result["time"],
                "file_original": pending_result["file_original"],
                "file_output": pending_result["file_output"],
                "status": status_value,
                "success": "true" if status_value == "处理完成" else "false",
                "duration_sec": duration_sec,
                "changes": "",
            })
            starts_by_file.pop(pending_result["file_original"], None)
            pending_result = None
            continue

        result_line_match = re.match(r"^\s+结果:\s+(.+)$", line)
        if result_line_match and rows:
            result_text = result_line_match.group(1)
            if "修改项：" in result_text:
                rows[-1]["changes"] = result_text.split("修改项：", 1)[1].strip()

    return rows


def _log_time_to_seconds(value):
    try:
        hh, mm, ss = value.split(":")
        return int(hh) * 3600 + int(mm) * 60 + int(ss)
    except Exception:
        return None


def _structured_log_row_from_event(event, start_events=None):
    status_text = event.get("status", "")
    if status_text not in ("处理完成", "处理失败", "已跳过", "已停止"):
        return None

    file_path = event.get("file_path", "")
    output_path = event.get("out_path", "") or file_path
    event_time = event.get("time", "")
    start_info = (start_events or {}).get(file_path, {})
    start_seconds = _log_time_to_seconds(start_info.get("time", ""))
    end_seconds = _log_time_to_seconds(event_time)
    duration_sec = ""
    if start_seconds is not None and end_seconds is not None:
        duration_sec = end_seconds - start_seconds
        if duration_sec < 0:
            duration_sec += 24 * 3600

    result_text = event.get("message", "")
    changes = ""
    if "修改项：" in result_text:
        changes = result_text.split("修改项：", 1)[1].strip()

    row = {
        "time": event_time,
        "file_original": file_path,
        "file_output": output_path,
        "status": status_text,
        "success": "true" if status_text == "处理完成" else "false",
        "duration_sec": duration_sec,
        "changes": changes,
    }

    return row


def _select_log_rows_for_export(structured_rows, log_text):
    if structured_rows:
        return list(structured_rows)
    return _render_logs_as_csv_rows(log_text)
