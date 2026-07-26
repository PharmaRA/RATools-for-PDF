import multiprocessing as mp
import os
import re
import tempfile
from datetime import datetime

from PySide6.QtCore import QThread, Signal

from ratools_pdf.config.features import ENABLE_UPDATE_CHECK
from ratools_pdf.controllers.io_actions import (
    _build_io_paths_for_file,
    _io_action_metadata,
    _normalize_io_action_types,
)
from ratools_pdf.pdf.processor import PDFProcessor

if ENABLE_UPDATE_CHECK:
    from ratools_pdf.services import update_checker
else:
    update_checker = None


def _process_document_task_pipe(file_path, out_path, options, processing_mode, conn):
    try:
        conn.send(PDFProcessor.process_document(file_path, out_path, options, processing_mode=processing_mode))
    except Exception as e:
        conn.send((False, f"处理进程异常: {str(e)}"))
    finally:
        conn.close()


class ProcessWorker(QThread):
    """
    后台处理线程：负责核心的 PDF 批量规则应用，防止 UI 卡死
    """
    progress = Signal(int, str, str)  # row_index, status_text, log_message
    structured_progress = Signal(object)
    finished_all = Signal(str)  # summary
    error = Signal(str)  # error_msg

    def __init__(self, files, options, output_dir, common_base="", overwrite_original=False, max_workers=1, processing_mode="smart"):
        super().__init__()
        self.files = files
        self.options = options
        self.processing_mode = str(processing_mode or "smart").lower()
        self.output_dir = output_dir
        self.common_base = common_base
        self.overwrite_original = overwrite_original
        self.max_workers = max(1, int(max_workers or 1))
        self._stop_requested = False
        self._skip_requested = False
        self._skip_requested_files = set()
        self._can_skip_current = False
        self._current_file_path = ""

    def request_stop(self):
        self._stop_requested = True

    def request_skip_current(self):
        if self._can_skip_current:
            self._skip_requested = True

    def request_skip_file(self, file_path):
        if file_path:
            self._skip_requested_files.add(os.path.normcase(os.path.normpath(file_path)))

    def run(self):
        if self.max_workers <= 1:
            self._run_serial()
        else:
            self._run_parallel()

    def _build_output_path(self, index, file_path, rename_ectd):
        base_name = os.path.basename(file_path)
        if rename_ectd:
            name, ext = os.path.splitext(base_name)
            name = name.lower().replace(" ", "-")
            name = re.sub(r'[^a-z0-9_-]', '', name)
            if not name:
                name = f"doc_{index + 1:03d}"
            base_name = f"{name}{ext.lower()}"

        if self.overwrite_original:
            return base_name, file_path + ".tmp_overwrite.pdf"

        if self.common_base:
            file_dir = os.path.dirname(os.path.abspath(file_path))
            rel_dir = os.path.relpath(file_dir, self.common_base)
            if rel_dir == '.':
                target_dir = self.output_dir
            else:
                target_dir = os.path.join(self.output_dir, rel_dir)
        else:
            target_dir = self.output_dir

        os.makedirs(target_dir, exist_ok=True)
        return base_name, os.path.join(target_dir, base_name)

    def _start_process_task(self, index, file_path, rename_ectd):
        base_name, out_path = self._build_output_path(index, file_path, rename_ectd)
        now = datetime.now().strftime('%H:%M:%S')
        self.structured_progress.emit({
            "type": "start",
            "time": now,
            "file_path": file_path,
            "out_path": out_path,
        })
        self.progress.emit(
            index,
            "正在处理...",
            f"\n[{now}] 开始处理: {file_path}\n    输出文件: {out_path}\n    显示名称: {base_name}",
        )
        parent_conn, child_conn = mp.Pipe(duplex=False)
        proc = mp.Process(target=_process_document_task_pipe, args=(file_path, out_path, self.options, self.processing_mode, child_conn))
        proc.start()
        child_conn.close()
        return {
            "index": index,
            "file_path": file_path,
            "base_name": base_name,
            "out_path": out_path,
            "parent_conn": parent_conn,
            "proc": proc,
        }

    def _terminate_process_task(self, task):
        proc = task["proc"]
        if proc.is_alive():
            try:
                proc.terminate()
            except Exception:
                pass
        proc.join(timeout=2)
        if proc.is_alive():
            try:
                proc.kill()
            except Exception:
                pass
            proc.join(timeout=1)

    def _join_process_task(self, task):
        proc = task["proc"]
        proc.join(timeout=2)
        if proc.is_alive():
            try:
                proc.kill()
            except Exception:
                pass
            proc.join(timeout=1)

    def _remove_partial_output(self, out_path):
        # 正式输出文件，以及子进程被强制终止时可能残留的中间临时文件
        # （pdf_processor 处理流程会先写 out_path + ".tmp.pdf" 再交给 qpdf）。
        for path in (out_path, f"{out_path}.tmp.pdf"):
            if os.path.exists(path):
                try:
                    os.remove(path)
                except Exception:
                    pass

    def _close_task_connection(self, task):
        try:
            task["parent_conn"].close()
        except Exception:
            pass

    def _finish_process_task(self, task, success, msg):
        out_path = task["out_path"]
        if success and self.overwrite_original:
            try:
                os.replace(out_path, task["file_path"])
                out_path = task["file_path"]
            except Exception as e:
                success = False
                msg = f"覆盖原文件失败: {str(e)}"
                self._remove_partial_output(task["out_path"])

        status = "处理完成" if success else "处理失败"
        if not success:
            self._remove_partial_output(task["out_path"])

        now = datetime.now().strftime('%H:%M:%S')
        self.structured_progress.emit({
            "type": "terminal",
            "time": now,
            "file_path": task["file_path"],
            "out_path": out_path,
            "status": status,
            "message": msg,
        })
        self.progress.emit(
            task["index"],
            status,
            f"[{now}] {task['file_path']}\n    状态: {status}\n    输出文件: {out_path}\n    结果: {msg}",
        )
        return success

    def _emit_stopped_task(self, task):
        self._remove_partial_output(task["out_path"])
        now = datetime.now().strftime('%H:%M:%S')
        self.structured_progress.emit({
            "type": "terminal",
            "time": now,
            "file_path": task["file_path"],
            "out_path": task["out_path"],
            "status": "已停止",
            "message": "用户手动停止处理",
        })
        self.progress.emit(
            task["index"],
            "已停止",
            f"[{now}] {task['file_path']}\n    状态: 已停止\n    输出文件: {task['out_path']}\n    原因: 用户手动停止处理",
        )

    def _emit_skipped_task(self, task, reason="已跳过当前文件"):
        self._remove_partial_output(task["out_path"])
        now = datetime.now().strftime('%H:%M:%S')
        self.structured_progress.emit({
            "type": "terminal",
            "time": now,
            "file_path": task["file_path"],
            "out_path": task["out_path"],
            "status": "已跳过",
            "message": reason,
        })
        self.progress.emit(
            task["index"],
            "已跳过",
            f"[{now}] {task['file_path']}\n    状态: 已跳过\n    输出文件: {task['out_path']}\n    原因: {reason}",
        )

    def _run_serial(self):
        try:
            started_at = datetime.now()
            success_count = 0
            rename_ectd = "filename_ectd_format" in self.options
            stopped = False

            for i, file_path in enumerate(self.files):
                if self._stop_requested:
                    stopped = True
                    break

                task = self._start_process_task(i, file_path, rename_ectd)
                base_name = task["base_name"]
                out_path = task["out_path"]
                parent_conn = task["parent_conn"]
                proc = task["proc"]

                success, msg = False, "处理中断"
                skipped_current = False
                self._can_skip_current = True
                self._current_file_path = file_path
                while proc.is_alive():
                    if self._stop_requested:
                        stopped = True
                        self._terminate_process_task(task)
                        break

                    if self._skip_requested:
                        skipped_current = True
                        self._skip_requested = False
                        self._terminate_process_task(task)
                        break

                    if parent_conn.poll(1.0):
                        break

                self._join_process_task(task)
                self._can_skip_current = False
                self._current_file_path = ""

                if stopped:
                    self._emit_stopped_task(task)
                    self._close_task_connection(task)
                    break

                if skipped_current:
                    self._emit_skipped_task(task)
                    self._close_task_connection(task)
                    continue

                if parent_conn.poll():
                    success, msg = parent_conn.recv()
                else:
                    success, msg = False, "处理进程无返回结果"
                parent_conn.close()

                if self._finish_process_task(task, success, msg):
                    success_count += 1

            if stopped:
                summary = f"任务已停止。已成功处理 {success_count} / {len(self.files)} 个文件。"
            else:
                summary = f"处理结束。共成功处理 {success_count} / {len(self.files)} 个文件。"
            elapsed_sec = int((datetime.now() - started_at).total_seconds())
            summary += f" 总耗时 {elapsed_sec}s。"
            self.finished_all.emit(summary)

        except Exception as e:
            self.error.emit(str(e))

    def _run_parallel(self):
        try:
            started_at = datetime.now()
            success_count = 0
            rename_ectd = "filename_ectd_format" in self.options
            stopped = False
            next_index = 0
            running = {}

            while next_index < len(self.files) or running:
                while not self._stop_requested and next_index < len(self.files) and len(running) < self.max_workers:
                    file_path = self.files[next_index]
                    task = self._start_process_task(next_index, file_path, rename_ectd)
                    running[os.path.normcase(os.path.normpath(file_path))] = task
                    next_index += 1

                if self._stop_requested:
                    stopped = True
                    for key, task in list(running.items()):
                        self._terminate_process_task(task)
                        self._emit_stopped_task(task)
                        self._close_task_connection(task)
                        running.pop(key, None)
                    break

                for skip_key in list(self._skip_requested_files):
                    task = running.pop(skip_key, None)
                    self._skip_requested_files.discard(skip_key)
                    if not task:
                        continue
                    self._terminate_process_task(task)
                    self._emit_skipped_task(task, "用户终止选中文件")
                    self._close_task_connection(task)

                finished_keys = []
                for key, task in list(running.items()):
                    parent_conn = task["parent_conn"]
                    proc = task["proc"]
                    if parent_conn.poll():
                        success, msg = parent_conn.recv()
                    elif not proc.is_alive():
                        success, msg = False, "处理进程无返回结果"
                    else:
                        continue

                    proc.join(timeout=2)
                    if proc.is_alive():
                        try:
                            proc.kill()
                        except Exception:
                            pass
                        proc.join(timeout=1)
                    self._close_task_connection(task)
                    if self._finish_process_task(task, success, msg):
                        success_count += 1
                    finished_keys.append(key)

                for key in finished_keys:
                    running.pop(key, None)

                if running and not finished_keys:
                    self.msleep(100)

            if stopped:
                summary = f"任务已停止。已成功处理 {success_count} / {len(self.files)} 个文件。"
            else:
                summary = f"处理结束。共成功处理 {success_count} / {len(self.files)} 个文件。"
            elapsed_sec = int((datetime.now() - started_at).total_seconds())
            summary += f" 总耗时 {elapsed_sec}s。"
            self.finished_all.emit(summary)

        except Exception as e:
            self.error.emit(str(e))


class PreCheckWorker(QThread):
    """
    后台预检线程：只读取 PDF 状态，生成建议勾选的处理项，不修改文件。
    """
    progress = Signal(int, str, str)
    result_ready = Signal(dict)
    finished_precheck = Signal(str)
    error_precheck = Signal(str)

    def __init__(self, files):
        super().__init__()
        self.files = list(files)

    def run(self):
        try:
            started_at = datetime.now()
            suggested_files = 0
            review_files = 0
            failed_files = 0

            for i, file_path in enumerate(self.files):
                base_name = os.path.basename(file_path)
                self.progress.emit(
                    i,
                    "正在预检...",
                    f"\n[{datetime.now().strftime('%H:%M:%S')}] 开始预检: {base_name}",
                )

                report = PDFProcessor.build_precheck_report(file_path)
                if not report.get("available"):
                    failed_files += 1
                    reason = report.get("error") or "无法读取PDF结构"
                    self.result_ready.emit({
                        "file_name": base_name,
                        "file_path": file_path,
                        "status": "预检失败",
                        "suggestions": "",
                        "suggestion_ids": "",
                        "error": reason,
                        "font_summary": "",
                        "font_details": "",
                        "annotation_summary": "",
                        "annotation_details": "",
                        "broken_reference_summary": "",
                        "broken_reference_details": "",
                    })
                    self.progress.emit(
                        i,
                        "预检失败",
                        f"[{datetime.now().strftime('%H:%M:%S')}] {base_name}\n    状态: 预检失败\n    原因: {reason}",
                    )
                    continue

                report_suggestions = report.get("suggestions", {})
                suggestions = list(report_suggestions.values())
                if suggestions:
                    actionable_ids = [
                        option_id
                        for option_id, item in report_suggestions.items()
                        if not item.get("report_only")
                    ]
                    has_report_only = any(item.get("report_only") for item in suggestions)
                    if actionable_ids:
                        suggested_files += 1
                    if has_report_only:
                        review_files += 1
                    suggestion_ids = ",".join(actionable_ids)
                    status = "建议处理" if actionable_ids else "需要复核"
                    advice_parts = []
                    for item in suggestions:
                        title = item.get("title", "")
                        if not title:
                            continue
                        if item.get("report_only") and item.get("reason"):
                            advice_parts.append(f"{title}：{item.get('reason')}")
                        else:
                            advice_parts.append(title)
                    advice = "、".join(advice_parts)
                    self.result_ready.emit({
                        "file_name": base_name,
                        "file_path": file_path,
                        "status": status,
                        "suggestions": advice,
                        "suggestion_ids": suggestion_ids,
                        "error": "",
                        "font_summary": report.get("font_summary", ""),
                        "font_details": report.get("font_details", ""),
                        "annotation_summary": report.get("annotation_summary", ""),
                        "annotation_details": report.get("annotation_details", ""),
                        "broken_reference_summary": report.get("broken_reference_summary", ""),
                        "broken_reference_details": report.get("broken_reference_details", ""),
                    })
                    self.progress.emit(
                        i,
                        status,
                        f"[{datetime.now().strftime('%H:%M:%S')}] {base_name}\n    状态: {status}\n    建议: {advice}",
                    )
                else:
                    self.result_ready.emit({
                        "file_name": base_name,
                        "file_path": file_path,
                        "status": "无需处理",
                        "suggestions": "",
                        "suggestion_ids": "",
                        "error": "",
                        "font_summary": report.get("font_summary", ""),
                        "font_details": report.get("font_details", ""),
                        "annotation_summary": report.get("annotation_summary", ""),
                        "annotation_details": report.get("annotation_details", ""),
                        "broken_reference_summary": report.get("broken_reference_summary", ""),
                        "broken_reference_details": report.get("broken_reference_details", ""),
                    })
                    self.progress.emit(
                        i,
                        "无需处理",
                        f"[{datetime.now().strftime('%H:%M:%S')}] {base_name}\n    状态: 无需处理\n    建议: 未发现当前可自动处理的明显问题",
                    )

            elapsed_sec = int((datetime.now() - started_at).total_seconds())
            summary = (
                f"预检结束。共检查 {len(self.files)} 个文件，"
                f"发现 {suggested_files} 个文件存在建议处理项，"
                f"{review_files} 个文件需要人工复核，"
                f"{failed_files} 个文件预检失败。总耗时 {elapsed_sec}s。"
            )
            self.finished_precheck.emit(summary)
        except Exception as e:
            self.error_precheck.emit(str(e))


class DetectionWorker(QThread):
    """后台检测线程：对整个队列执行单项只读检测（批注 / 失效引用文本），不修改文件。

    detection_kind:
        - "annotation"：检测便签、高亮等批注
        - "broken_reference"：检测 Word 转 PDF 残留的失效引用/链接占位文本
    """
    progress = Signal(int, str, str)
    result_ready = Signal(dict)
    finished_detection = Signal(str, list)
    error_detection = Signal(str)

    KIND_LABELS = {
        "annotation": "批注检测",
        "broken_reference": "失效引用检测",
    }

    def __init__(self, files, detection_kind):
        super().__init__()
        self.files = list(files)
        self.detection_kind = detection_kind if detection_kind in self.KIND_LABELS else "annotation"

    def _collect(self, file_path):
        if self.detection_kind == "annotation":
            findings = PDFProcessor._collect_annotation_findings_for_path(file_path)
            hit = bool(findings.get("has_annotations"))
        else:
            findings = PDFProcessor._collect_broken_reference_findings_for_path(file_path)
            hit = bool(findings.get("has_broken_reference"))
        return findings, hit

    def run(self):
        try:
            started_at = datetime.now()
            kind_label = self.KIND_LABELS[self.detection_kind]
            hit_files = 0
            failed_files = 0
            results = []

            for i, file_path in enumerate(self.files):
                base_name = os.path.basename(file_path)
                self.progress.emit(
                    i,
                    "正在检测...",
                    f"\n[{datetime.now().strftime('%H:%M:%S')}] 开始{kind_label}: {base_name}",
                )

                findings, hit = self._collect(file_path)
                summary = findings.get("summary", "")
                details = findings.get("details", "")

                if not findings.get("available"):
                    failed_files += 1
                    reason = findings.get("error") or "无法读取PDF结构"
                    status = "检测失败"
                    row = {
                        "file_name": base_name,
                        "file_path": file_path,
                        "detection_kind": self.detection_kind,
                        "status": status,
                        "summary": "",
                        "details": "",
                        "error": reason,
                    }
                    results.append(row)
                    self.result_ready.emit(row)
                    self.progress.emit(
                        i,
                        status,
                        f"[{datetime.now().strftime('%H:%M:%S')}] {base_name}\n    状态: {status}\n    原因: {reason}",
                    )
                    continue

                status = "发现问题" if hit else "未发现"
                if hit:
                    hit_files += 1
                row = {
                    "file_name": base_name,
                    "file_path": file_path,
                    "detection_kind": self.detection_kind,
                    "status": status,
                    "summary": summary,
                    "details": details,
                    "error": "",
                }
                results.append(row)
                self.result_ready.emit(row)
                log_detail = summary
                if details:
                    log_detail = f"{summary}；明细：{details}" if summary else details
                self.progress.emit(
                    i,
                    status,
                    f"[{datetime.now().strftime('%H:%M:%S')}] {base_name}\n    状态: {status}\n    结果: {log_detail or '未发现相关内容'}",
                )

            elapsed_sec = int((datetime.now() - started_at).total_seconds())
            summary = (
                f"{kind_label}结束。共检查 {len(self.files)} 个文件，"
                f"发现 {hit_files} 个文件存在相关内容，"
                f"{failed_files} 个文件检测失败。总耗时 {elapsed_sec}s。"
            )
            self.finished_detection.emit(summary, results)
        except Exception as e:
            self.error_detection.emit(str(e))


class IOActionWorker(QThread):
    """
    高级 IO 操作后台线程：处理书签、链接等需要长时间读写的批量导入/导出操作
    """
    progress = Signal(int, str, str)
    finished_action = Signal(str)
    error_action = Signal(str)

    def __init__(self, action_type, files, target_dir, output_dir=None, common_base="",
                 link_scope="all", link_mode="overwrite"):
        super().__init__()
        self.action_types = _normalize_io_action_types(action_type)
        self.action_type = self.action_types[0] if self.action_types else ""
        self.files = list(files)
        self.target_dir = target_dir
        self.output_dir = output_dir
        self.common_base = common_base
        self.link_scope = "external" if link_scope == "external" else "all"
        self.link_mode = "incremental" if link_mode == "incremental" else "overwrite"

    def run(self):
        try:
            for row_idx, file_path in enumerate(self.files):
                base_name = os.path.basename(file_path)

                self.progress.emit(row_idx, "正在执行...",
                                   f"[{datetime.now().strftime('%H:%M:%S')}] 正在处理: {base_name}")
                success, messages = False, []

                if self.action_type.startswith("export"):
                    for action_type in self.action_types:
                        meta = _io_action_metadata(action_type)
                        data_path, _ = _build_io_paths_for_file(
                            file_path, meta["data_kind"], self.target_dir, common_base=self.common_base
                        )
                        os.makedirs(os.path.dirname(data_path), exist_ok=True)
                        if action_type == 'export_bookmarks':
                            PDFProcessor.export_bookmarks(file_path, data_path)
                            messages.append("✅ 导出书签成功")
                        elif action_type == 'export_links':
                            PDFProcessor.export_links(file_path, data_path, scope=self.link_scope)
                            messages.append("✅ 导出%s成功" % ("外部链接" if self.link_scope == "external" else "链接"))
                    success = bool(messages)

                elif self.action_type.startswith("import"):
                    matched_actions = []
                    final_out_pdf = None
                    for action_type in self.action_types:
                        meta = _io_action_metadata(action_type)
                        data_path, out_pdf = _build_io_paths_for_file(
                            file_path, meta["data_kind"], self.target_dir, self.output_dir, self.common_base
                        )
                        final_out_pdf = final_out_pdf or out_pdf
                        if os.path.exists(data_path):
                            matched_actions.append((action_type, data_path))
                        else:
                            messages.append(f"⚠️ 未找到匹配的{meta['data_type']}文件")

                    if matched_actions:
                        if not final_out_pdf:
                            raise ValueError("导入书签或链接时缺少输出目录")
                        os.makedirs(os.path.dirname(final_out_pdf), exist_ok=True)
                        current_source = file_path
                        temp_paths = []
                        try:
                            for action_index, (action_type, data_path) in enumerate(matched_actions):
                                is_last_action = action_index == len(matched_actions) - 1
                                if is_last_action:
                                    current_output = final_out_pdf
                                else:
                                    fd, current_output = tempfile.mkstemp(suffix=".pdf", prefix="ratools_io_")
                                    os.close(fd)
                                    temp_paths.append(current_output)

                                if action_type == 'import_bookmarks':
                                    PDFProcessor.import_bookmarks(current_source, data_path, current_output)
                                    messages.append("✅ 导入书签成功")
                                elif action_type == 'import_links':
                                    PDFProcessor.import_links(
                                        current_source, data_path, current_output,
                                        scope=self.link_scope, mode=self.link_mode,
                                    )
                                    scope_label = "外部链接" if self.link_scope == "external" else "链接"
                                    mode_label = "增量" if self.link_mode == "incremental" else "覆盖"
                                    messages.append(f"✅ 导入{scope_label}成功（{mode_label}）")
                                current_source = current_output
                            success = True
                        finally:
                            for temp_path in temp_paths:
                                try:
                                    if os.path.exists(temp_path):
                                        os.remove(temp_path)
                                except OSError:
                                    pass

                msg = "；".join(messages)

                status = "操作成功" if success else "操作失败"
                if "未找到匹配" in msg:
                    status = "未匹配跳过"
                    if success:
                        status = "操作成功"
                self.progress.emit(row_idx, status, f"   ↳ 结果: {msg}")

            action_name = "导出" if self.action_type.startswith("export") else "导入"
            self.finished_action.emit(f"批量高级 '{action_name}' 任务执行完成。")

        except Exception as e:
            self.error_action.emit(str(e))


class UpdateCheckWorker(QThread):
    finished_check = Signal(object, bool)

    def __init__(self, silent=False, parent=None):
        super().__init__(parent)
        self.silent = silent

    def run(self):
        if not ENABLE_UPDATE_CHECK or update_checker is None:
            return
        result = update_checker.check_for_updates()
        self.finished_check.emit(result, self.silent)
