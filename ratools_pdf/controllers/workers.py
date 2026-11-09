import multiprocessing as mp
import os
import tempfile
import threading
import time
import uuid
from datetime import datetime

from PySide6.QtCore import QThread, Signal

from ratools_pdf.config.features import ENABLE_UPDATE_CHECK
from ratools_pdf.controllers.io_actions import (
    _build_io_paths_for_file,
    _io_action_metadata,
    _normalize_io_action_types,
    _normalized_ectd_name,
)
from ratools_pdf.pdf import inspect as pdf_inspect
from ratools_pdf.pdf.processor import PDFProcessor
from ratools_pdf.services.file_safety import check_unlinked_path, recycle_owned_path
from ratools_pdf.services.output_paths import common_base_dir, path_key, relative_source_directory, validate_output_plan
from ratools_pdf.services.pdf_facts import InspectionCancelled, capture_pdf_snapshot, verify_output
from ratools_pdf.services.processing_plan import configuration_matches_plan, input_matches_plan, source_file_id, timestamp
from ratools_pdf.services import process_control
from ratools_pdf.services.rule_results import actual_changed, merge_execution

if ENABLE_UPDATE_CHECK:
    from ratools_pdf.services import update_checker
else:
    update_checker = None


def _process_document_task_pipe(file_path, out_path, options, processing_mode, compression_settings, conn):
    try:
        conn.send(PDFProcessor.process_document_detailed(
            file_path, out_path, options,
            processing_mode=processing_mode,
            compression_settings=compression_settings,
        ))
    except Exception as e:
        conn.send((False, f"处理进程异常: {str(e)}"))
    finally:
        conn.close()


class ProcessWorker(QThread):
    """
    后台处理线程：负责核心的 PDF 批量规则应用，防止 UI 卡死
    """
    progress = Signal(str, str, str)  # file_path, status_text, log_message
    structured_progress = Signal(object)
    finished_all = Signal(str)  # summary
    error = Signal(str)  # error_msg

    def __init__(self, files, options, output_dir, common_base="", overwrite_original=False, max_workers=1, processing_mode="smart", compression_settings=None, processing_plan=None, file_timeout_seconds=0):
        super().__init__()
        self.files = list(files)
        self.options = set(options)
        self.processing_mode = str(processing_mode or "smart").lower()
        self.compression_settings = dict(compression_settings) if compression_settings is not None else None
        self.processing_plan = processing_plan
        self.batch_id = processing_plan['batch_id'] if processing_plan else uuid.uuid4().hex
        self.file_plans = {item['file_id']: item for item in (processing_plan or {}).get('files', [])}
        self.output_dir = output_dir
        self.common_base = common_base or common_base_dir(self.files)
        self.overwrite_original = overwrite_original
        self.max_workers = max(1, int(max_workers or 1))
        self.effective_max_workers = self.max_workers
        self.file_timeout_seconds = process_control.normalize_file_timeout(file_timeout_seconds)
        self._control_lock = threading.RLock()
        self._stop_requested = False
        self._skip_requested = False
        self._skip_requested_files = set()
        self._can_skip_current = False
        self._current_file_path = ""
        self._verification_file_path = ""

    def request_stop(self):
        with self._control_lock:
            self._stop_requested = True

    def request_skip_current(self):
        with self._control_lock:
            if self._can_skip_current or self._verification_file_path:
                self._skip_requested = True

    def request_skip_file(self, file_path):
        with self._control_lock:
            if file_path:
                self._skip_requested_files.add(path_key(file_path))

    def _clear_current_skip_state(self):
        with self._control_lock:
            self._can_skip_current = False
            self._current_file_path = ""
            self._verification_file_path = ""
            self._skip_requested = False

    def run(self):
        try:
            rename_ectd = "filename_ectd_format" in self.options
            pairs = [(path, self._build_output_path(i, path, rename_ectd)[1])
                     for i, path in enumerate(self.files)]
            validate_output_plan(pairs, self.files, allow_in_place=self.overwrite_original)
        except Exception as exc:
            self.error.emit(str(exc))
            return
        input_sizes = []
        for path in self.files:
            try:
                input_sizes.append(os.path.getsize(path))
            except OSError:
                input_sizes.append(None)
        self.effective_max_workers = process_control.concurrency_limit(
            self.max_workers, input_sizes, process_control.available_memory_bytes())
        if self.effective_max_workers <= 1:
            self._run_serial()
        else:
            self._run_parallel()

    def _build_output_path(self, index, file_path, rename_ectd):
        base_name = os.path.basename(file_path)
        if rename_ectd:
            # 与 IO 向导预览共用同一实现，保证"预览名 == 实际输出名"
            base_name = _normalized_ectd_name(base_name, index + 1)

        if self.overwrite_original:
            return base_name, file_path

        rel_dir = relative_source_directory(file_path, self.common_base)
        target_dir = os.path.join(self.output_dir, rel_dir) if rel_dir else self.output_dir
        return base_name, os.path.join(target_dir, base_name)

    def _start_process_task(self, index, file_path, rename_ectd):
        base_name, out_path = self._build_output_path(index, file_path, rename_ectd)
        file_plan = self.file_plans.get(source_file_id(file_path))
        task = {'index': index, 'file_path': file_path, 'base_name': base_name, 'out_path': out_path,
                'file_plan': file_plan, 'started_at': timestamp()}
        seconds = process_control.normalize_file_timeout(self.file_timeout_seconds)
        task['deadline'] = time.monotonic() + seconds if seconds else None
        cancelled = lambda: self._task_cancelled(task)
        try:
            if self.processing_plan:
                if not file_plan or file_plan['status'] != 'ready':
                    task.update(rejection='文件计划检查失败，未执行处理', outcome='planning_failed')
                    return task
                task['before'] = file_plan['before']
                if not configuration_matches_plan(self.processing_plan, self.options, self.processing_mode, self.compression_settings):
                    task.update(rejection='处理配置已变化，请重新生成计划', outcome='stale_plan')
                    return task
                if not input_matches_plan(file_path, file_plan, cancelled):
                    task.update(rejection='输入文件已变化，请重新预检并生成计划', outcome='stale_plan')
                    return task
            else:
                task['before'] = capture_pdf_snapshot(file_path, cancelled)
        except InspectionCancelled:
            task['rejection'] = '输入检查已中断，未执行处理'
            return task
        if cancelled():
            task['rejection'] = '输入检查已中断，未执行处理'
            return task
        check_unlinked_path(file_path)
        check_unlinked_path(out_path)
        same_path = os.path.normcase(os.path.abspath(file_path)) == os.path.normcase(os.path.abspath(out_path))
        if not self.overwrite_original and same_path:
            raise ValueError("输出路径与输入文件相同，请选择其他输出目录或明确启用覆盖原文件")
        # Same filesystem as the final destination, so a successful publish is atomic.
        try:
            os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
            process_control.ensure_output_capacity(os.path.dirname(os.path.abspath(out_path)), os.path.getsize(file_path))
        except OSError as exc:
            task.update(rejection=f'输出预检失败: {exc}', outcome='resource_preflight_failed')
            return task
        staging_dir = tempfile.mkdtemp(prefix=".ratools-task-", dir=os.path.dirname(os.path.abspath(out_path)))
        work_path = os.path.join(staging_dir, "result.pdf")
        now = datetime.now().strftime('%H:%M:%S')
        self.structured_progress.emit({
            "type": "start",
            "time": now,
            "file_path": file_path,
            "out_path": out_path,
            **self._task_facts(task),
        })
        self.progress.emit(
            file_path,
            "正在处理...",
            f"\n[{now}] 开始处理: {file_path}\n    输出文件: {out_path}\n    显示名称: {base_name}",
        )
        task.update(work_path=work_path, staging_dir=staging_dir)
        parent_conn = child_conn = None
        proc = None
        try:
            parent_conn, child_conn = mp.Pipe(duplex=False)
            proc = mp.Process(
                target=_process_document_task_pipe,
                args=(file_path, work_path, set(file_plan['effective_options']) if file_plan else self.options,
                      'force' if file_plan else self.processing_mode, self.compression_settings, child_conn),
            )
            proc.start()
            task.update(parent_conn=parent_conn, proc=proc)
            return task
        except Exception as exc:
            if proc is not None and proc.pid is not None:
                self._terminate_process_task({'proc': proc})
            if parent_conn is not None:
                parent_conn.close()
            warning = self._recycle_task_output(task)
            if warning:
                raise RuntimeError(f"{exc}；{warning}") from exc
            raise
        finally:
            if child_conn is not None:
                child_conn.close()

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

    def _recycle_task_output(self, task):
        staging_dir = task.get("staging_dir")
        if not staging_dir:
            return ""
        try:
            recycle_owned_path(staging_dir, staging_dir)
            return ""
        except Exception:
            return f"临时文件未能移入回收站，已保留于: {staging_dir}"

    def _close_task_connection(self, task):
        try:
            task["parent_conn"].close()
        except Exception:
            pass

    def _fail_process_task(self, task, message):
        try:
            if "proc" in task:
                self._terminate_process_task(task)
        except Exception as exc:
            message += f"；进程收尾异常: {exc}"
        finally:
            if "parent_conn" in task:
                self._close_task_connection(task)
        self._finish_process_task(task, False, message)

    def _task_facts(self, task):
        output_relative_path = None
        if task.get('out_path'):
            base = self.common_base if self.overwrite_original else self.output_dir
            try:
                relative = os.path.relpath(task['out_path'], base)
                if relative != '..' and not relative.startswith('..' + os.sep) and not os.path.isabs(relative):
                    output_relative_path = relative.replace(os.sep, '/')
            except ValueError:
                pass
        return {'batch_id': self.batch_id, 'file_id': source_file_id(task['file_path']),
                'output_relative_path': output_relative_path,
                'started_at': task.get('started_at'), 'before': task.get('before'),
                'after': task.get('after'), 'verification': task.get('verification'),
                'engine_success': task.get('engine_success'), 'published': task.get('published', False),
                'outcome': task.get('outcome'),
                'rule_actual_changed': actual_changed(task.get('rule_results', [])),
                'rule_results': task.get('rule_results', []),
                'rule_events': task.get('rule_events', [])}

    def _consume_engine_result(self, task, payload):
        if not isinstance(payload, dict):
            return payload
        results, events = merge_execution(
            (task.get('file_plan') or {}).get('rule_results', []),
            payload.get('rule_results', []), payload.get('rule_events', []))
        task.update(rule_results=results, rule_events=events)
        return bool(payload.get('success')), payload.get('message', '')

    def _task_expired(self, task):
        return task.get('deadline') is not None and time.monotonic() >= task['deadline']

    def _task_cancelled(self, task):
        return (self._stop_requested or self._skip_requested or
                path_key(task['file_path']) in self._skip_requested_files or self._task_expired(task))

    def _emit_timed_out_task(self, task):
        self._clear_current_skip_state()
        task['outcome'] = 'timed_out'
        task['published'] = False
        warning = self._recycle_task_output(task)
        message = '每文件处理超时，已停止处理并保留既有输出；可调整超时后重试'
        if warning:
            message += f'；{warning}'
        now = datetime.now().strftime('%H:%M:%S')
        self.structured_progress.emit({'type': 'terminal', 'time': now,
                                      'file_path': task['file_path'], 'out_path': task['out_path'],
                                      'status': '处理失败', 'message': message,
                                      'ended_at': timestamp(), **self._task_facts(task)})
        self.progress.emit(task['file_path'], '处理失败', f'[{now}] {message}')

    def _finish_interrupted_task(self, task):
        """Caller holds _control_lock; a terminal decision precedes publication."""
        current_key = path_key(task['file_path'])
        if self._stop_requested:
            task['outcome'] = 'stopped'
            self._emit_stopped_task(task)
            return True
        if self._skip_requested or current_key in self._skip_requested_files:
            task['outcome'] = 'skipped'
            self._skip_requested = False
            self._skip_requested_files.discard(current_key)
            self._emit_skipped_task(task)
            return True
        if self._task_expired(task):
            self._emit_timed_out_task(task)
            return True
        return False

    def _finish_process_task(self, task, success, msg):
        out_path = task["out_path"]
        task['engine_success'] = bool(success)
        with self._control_lock:
            if self._finish_interrupted_task(task):
                return False
            self._verification_file_path = task['file_path']
        cancelled = lambda: self._task_cancelled(task)
        try:
            if 'before' not in task:
                task['before'] = capture_pdf_snapshot(task['file_path'], cancelled)
            if task.get('work_path') and os.path.isfile(task['work_path']):
                task['after'] = capture_pdf_snapshot(task['work_path'], cancelled)
            task['verification'] = verify_output(task.get('before') or {}, task.get('after') or {},
                                                 (task.get('file_plan') or {}).get('effective_options', self.options))
        except InspectionCancelled:
            with self._control_lock:
                if self._finish_interrupted_task(task):
                    return False
            raise
        if success and not task['verification']['publish_allowed']:
            success = False
            task['outcome'] = 'verification_failed'
            msg = '处理输出未通过验证，未发布；保存重开：{}；结构预期：{}'.format(
                task['verification']['save_reopen'], task['verification']['structural_expectations'])
        try:
            plan_matches = not (success and task.get('file_plan')) or input_matches_plan(
                task['file_path'], task['file_plan'], cancelled)
        except InspectionCancelled:
            with self._control_lock:
                if self._finish_interrupted_task(task):
                    return False
            raise
        if not plan_matches:
            success = False
            task['outcome'] = 'stale_plan'
            msg = '处理期间输入文件发生变化，未发布；请重新生成计划'
        # Stop/skip requests may arrive while structural facts are being read.
        # Recheck immediately before os.replace so a verified stage cannot be
        # published after the user has cancelled the current task.
        with self._control_lock:
            if self._finish_interrupted_task(task):
                return False
            if not success:
                self._clear_current_skip_state()
        if success:
            try:
                check_unlinked_path(out_path)
                check_unlinked_path(task["work_path"])
                # Path/device checks may block. Recheck under the same lock used
                # by stop/skip, then atomically order the request vs publication.
                with self._control_lock:
                    if self._finish_interrupted_task(task):
                        return False
                    os.replace(task["work_path"], out_path)
                    task['published'] = True
                    self._clear_current_skip_state()
            except Exception as e:
                with self._control_lock:
                    if self._finish_interrupted_task(task):
                        return False
                    self._clear_current_skip_state()
                success = False
                msg = f"提交输出失败（原文件已保留）: {str(e)}"
                task['outcome'] = 'publish_failed'

        task.setdefault('outcome', 'succeeded' if success else 'processing_failed')

        status = "处理完成" if success else "处理失败"
        warning = self._recycle_task_output(task)
        if warning:
            msg = f"{msg}；{warning}"

        now = datetime.now().strftime('%H:%M:%S')
        self.structured_progress.emit({
            "type": "terminal",
            "time": now,
            "file_path": task["file_path"],
            "out_path": out_path,
            "status": status,
            "message": msg,
            "ended_at": timestamp(),
            **self._task_facts(task),
        })
        self.progress.emit(
            task["file_path"],
            status,
            f"[{now}] {task['file_path']}\n    状态: {status}\n    输出文件: {out_path}\n    结果: {msg}"
            f"\n    验证：保存重开 {task['verification']['save_reopen']}；结构预期 {task['verification']['structural_expectations']}；规则级 not_verified",
        )
        return success

    def _emit_stopped_task(self, task):
        self._clear_current_skip_state()
        warning = self._recycle_task_output(task)
        reason = "用户手动停止处理" + (f"；{warning}" if warning else "")
        now = datetime.now().strftime('%H:%M:%S')
        self.structured_progress.emit({
            "type": "terminal",
            "time": now,
            "file_path": task["file_path"],
            "out_path": task["out_path"],
            "status": "已停止",
            "message": reason,
            "ended_at": timestamp(),
            **self._task_facts(dict(task, outcome='stopped')),
        })
        self.progress.emit(
            task["file_path"],
            "已停止",
            f"[{now}] {task['file_path']}\n    状态: 已停止\n    输出文件: {task['out_path']}\n    原因: {reason}",
        )

    def _emit_skipped_task(self, task, reason="已跳过当前文件"):
        self._clear_current_skip_state()
        warning = self._recycle_task_output(task)
        if warning:
            reason = f"{reason}；{warning}"
        now = datetime.now().strftime('%H:%M:%S')
        self.structured_progress.emit({
            "type": "terminal",
            "time": now,
            "file_path": task["file_path"],
            "out_path": task["out_path"],
            "status": "已跳过",
            "message": reason,
            "ended_at": timestamp(),
            **self._task_facts(dict(task, outcome='skipped')),
        })
        self.progress.emit(
            task["file_path"],
            "已跳过",
            f"[{now}] {task['file_path']}\n    状态: 已跳过\n    输出文件: {task['out_path']}\n    原因: {reason}",
        )

    def _run_serial(self):
        task = None
        failure = None
        try:
            started_at = datetime.now()
            success_count = 0
            rename_ectd = "filename_ectd_format" in self.options
            stopped = False

            for i, file_path in enumerate(self.files):
                if self._stop_requested:
                    stopped = True
                    break

                # Keep a retryable identity even if preparation/spawn fails.
                task = {"file_path": file_path, "out_path": ""}
                task["out_path"] = self._build_output_path(i, file_path, rename_ectd)[1]
                task = self._start_process_task(i, file_path, rename_ectd)
                if task.get('rejection'):
                    self._finish_process_task(task, False, task['rejection'])
                    task = None
                    continue
                parent_conn = task["parent_conn"]
                proc = task["proc"]

                success, msg = False, "处理中断"
                skipped_current = False
                timed_out = False
                with self._control_lock:
                    self._can_skip_current = True
                    self._current_file_path = file_path
                while proc.is_alive():
                    if self._stop_requested:
                        stopped = True
                        self._terminate_process_task(task)
                        break

                    if self._skip_requested or path_key(file_path) in self._skip_requested_files:
                        skipped_current = True
                        self._skip_requested = False
                        self._skip_requested_files.discard(path_key(file_path))
                        self._terminate_process_task(task)
                        break

                    if self._task_expired(task):
                        timed_out = True
                        self._terminate_process_task(task)
                        break

                    if parent_conn.poll(0.1):
                        break

                if stopped:
                    self._emit_stopped_task(task)
                    self._close_task_connection(task)
                    task = None
                    break

                if skipped_current:
                    self._emit_skipped_task(task)
                    self._close_task_connection(task)
                    task = None
                    continue

                if timed_out:
                    self._emit_timed_out_task(task)
                    self._close_task_connection(task)
                    task = None
                    continue

                # Drain the pipe before joining: send() may still be blocked on
                # a result larger than the OS pipe buffer even after poll().
                if parent_conn.poll():
                    success, msg = self._consume_engine_result(task, parent_conn.recv())
                else:
                    success, msg = False, "处理进程无返回结果"
                self._join_process_task(task)

                if self._finish_process_task(task, success, msg):
                    success_count += 1
                self._close_task_connection(task)
                task = None

        except Exception as exc:
            failure = str(exc)
        finally:
            self._clear_current_skip_state()
            if task is not None:
                self._fail_process_task(task, f"批处理异常: {failure or '处理中断'}")

        if failure is not None:
            self.error.emit(failure)
        else:
            if stopped:
                summary = f"任务已停止。已成功处理 {success_count} / {len(self.files)} 个文件。"
            else:
                summary = f"处理结束。共成功处理 {success_count} / {len(self.files)} 个文件。"
            elapsed_sec = int((datetime.now() - started_at).total_seconds())
            summary += f" 总耗时 {elapsed_sec}s。"
            self.finished_all.emit(summary)

    def _run_parallel(self):
        running = {}
        starting_task = None
        failure = None
        try:
            started_at = datetime.now()
            success_count = 0
            rename_ectd = "filename_ectd_format" in self.options
            stopped = False
            next_index = 0

            while next_index < len(self.files) or running:
                while not self._stop_requested and next_index < len(self.files) and len(running) < self.effective_max_workers:
                    file_path = self.files[next_index]
                    starting_task = {"file_path": file_path, "out_path": ""}
                    starting_task["out_path"] = self._build_output_path(next_index, file_path, rename_ectd)[1]
                    task = self._start_process_task(next_index, file_path, rename_ectd)
                    if task.get('rejection'):
                        self._finish_process_task(task, False, task['rejection'])
                        starting_task = None
                        next_index += 1
                        continue
                    running[next_index] = task
                    starting_task = None
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
                    self._skip_requested_files.discard(skip_key)
                    for task_id, task in list(running.items()):
                        if path_key(task['file_path']) != skip_key:
                            continue
                        self._terminate_process_task(task)
                        self._emit_skipped_task(task, "用户终止选中文件")
                        self._close_task_connection(task)
                        running.pop(task_id)

                completed_any = False
                for key, task in list(running.items()):
                    parent_conn = task["parent_conn"]
                    proc = task["proc"]
                    if self._task_expired(task):
                        self._terminate_process_task(task)
                        self._emit_timed_out_task(task)
                        self._close_task_connection(task)
                        running.pop(key, None)
                        continue
                    if parent_conn.poll():
                        success, msg = self._consume_engine_result(task, parent_conn.recv())
                    elif not proc.is_alive():
                        success, msg = False, "处理进程无返回结果"
                    else:
                        continue

                    self._join_process_task(task)
                    if self._finish_process_task(task, success, msg):
                        success_count += 1
                    self._close_task_connection(task)
                    running.pop(key, None)
                    completed_any = True

                if running and not completed_any:
                    self.msleep(100)

        except Exception as exc:
            failure = str(exc)
        finally:
            for task in running.values():
                self._fail_process_task(task, f"批处理异常: {failure or '处理中断'}")
            if starting_task is not None:
                self._fail_process_task(starting_task, f"启动处理失败: {failure or '处理中断'}")

        # Only release the UI after child processes, connections and stages have
        # been handled and every started/failed-to-start file has a terminal state.
        if failure is not None:
            self.error.emit(failure)
        else:
            if stopped:
                summary = f"任务已停止。已成功处理 {success_count} / {len(self.files)} 个文件。"
            else:
                summary = f"处理结束。共成功处理 {success_count} / {len(self.files)} 个文件。"
            elapsed_sec = int((datetime.now() - started_at).total_seconds())
            summary += f" 总耗时 {elapsed_sec}s。"
            self.finished_all.emit(summary)

class PreCheckWorker(QThread):
    """
    后台预检线程：只读取 PDF 状态，生成建议勾选的处理项，不修改文件。
    """
    progress = Signal(str, str, str)
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
                    file_path,
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
                        file_path,
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
                        file_path,
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
                        file_path,
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
    progress = Signal(str, str, str)
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
            findings = pdf_inspect.collect_annotation_findings_for_path(file_path)
            hit = bool(findings.get("has_annotations"))
        else:
            findings = pdf_inspect.collect_broken_reference_findings_for_path(file_path)
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
                    file_path,
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
                        file_path,
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
                    file_path,
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
    progress = Signal(str, str, str)
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
        self.common_base = common_base or common_base_dir(self.files)
        self.link_scope = "external" if link_scope == "external" else "all"
        self.link_mode = "incremental" if link_mode == "incremental" else "overwrite"

    def run(self):
        try:
            pairs = []
            for file_path in self.files:
                file_outputs = set()
                for action_type in self.action_types:
                    meta = _io_action_metadata(action_type)
                    data_path, out_pdf = _build_io_paths_for_file(
                        file_path, meta['data_kind'], self.target_dir, self.output_dir, self.common_base,
                    )
                    destination = data_path if meta['is_export'] else out_pdf
                    if not destination:
                        raise ValueError('导入书签或链接时缺少输出目录')
                    # Several import actions intentionally form one pipeline per PDF.
                    if destination not in file_outputs:
                        pairs.append((file_path, destination))
                        file_outputs.add(destination)
            validate_output_plan(pairs, self.files)
            for row_idx, file_path in enumerate(self.files):
                base_name = os.path.basename(file_path)

                self.progress.emit(file_path, "正在执行...",
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
                self.progress.emit(file_path, status, f"   ↳ 结果: {msg}")

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
