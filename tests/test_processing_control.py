"""PD06: real owned trees, per-file deadlines and conservative admission.

Mutations caught: leaving descendants alive, publishing an expired stage,
ignoring skip in serial mode, losing retryable timeout identity, oversubscribing
unknown memory, ignoring low disk, or accepting invalid persisted deadlines.
"""

import ctypes
import importlib
import json
import multiprocessing as mp
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import threading
import time
from types import SimpleNamespace
import unittest
from unittest.mock import patch

import fitz
from PySide6.QtWidgets import QApplication

from ratools_pdf.controllers.workers import ProcessWorker
from ratools_pdf.controllers.main_controller import MainController
from ratools_pdf.pdf import qpdf
from ratools_pdf.services.processing_plan import build_processing_plan
from ratools_pdf.ui.dialogs.settings_dialog import SettingsDialog


def write_pdf(path, text='complete PDF'):
    with fitz.open() as doc:
        doc.new_page().insert_text((72, 72), text)
        doc.save(path)


def slow_pdf_child(file_path, out_path, options, mode, compression, conn):
    write_pdf(out_path, 'staged but not published')
    time.sleep(3)
    conn.send((True, 'done'))
    conn.close()


def crashed_qpdf_owner(script, marker):
    with patch.object(qpdf, 'get_qpdf_path', return_value=sys.executable):
        qpdf.run_qpdf_command(['-B', script, marker], timeout=30)


def crash_owner_during_post_create(script, marker, child_pid_path):
    from ratools_pdf.services import process_control
    real_popen = subprocess.Popen
    def launch(*args, **kwargs):
        process = real_popen(*args, **kwargs)
        # A reviewer reproduced owner death in the post-create assignment gap.
        # The new atomic launcher exposes the same observation after CreateProcess
        # for this test by wrapping its real CreateProcess boundary below.
        Path(child_pid_path).write_text(str(process.pid))
        os._exit(37)
    if hasattr(process_control, '_create_windows_process'):
        real_create = process_control._create_windows_process
        def create(*args, **kwargs):
            created = real_create(*args, **kwargs)
            Path(child_pid_path).write_text(str(created.dwProcessId))
            os._exit(37)
        boundary = patch.object(process_control, '_create_windows_process', side_effect=create)
    else:
        boundary = patch.object(process_control.subprocess, 'Popen', side_effect=launch)
    with boundary, patch.object(qpdf, 'get_qpdf_path', return_value=sys.executable):
        qpdf.run_qpdf_command(['-B', script, marker], timeout=30)


def qpdf_tree_pdf_child(file_path, out_path, options, mode, compression, conn):
    root = Path(file_path).parent
    write_pdf(out_path, 'staged output before cancellation')
    with patch.object(qpdf, 'get_qpdf_path', return_value=sys.executable):
        qpdf.run_qpdf_command(['-B', str(root / 'tree.py'), str(root / 'tree.txt')], timeout=30)
    conn.send((True, 'done'))
    conn.close()


def process_alive(pid):
    kernel = ctypes.WinDLL('kernel32', use_last_error=True)
    kernel.OpenProcess.restype = ctypes.c_void_p
    handle = kernel.OpenProcess(0x100000 | 0x1000, False, pid)
    if not handle:
        return False
    try:
        kernel.WaitForSingleObject.argtypes = [ctypes.c_void_p, ctypes.c_ulong]
        return kernel.WaitForSingleObject(handle, 0) == 258
    finally:
        kernel.CloseHandle.argtypes = [ctypes.c_void_p]
        kernel.CloseHandle(handle)


def terminate_test_pid(pid):
    kernel = ctypes.WinDLL('kernel32', use_last_error=True)
    kernel.OpenProcess.restype = ctypes.c_void_p
    handle = kernel.OpenProcess(0x100000 | 1, False, pid)
    if handle:
        kernel.TerminateProcess.argtypes = [ctypes.c_void_p, ctypes.c_uint]
        kernel.WaitForSingleObject.argtypes = [ctypes.c_void_p, ctypes.c_ulong]
        kernel.CloseHandle.argtypes = [ctypes.c_void_p]
        kernel.TerminateProcess(handle, 1)
        kernel.WaitForSingleObject(handle, 3000)
        kernel.CloseHandle(handle)


class ProcessingControlTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.root = Path(self.directory.name)
        self.source = self.root / 'source.pdf'
        write_pdf(self.source, 'original')
        self.output = self.root / 'out'
        self.output.mkdir()
        self.destination = self.output / 'source.pdf'
        self.destination.write_bytes(b'previous successful output')

    def service(self):
        try:
            return importlib.import_module('ratools_pdf.services.process_control')
        except ImportError:
            self.fail('PD06 process and resource control service is missing')

    def worker(self, parallel=False):
        worker = ProcessWorker([str(self.source)], set(), str(self.output),
                               max_workers=2 if parallel else 1)
        worker.file_timeout_seconds = 0.5
        events, errors = [], []
        worker.structured_progress.connect(events.append)
        worker.error.connect(errors.append)
        return worker, events, errors

    def test_serial_and_parallel_timeout_preserve_old_output_and_can_retry(self):
        for parallel in (False, True):
            with self.subTest(parallel=parallel):
                worker, events, errors = self.worker(parallel)
                with patch('ratools_pdf.controllers.workers._process_document_task_pipe', slow_pdf_child):
                    worker.run()
                terminals = [event for event in events if event['type'] == 'terminal']
                self.assertEqual([event['outcome'] for event in terminals], ['timed_out'])
                self.assertFalse(terminals[0]['published'])
                self.assertEqual(terminals[0]['status'], '处理失败')
                self.assertIn('超时', terminals[0]['message'])
                self.assertEqual(self.destination.read_bytes(), b'previous successful output')
                self.assertEqual(errors, [])
                worker.file_timeout_seconds = 0
                worker.run()
                self.assertEqual(events[-1]['outcome'], 'succeeded')
                with fitz.open(self.destination) as doc:
                    self.assertIn('original', doc[0].get_text())
                self.destination.write_bytes(b'previous successful output')

    def test_timeout_during_output_verification_never_publishes(self):
        worker, events, _ = self.worker()
        from ratools_pdf.services.pdf_facts import verify_output
        def expire(*args):
            result = verify_output(*args)
            time.sleep(0.2)
            return result
        # Real PDF processing succeeds; only the verification phase crosses expiry.
        worker.file_timeout_seconds = 10
        real_finish = worker._finish_process_task
        def finish(task, *args):
            task['deadline'] = time.monotonic() + 0.05
            return real_finish(task, *args)
        with patch.object(worker, '_finish_process_task', side_effect=finish), \
                patch('ratools_pdf.controllers.workers.verify_output', side_effect=expire):
            worker.run()
        self.assertEqual(events[-1]['outcome'], 'timed_out')
        self.assertEqual(self.destination.read_bytes(), b'previous successful output')

    def test_control_request_during_final_path_check_prevents_publication(self):
        from ratools_pdf.controllers.workers import check_unlinked_path
        for action, expected in (('stop', 'stopped'), ('skip', 'skipped'), ('expire', 'timed_out')):
            with self.subTest(action=action):
                worker, events, _ = self.worker()
                worker.file_timeout_seconds = 0
                current_task = {}
                original_finish = worker._finish_process_task
                def finish(task, *args):
                    current_task.update(task=task)
                    return original_finish(task, *args)
                def check(path):
                    result = check_unlinked_path(path)
                    if Path(path).name == 'result.pdf':
                        if action == 'stop':
                            worker.request_stop()
                        elif action == 'skip':
                            worker.request_skip_file(str(self.source))
                        else:
                            current_task['task']['deadline'] = time.monotonic() - 1
                    return result
                with patch.object(worker, '_finish_process_task', side_effect=finish), \
                        patch('ratools_pdf.controllers.workers.check_unlinked_path', side_effect=check):
                    worker.run()
                self.assertEqual(events[-1]['outcome'], expected)
                self.assertFalse(events[-1]['published'])
                self.assertEqual(self.destination.read_bytes(), b'previous successful output')

    def test_serial_skip_by_path_remains_distinct_from_timeout(self):
        worker, events, _ = self.worker()
        worker.request_skip_file(str(self.source))
        worker.run()
        self.assertEqual(events[-1]['outcome'], 'skipped')
        self.assertEqual(self.destination.read_bytes(), b'previous successful output')

    def test_skip_current_during_final_fingerprint_preserves_output_and_next_file(self):
        from ratools_pdf.services import processing_plan
        next_source = self.root / 'next.pdf'
        write_pdf(next_source, 'next original')
        worker, events, errors = self.worker()
        worker.file_timeout_seconds = 0
        worker.files.append(str(next_source))
        worker.processing_plan = build_processing_plan(worker.files, set(), common_base=str(self.root))
        worker.file_plans = {item['file_id']: item for item in worker.processing_plan['files']}
        source_fingerprints = []
        real_fingerprint = processing_plan.fingerprint

        def fingerprint(path, cancelled=None):
            if Path(path) == self.source:
                source_fingerprints.append(path)
                if len(source_fingerprints) == 2:
                    worker.request_skip_current()
            return real_fingerprint(path, cancelled)

        def request_after_terminal(event):
            if event['type'] == 'terminal' and event['file_path'] == str(self.source):
                worker.request_skip_current()

        worker.structured_progress.connect(request_after_terminal)
        with patch.object(processing_plan, 'fingerprint', side_effect=fingerprint):
            worker.run()
        terminals = [event for event in events if event['type'] == 'terminal']
        self.assertEqual(errors, [])
        self.assertEqual(len(source_fingerprints), 2)
        self.assertEqual([event['outcome'] for event in terminals], ['skipped', 'succeeded'])
        self.assertFalse(terminals[0]['published'])
        self.assertEqual(self.destination.read_bytes(), b'previous successful output')
        self.assertTrue(terminals[1]['published'])
        with fitz.open(self.output / 'next.pdf') as doc:
            self.assertIn('next original', doc[0].get_text())

    def test_skip_current_during_final_path_check_preserves_output_and_next_file(self):
        from ratools_pdf.controllers.workers import check_unlinked_path
        next_source = self.root / 'next.pdf'
        write_pdf(next_source, 'next original')
        worker, events, errors = self.worker()
        worker.file_timeout_seconds = 0
        worker.files.append(str(next_source))
        requests = []

        def check(path):
            result = check_unlinked_path(path)
            if Path(path).name == 'result.pdf' and not requests:
                requests.append(path)
                worker.request_skip_current()
            return result

        def request_after_terminal(event):
            if event['type'] == 'terminal' and event['file_path'] == str(self.source):
                worker.request_skip_current()

        worker.structured_progress.connect(request_after_terminal)
        with patch('ratools_pdf.controllers.workers.check_unlinked_path', side_effect=check):
            worker.run()
        terminals = [event for event in events if event['type'] == 'terminal']
        self.assertEqual(errors, [])
        self.assertEqual(len(requests), 1)
        self.assertEqual([event['outcome'] for event in terminals], ['skipped', 'succeeded'])
        self.assertFalse(terminals[0]['published'])
        self.assertEqual(self.destination.read_bytes(), b'previous successful output')
        self.assertTrue(terminals[1]['published'])
        with fitz.open(self.output / 'next.pdf') as doc:
            self.assertIn('next original', doc[0].get_text())

    def test_disabled_deadline_allows_successful_real_pdf(self):
        worker, events, errors = self.worker()
        worker.file_timeout_seconds = 0
        worker.run()
        self.assertEqual(errors, [])
        self.assertEqual(events[-1]['outcome'], 'succeeded')
        with fitz.open(self.destination) as doc:
            self.assertIn('original', doc[0].get_text())

    def test_memory_and_input_size_lower_user_limit_and_unknown_falls_back(self):
        service = self.service()
        cases = [(4, [1], 4 * 1024**3, 4),
                 (4, [128 * 1024**2], 4 * 1024**3, 2),
                 (4, [1], 512 * 1024**2, 1),
                 (4, [1], None, 1), (1, [1], 8 * 1024**3, 1)]
        for requested, sizes, available, expected in cases:
            with self.subTest(available=available, sizes=sizes):
                self.assertEqual(service.concurrency_limit(requested, sizes, available), expected)

    def test_scheduler_uses_available_memory_cap(self):
        service = self.service()
        worker, _, _ = self.worker(parallel=True)
        with patch.object(service, 'available_memory_bytes', return_value=None), \
                patch.object(worker, '_run_serial') as serial, \
                patch.object(worker, '_run_parallel') as parallel:
            worker.run()
        self.assertEqual(worker.effective_max_workers, 1)
        serial.assert_called_once()
        parallel.assert_not_called()

    def test_low_disk_and_unavailable_output_preflight_preserve_old_output(self):
        service = self.service()
        for usage in (type('Usage', (), {'free': 1})(), OSError('network unavailable')):
            with self.subTest(usage=type(usage).__name__):
                worker, events, _ = self.worker()
                with patch.object(service.shutil, 'disk_usage', side_effect=usage if isinstance(usage, Exception) else None,
                                  return_value=usage):
                    worker.run()
                self.assertEqual(events[-1]['outcome'], 'resource_preflight_failed')
                self.assertFalse(events[-1]['published'])
                self.assertEqual(self.destination.read_bytes(), b'previous successful output')
                self.assertTrue('空间' in events[-1]['message'] or '输出' in events[-1]['message'])

    def test_timeout_validation_rejects_negative_infinite_or_nonnumeric(self):
        service = self.service()
        for value in (-1, float('nan'), float('inf'), 'invalid', 86401):
            with self.subTest(value=value):
                with self.assertRaises(ValueError):
                    service.normalize_file_timeout(value)
        self.assertEqual(service.normalize_file_timeout(0), 0)
        self.assertEqual(service.normalize_file_timeout('120'), 120)

    def test_snapshot_and_plan_hash_cancel_at_existing_cooperative_checkpoints(self):
        from ratools_pdf.services.pdf_facts import InspectionCancelled, capture_pdf_snapshot
        plan = build_processing_plan([str(self.source)], set(), common_base=str(self.root))
        for phase in ('before', 'plan_hash', 'after'):
            with self.subTest(phase=phase):
                worker, events, _ = self.worker()
                worker.file_timeout_seconds = 0.05
                checkpoints = []
                if phase == 'after':
                    worker.file_timeout_seconds = 0
                    original_finish = worker._finish_process_task
                    def finish(task, *args):
                        task['deadline'] = time.monotonic() + 0.05
                        return original_finish(task, *args)
                    worker._finish_process_task = finish
                if phase == 'plan_hash':
                    worker.processing_plan = plan
                    worker.file_plans = {item['file_id']: item for item in plan['files']}
                real_fingerprint = importlib.import_module('ratools_pdf.services.processing_plan').fingerprint
                def slow_checkpoint(path, cancelled=None):
                    # Model the actual per-page / 1MiB hash checkpoints without
                    # claiming to interrupt any individual OS or fitz call.
                    for _ in range(30):
                        time.sleep(0.02)
                        checkpoints.append(1)
                        if cancelled and cancelled():
                            raise InspectionCancelled()
                    if phase == 'plan_hash':
                        return real_fingerprint(path, cancelled)
                    return capture_pdf_snapshot(path, cancelled)
                if phase == 'plan_hash':
                    boundary = patch('ratools_pdf.services.processing_plan.fingerprint', side_effect=slow_checkpoint)
                else:
                    def snapshot(path, cancelled=None):
                        if phase == 'before' or Path(path).name == 'result.pdf':
                            return slow_checkpoint(path, cancelled)
                        return capture_pdf_snapshot(path, cancelled)
                    boundary = patch('ratools_pdf.controllers.workers.capture_pdf_snapshot', side_effect=snapshot)
                start = time.monotonic()
                with boundary:
                    worker.run()
                elapsed = time.monotonic() - start
                self.assertEqual(events[-1]['outcome'], 'timed_out')
                self.assertGreater(len(checkpoints), 0)
                self.assertLess(len(checkpoints), 10, f'cooperative checkpoints ignored ({elapsed:.3f}s)')
                self.assertEqual(self.destination.read_bytes(), b'previous successful output')

    def test_audit_timeout_exports_distinct_stable_reason(self):
        from ratools_pdf.services.audit_evidence import export_evidence_package
        host = SimpleNamespace(process_log_starts={}, process_log_rows=[])
        plan = build_processing_plan([str(self.source)], set(), common_base=str(self.root))
        MainController._retain_processing_plan(host, plan, 'running')
        MainController._finish_processing_evidence(host, 'finished')
        host.processing_evidence['files'][plan['files'][0]['file_id']]['outcome'] = 'timed_out'
        export_evidence_package(host.processing_evidence, self.root / 'audit')
        data = json.loads((self.root / 'audit' / 'evidence.json').read_text(encoding='utf-8'))
        self.assertIn('file_timeout', data['files'][0]['reason_codes'])
        self.assertNotIn('engine_failure', data['files'][0]['reason_codes'])


@unittest.skipUnless(os.name == 'nt', 'Windows Job Object qualification')
class OwnedTreeTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.root = Path(self.directory.name)
        self.marker = self.root / 'tree.txt'
        self.script = self.root / 'tree.py'
        self.script.write_text(
            'import os, subprocess, sys, time\n'
            'from pathlib import Path\n'
            'child = subprocess.Popen([sys.executable, "-B", "-c", "import time; time.sleep(30)"], '
            'creationflags=subprocess.CREATE_NO_WINDOW)\n'
            'Path(sys.argv[1]).write_text(str(os.getpid()) + "," + str(child.pid))\n'
            'time.sleep(30)\n', encoding='utf-8')
        self.pids = []
        self.addCleanup(self.cleanup_pids)

    def cleanup_pids(self):
        if self.marker.exists():
            self.pids.extend(int(pid) for pid in self.marker.read_text().split(','))
        for pid in set(self.pids):
            terminate_test_pid(pid)

    def wait_marker(self):
        deadline = time.monotonic() + 10
        while not self.marker.exists() and time.monotonic() < deadline:
            time.sleep(0.02)
        self.assertTrue(self.marker.exists(), 'controlled process tree did not start')
        self.pids = [int(pid) for pid in self.marker.read_text().split(',')]

    def assert_tree_gone(self):
        deadline = time.monotonic() + 5
        while any(process_alive(pid) for pid in self.pids) and time.monotonic() < deadline:
            time.sleep(0.02)
        self.assertEqual([pid for pid in self.pids if process_alive(pid)], [])

    def test_real_timeout_reaps_parent_and_descendant_keeps_unrelated_process(self):
        unrelated = subprocess.Popen([sys.executable, '-B', '-c', 'import time; time.sleep(30)'],
                                     creationflags=subprocess.CREATE_NO_WINDOW)
        self.addCleanup(lambda: (unrelated.terminate(), unrelated.wait(5)))
        with patch.object(qpdf, 'get_qpdf_path', return_value=sys.executable):
            with self.assertRaises(subprocess.TimeoutExpired):
                qpdf.run_qpdf_command(['-B', str(self.script), str(self.marker)], timeout=1)
        self.wait_marker()
        self.assert_tree_gone()
        self.assertIsNone(unrelated.poll())

    def test_python_owner_crash_reaps_qpdf_and_descendant(self):
        owner = mp.get_context('spawn').Process(target=crashed_qpdf_owner,
                                               args=(str(self.script), str(self.marker)))
        owner.start()
        try:
            self.wait_marker()
            self.assertTrue(all(process_alive(pid) for pid in self.pids))
            owner.terminate()
            owner.join(5)
            self.assert_tree_gone()
        finally:
            if owner.is_alive():
                owner.terminate()
            owner.join(5)

    def test_owner_dies_immediately_after_creation_before_python_setup(self):
        child_pid_path = self.root / 'created-pid.txt'
        owner = mp.get_context('spawn').Process(target=crash_owner_during_post_create,
                    args=(str(self.script), str(self.marker), str(child_pid_path)))
        owner.start()
        try:
            owner.join(10)
            self.assertEqual(owner.exitcode, 37)
            self.assertTrue(child_pid_path.exists())
            self.pids = [int(child_pid_path.read_text())]
            self.assert_tree_gone()
        finally:
            if owner.is_alive():
                owner.terminate()
            owner.join(5)

    def test_failed_job_assignment_never_runs_suspended_command(self):
        try:
            service = importlib.import_module('ratools_pdf.services.process_control')
        except ImportError:
            self.fail('PD06 owned process launcher is missing')
        with patch.object(service.WindowsJob, 'assign', side_effect=OSError('nested job refused')), \
                patch.object(qpdf, 'get_qpdf_path', return_value=sys.executable):
            with self.assertRaisesRegex(OSError, 'nested job refused'):
                qpdf.run_qpdf_command(['-B', str(self.script), str(self.marker)], timeout=3)
        self.assertFalse(self.marker.exists())

    def test_failed_resume_reaps_suspended_parent_without_running_code(self):
        try:
            service = importlib.import_module('ratools_pdf.services.process_control')
        except ImportError:
            self.fail('PD06 owned process launcher is missing')
        with patch.object(service.WindowsJob, 'resume', side_effect=OSError('resume refused')), \
                patch.object(qpdf, 'get_qpdf_path', return_value=sys.executable):
            with self.assertRaisesRegex(OSError, 'resume refused'):
                qpdf.run_qpdf_command(['-B', str(self.script), str(self.marker)], timeout=3)
        self.assertFalse(self.marker.exists())

    def test_worker_stop_and_skip_reap_real_qpdf_tree_preserve_old_output(self):
        source = self.root / 'source.pdf'
        write_pdf(source, 'original')
        output = self.root / 'out'
        output.mkdir()
        destination = output / 'source.pdf'
        for parallel, action, expected in ((False, 'stop', 'stopped'),
                                           (False, 'skip', 'skipped'),
                                           (True, 'skip', 'skipped')):
            with self.subTest(parallel=parallel, action=action):
                destination.write_bytes(b'previous output')
                if self.marker.exists():
                    self.marker.rename(self.root / f'previous-{parallel}-{action}.txt')
                worker = ProcessWorker([str(source)], set(), str(output), max_workers=2 if parallel else 1)
                events, failure = [], []
                worker.structured_progress.connect(events.append)
                def cancel():
                    try:
                        self.wait_marker()
                        if action == 'stop':
                            worker.request_stop()
                        else:
                            worker.request_skip_file(str(source))
                    except Exception as exc:
                        failure.append(exc)
                        worker.request_stop()
                watcher = threading.Thread(target=cancel, daemon=True)
                with patch('ratools_pdf.controllers.workers._process_document_task_pipe', qpdf_tree_pdf_child):
                    watcher.start()
                    worker.run()
                watcher.join(12)
                self.assertFalse(watcher.is_alive())
                self.assertEqual(failure, [])
                self.assertEqual(events[-1]['outcome'], expected)
                self.assert_tree_gone()
                self.assertEqual(destination.read_bytes(), b'previous output')
                self.assertFalse(any(path.is_dir() for path in output.iterdir()))

    def test_qpdf_public_result_contract_and_launch_failure(self):
        result = qpdf.run_qpdf_command(['--version'], timeout=5)
        self.assertIsInstance(result, qpdf.QpdfResult)
        self.assertTrue(result.is_success)
        self.assertIn('qpdf', result.stdout)
        with patch.object(qpdf, 'get_qpdf_path', return_value=str(self.root / 'missing.exe')):
            with self.assertRaises(FileNotFoundError):
                qpdf.run_qpdf_command([], timeout=1)


class TimeoutUITests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def test_user_can_disable_or_set_per_file_timeout_with_bounded_values(self):
        dialog = SettingsDialog()
        self.addCleanup(dialog.close)
        self.assertTrue(hasattr(dialog, 'spin_file_timeout'), 'per-file timeout control is missing')
        dialog.spin_file_timeout.setValue(0)
        self.assertEqual(dialog.spin_file_timeout.value(), 0)
        self.assertIn('关闭', dialog.spin_file_timeout.text())
        dialog.spin_file_timeout.setValue(120)
        self.assertEqual(dialog.spin_file_timeout.value(), 120)
        dialog.spin_file_timeout.setValue(-1)
        self.assertEqual(dialog.spin_file_timeout.value(), 0)

    def test_timeout_persists_and_controller_consumes_the_user_setting(self):
        from PySide6.QtCore import QSettings
        from ratools_pdf.ui.main_window import MainWindow
        with tempfile.TemporaryDirectory() as directory:
            settings = QSettings(str(Path(directory) / 'settings.ini'), QSettings.IniFormat)
            settings.setValue('Settings/FileTimeoutSeconds', 120)
            settings.setValue('Settings/ParallelProcessingEnabled', True)
            settings.sync()
            with patch('ratools_pdf.ui.main_window.get_app_dir', return_value=directory):
                window = MainWindow()
            try:
                self.assertTrue(hasattr(window.settings_dialog, 'spin_file_timeout'), 'timeout setting is missing')
                self.assertEqual(window.settings_dialog.spin_file_timeout.value(), 120)
                host = SimpleNamespace(view=window)
                self.assertEqual(MainController._get_file_timeout_seconds(host), 120)
                window.settings_dialog.spin_file_timeout.setValue(0)
                window.persist_all_settings()
                self.assertEqual(int(window.app_settings.value('Settings/FileTimeoutSeconds')), 0)
            finally:
                window.close()

    def test_invalid_saved_timeout_loads_as_disabled(self):
        from PySide6.QtCore import QSettings
        from ratools_pdf.ui.main_window import MainWindow
        with tempfile.TemporaryDirectory() as directory:
            settings = QSettings(str(Path(directory) / 'settings.ini'), QSettings.IniFormat)
            settings.setValue('Settings/FileTimeoutSeconds', 'nan')
            settings.sync()
            with patch('ratools_pdf.ui.main_window.get_app_dir', return_value=directory):
                window = MainWindow()
            try:
                self.assertTrue(hasattr(window.settings_dialog, 'spin_file_timeout'), 'timeout setting is missing')
                self.assertEqual(window.settings_dialog.spin_file_timeout.value(), 0)
            finally:
                window.close()


if __name__ == '__main__':
    unittest.main()
