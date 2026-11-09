"""Scheduler exceptions reap every child and preserve retryable terminal state."""

import multiprocessing as mp
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

import fitz

from ratools_pdf.controllers.workers import ProcessWorker


LARGE_RESULT = 'page result: checked content, annotations and navigation\n' * 25000


def _write_pdf(path, text):
    with fitz.open() as doc:
        doc.new_page().insert_text((72, 72), text)
        doc.save(path)


def large_result_child(file_path, out_path, options, mode, compression, conn):
    with fitz.open() as doc:
        doc.new_page().insert_text((72, 72), 'Completed large-result task')
        doc.save(out_path)
    try:
        conn.send((True, LARGE_RESULT))
    finally:
        conn.close()


class WorkerLifecycleTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.root = Path(self.directory.name)
        self.output = self.root/'out'
        self.output.mkdir()
        self.files = []
        for name in ('first.pdf', 'second.pdf'):
            source = self.root/name
            _write_pdf(source, 'original')
            (self.output/name).write_bytes(b'earlier output')
            self.files.append(str(source))

    def worker(self, parallel=True):
        worker = ProcessWorker(self.files, set(), str(self.output), max_workers=2 if parallel else 1)
        events, errors, summaries = [], [], []
        worker.structured_progress.connect(events.append)
        worker.error.connect(errors.append)
        worker.finished_all.connect(summaries.append)
        return worker, events, errors, summaries

    def task(self, index, alive=True):
        stage = Path(tempfile.mkdtemp(prefix=f'stage-{index}-', dir=self.root))
        work = stage/'result.pdf'
        _write_pdf(work, 'new complete output')
        proc = Mock()
        proc.is_alive.return_value = alive
        proc.terminate.side_effect = lambda: setattr(proc.is_alive, 'return_value', False)
        proc.kill.side_effect = proc.terminate.side_effect
        conn = Mock()
        conn.poll.return_value = True
        conn.recv.return_value = (True, 'done')
        return dict(index=index, file_path=self.files[index], base_name=Path(self.files[index]).name,
                    out_path=str(self.output/Path(self.files[index]).name), staging_dir=str(stage),
                    work_path=str(work), proc=proc, parent_conn=conn)

    def assert_reaped(self, task):
        task['proc'].join.assert_called()
        self.assertFalse(task['proc'].is_alive())
        task['parent_conn'].close.assert_called()
        self.assertFalse(Path(task['staging_dir']).exists())

    def assert_retryable(self, events, paths):
        terminals = [event for event in events if event['type'] == 'terminal']
        self.assertEqual({event['file_path'] for event in terminals}, set(paths))
        self.assertTrue(all(event['status'] == '处理失败' for event in terminals))
        self.assertEqual(len(terminals), len(paths))

    def test_second_spawn_failure_reaps_first_and_marks_both_retryable(self):
        worker, events, errors, _ = self.worker()
        first = self.task(0)
        with patch.object(worker, '_start_process_task', side_effect=[first, OSError('spawn failed')]):
            worker.run()
        self.assert_reaped(first)
        self.assert_retryable(events, self.files)
        self.assertEqual(len(errors), 1)
        for source in self.files:
            self.assertEqual((self.output/Path(source).name).read_bytes(), b'earlier output')

    def test_poll_and_receive_failure_reap_all_registered_tasks(self):
        for operation in ('poll', 'recv'):
            with self.subTest(operation=operation):
                worker, events, errors, _ = self.worker()
                tasks = [self.task(0), self.task(1)]
                getattr(tasks[0]['parent_conn'], operation).side_effect = OSError('pipe failed')
                with patch.object(worker, '_start_process_task', side_effect=tasks):
                    worker.run()
                for task in tasks:
                    self.assert_reaped(task)
                self.assert_retryable(events, self.files)
                self.assertEqual(len(errors), 1)

    def test_later_communication_failure_does_not_overwrite_prior_success(self):
        worker, events, errors, _ = self.worker()
        tasks = [self.task(0, alive=False), self.task(1)]
        tasks[1]['parent_conn'].poll.side_effect = OSError('pipe failed')
        with patch.object(worker, '_start_process_task', side_effect=tasks):
            worker.run()
        self.assertEqual([event['status'] for event in events], ['处理完成', '处理失败'])
        for task in tasks:
            self.assert_reaped(task)
        with fitz.open(tasks[0]['out_path']) as doc:
            self.assertIn('new complete output', doc[0].get_text())
        self.assertEqual(len(errors), 1)

    def test_publish_failure_closes_connections_and_preserves_previous_outputs(self):
        worker, events, errors, summaries = self.worker()
        tasks = [self.task(0, alive=False), self.task(1, alive=False)]
        with patch.object(worker, '_start_process_task', side_effect=tasks), \
                patch('ratools_pdf.controllers.workers.os.replace', side_effect=PermissionError('locked')):
            worker.run()
        for task in tasks:
            self.assert_reaped(task)
            self.assertEqual(Path(task['out_path']).read_bytes(), b'earlier output')
        self.assert_retryable(events, self.files)
        self.assertEqual(errors, [])
        self.assertEqual(len(summaries), 1)

    def test_serial_communication_failure_reaps_child_and_clears_skip_state(self):
        worker, events, errors, _ = self.worker(parallel=False)
        task = self.task(0)
        task['parent_conn'].poll.side_effect = OSError('poll failed')
        with patch.object(worker, '_start_process_task', return_value=task):
            worker.run()
        self.assert_reaped(task)
        self.assert_retryable(events, [self.files[0]])
        self.assertFalse(worker._can_skip_current)
        self.assertEqual(worker._current_file_path, '')
        self.assertEqual(len(errors), 1)

    def test_serial_start_failure_is_a_retryable_terminal_event(self):
        worker, events, errors, _ = self.worker(parallel=False)
        with patch.object(worker, '_start_process_task', side_effect=OSError('spawn failed')):
            worker.run()
        self.assert_retryable(events, [self.files[0]])
        self.assertEqual(len(errors), 1)

    def test_serial_receives_large_real_pipe_result_before_joining_child(self):
        worker, events, errors, summaries = self.worker(parallel=False)
        worker.files = self.files[:1]
        children = []
        def launch(**kwargs):
            child = mp.get_context('spawn').Process(**kwargs)
            children.append(child)
            return child
        try:
            with patch('ratools_pdf.controllers.workers._process_document_task_pipe', large_result_child), \
                    patch('ratools_pdf.controllers.workers.mp.Process', side_effect=launch):
                worker.run()
            self.assertEqual(errors, [])
            self.assertEqual(len(summaries), 1)
            terminal = [event for event in events if event['type'] == 'terminal']
            self.assertEqual(terminal[0]['message'], LARGE_RESULT)
            self.assertEqual(terminal[0]['status'], '处理完成')
            self.assertEqual(children[0].exitcode, 0)
            with fitz.open(self.output/'first.pdf') as doc:
                self.assertIn('Completed large-result task', doc[0].get_text())
        finally:
            for child in children:
                if child.is_alive():
                    child.terminate()
                child.join(5)


if __name__ == '__main__':
    unittest.main()
