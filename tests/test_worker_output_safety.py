"""A failed/stopped task must never remove results from an earlier batch."""

import os
import multiprocessing as mp
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

import fitz

from ratools_pdf.controllers.workers import ProcessWorker


def _child_cleanup_probe(conn, directory):
    blocked = False
    try:
        sys.audit('os.remove', 'audit-only-no-file', -1)
    except RuntimeError:
        blocked = True
    if blocked:
        path = Path(directory) / 'child-owned.tmp'
        path.write_bytes(b'temporary child data')
        os.remove(path)  # Must be redirected; the audit hook prevents physical deletion.
        conn.send((blocked, not path.exists()))
    else:
        conn.send((False, False))
    conn.close()


class WorkerOutputSafetyTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.root = Path(self.directory.name)
        self.source = self.root / 'source.pdf'
        self.source.write_bytes(b'original document')
        self.output_dir = self.root / 'output'
        self.output_dir.mkdir()
        self.destination = self.output_dir / 'source.pdf'
        self.destination.write_bytes(b'previous successful result')
        self.worker = ProcessWorker([str(self.source)], set(), str(self.output_dir))
        self.events = []
        self.worker.structured_progress.connect(self.events.append)

    def start_task(self):
        # Keep path allocation and output lifecycle real; do not launch a PDF subprocess.
        with patch('ratools_pdf.controllers.workers.mp.Process'), \
                patch('ratools_pdf.controllers.workers.mp.Pipe', return_value=(MagicMock(), MagicMock())):
            return self.worker._start_process_task(0, str(self.source), False)

    def test_failure_before_writing_preserves_existing_result_and_legacy_temp(self):
        legacy = Path(str(self.destination) + '.tmp.pdf')
        legacy.write_bytes(b'not owned by this task')
        task = self.start_task()
        self.assertFalse(self.worker._finish_process_task(task, False, 'input could not be opened'))
        self.assertTrue(self.destination.exists(), 'earlier successful output was removed')
        self.assertEqual(self.destination.read_bytes(), b'previous successful result')
        self.assertEqual(legacy.read_bytes(), b'not owned by this task')

    def test_stop_preserves_existing_result(self):
        self.worker._emit_stopped_task(self.start_task())
        self.assertTrue(self.destination.exists(), 'stop removed an earlier output')
        self.assertEqual(self.destination.read_bytes(), b'previous successful result')

    def test_skip_preserves_existing_result(self):
        self.worker._emit_skipped_task(self.start_task())
        self.assertTrue(self.destination.exists(), 'skip removed an earlier output')
        self.assertEqual(self.destination.read_bytes(), b'previous successful result')

    def test_failure_after_partial_write_does_not_change_published_result(self):
        with patch('ratools_pdf.controllers.workers.mp.Process') as process, \
                patch('ratools_pdf.controllers.workers.mp.Pipe', return_value=(MagicMock(), MagicMock())):
            task = self.worker._start_process_task(0, str(self.source), False)
        work_path = Path(process.call_args.kwargs['args'][1])
        work_path.write_bytes(b'partial new document')
        self.worker._finish_process_task(task, False, 'write failed')
        self.assertTrue(self.destination.exists(), 'failure removed published output')
        self.assertEqual(self.destination.read_bytes(), b'previous successful result')
        self.assertFalse(work_path.exists(), 'owned temporary output was not recycled')

    def test_success_publishes_staged_content_and_reports_final_path(self):
        with patch('ratools_pdf.controllers.workers.mp.Process') as process, \
                patch('ratools_pdf.controllers.workers.mp.Pipe', return_value=(MagicMock(), MagicMock())):
            task = self.worker._start_process_task(0, str(self.source), False)
        work_path = Path(process.call_args.kwargs['args'][1])
        self.assertNotEqual(work_path, self.destination, 'child writes directly to published result')
        work_path.write_bytes(b'complete new document')
        self.assertEqual(self.destination.read_bytes(), b'previous successful result')
        self.assertTrue(self.worker._finish_process_task(task, True, 'done'))
        self.assertEqual(self.destination.read_bytes(), b'complete new document')
        self.assertEqual(self.events[-1]['out_path'], str(self.destination))
        self.assertFalse(work_path.parent.exists(), 'owned temporary directory was not recycled')

    def test_failed_overwrite_commit_preserves_original(self):
        self.worker.overwrite_original = True
        with patch('ratools_pdf.controllers.workers.mp.Process') as process, \
                patch('ratools_pdf.controllers.workers.mp.Pipe', return_value=(MagicMock(), MagicMock())):
            task = self.worker._start_process_task(0, str(self.source), False)
        work_path = Path(process.call_args.kwargs['args'][1])
        work_path.write_bytes(b'new document')
        with patch('ratools_pdf.controllers.workers.os.replace', side_effect=PermissionError('locked')):
            self.assertFalse(self.worker._finish_process_task(task, True, 'done'))
        self.assertEqual(self.source.read_bytes(), b'original document')
        self.assertFalse(work_path.exists())

    def test_without_overwrite_an_input_output_alias_is_rejected(self):
        self.worker.output_dir = str(self.root)
        with self.assertRaises(ValueError):
            self.start_task()
        self.assertEqual(self.source.read_bytes(), b'original document')

    def test_successful_overwrite_replaces_original_only_after_success(self):
        self.worker.overwrite_original = True
        with patch('ratools_pdf.controllers.workers.mp.Process') as process, \
                patch('ratools_pdf.controllers.workers.mp.Pipe', return_value=(MagicMock(), MagicMock())):
            task = self.worker._start_process_task(0, str(self.source), False)
        work_path = Path(process.call_args.kwargs['args'][1])
        work_path.write_bytes(b'complete new document')
        self.assertEqual(self.source.read_bytes(), b'original document')
        self.assertTrue(self.worker._finish_process_task(task, True, 'done'))
        self.assertEqual(self.source.read_bytes(), b'complete new document')
        self.assertEqual(self.events[-1]['out_path'], str(self.source))

    def test_spawn_failure_preserves_old_result_and_recycles_allocated_directory(self):
        with patch('ratools_pdf.controllers.workers.mp.Process') as process, \
                patch('ratools_pdf.controllers.workers.mp.Pipe', return_value=(MagicMock(), MagicMock())):
            process.return_value.start.side_effect = OSError('cannot start process')
            with self.assertRaisesRegex(OSError, 'cannot start process'):
                self.worker._start_process_task(0, str(self.source), False)
        self.assertEqual(self.destination.read_bytes(), b'previous successful result')
        self.assertEqual(list(self.output_dir.iterdir()), [self.destination])

    def test_recycle_failure_retains_stage_and_reports_location(self):
        task = self.start_task()
        stage = Path(task['staging_dir'])
        (stage / 'result.pdf').write_bytes(b'partial document')
        with patch('ratools_pdf.controllers.workers.recycle_owned_path', side_effect=OSError('unavailable')):
            self.worker._finish_process_task(task, False, 'processing failed')
        self.assertEqual(self.destination.read_bytes(), b'previous successful result')
        self.assertEqual((stage / 'result.pdf').read_bytes(), b'partial document')
        self.assertIn(str(stage), self.events[-1]['message'])

    def test_real_serial_pdf_process_publishes_valid_output(self):
        self.check_real_process(max_workers=1)

    def test_real_parallel_pdf_process_publishes_valid_outputs(self):
        self.check_real_process(max_workers=2)

    @unittest.skipUnless(os.environ.get('RATOOLS_SAFE_TEST_ROOT'), 'requires safe test runner')
    def test_spawned_child_blocks_physical_deletion_and_recycles_cleanup(self):
        parent, child = mp.Pipe(duplex=False)
        process = mp.Process(target=_child_cleanup_probe, args=(child, str(self.root)))
        process.start()
        child.close()
        try:
            self.assertTrue(parent.poll(15), 'child safety probe did not return')
            self.assertEqual(parent.recv(), (True, True))
        finally:
            process.join(5)
            if process.is_alive():
                process.terminate()
                process.join(5)
            parent.close()
        self.assertEqual(process.exitcode, 0)

    def test_real_serial_qpdf_process_recycles_intermediate_output(self):
        self.check_real_process(max_workers=1, extra_options={'fast_web_view'})

    def check_real_process(self, max_workers, extra_options=None):
        sources = []
        for index in range(2):
            source = self.root / f'input_{index}.pdf'
            with fitz.open() as doc:
                doc.new_page().insert_text((72, 72), f'Original text {index}')
                doc.save(source)
            (self.output_dir / source.name).write_bytes(b'old output')
            sources.append(str(source))
        worker = ProcessWorker(sources, {'title_from_filename'} | (extra_options or set()), str(self.output_dir),
                               max_workers=max_workers, processing_mode='force')
        errors, summaries = [], []
        worker.error.connect(errors.append)
        worker.finished_all.connect(summaries.append)
        worker.run()
        self.assertEqual(errors, [])
        self.assertEqual(len(summaries), 1)
        for index, source in enumerate(sources):
            with fitz.open(self.output_dir / Path(source).name) as doc:
                self.assertEqual(doc.metadata['title'], f'input_{index}')
                self.assertIn(f'Original text {index}', doc[0].get_text())
            with fitz.open(source) as doc:
                self.assertEqual(doc.metadata['title'], '')
        self.assertFalse(any(p.is_dir() for p in self.output_dir.iterdir()))


if __name__ == '__main__':
    unittest.main()
