"""Distinct sources must not silently share a destination or overwrite queued inputs."""

import os
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

import fitz

from ratools_pdf.controllers.io_actions import _build_io_paths_for_file
from ratools_pdf.controllers.workers import IOActionWorker, ProcessWorker
from ratools_pdf.pdf import qpdf
from ratools_pdf.services.file_safety import check_unlinked_path
from ratools_pdf.ui.dialogs.page_assembly_dialog import PageAssemblyWorker
from ratools_pdf.ui.dialogs.security_center_dialog import DecryptWorker


class OutputCollisionTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.root = Path(self.directory.name)
        self.output = self.root / 'out'
        self.output.mkdir()

    def pdf(self, relative, text='Original'):
        path = self.root / relative
        path.parent.mkdir(parents=True, exist_ok=True)
        with fitz.open() as doc:
            doc.new_page().insert_text((72, 72), text)
            doc.set_toc([[1, text, 1]])
            doc.save(path)
        return str(path)

    @unittest.skipUnless(os.name == 'nt', 'Windows drive path semantics')
    def test_cross_drive_batch_outputs_keep_distinct_sources(self):
        sources = [r'C:\one\report.pdf', r'D:\two\report.pdf']
        worker = ProcessWorker(sources, set(), str(self.output))
        paths = [worker._build_output_path(i, path, False)[1] for i, path in enumerate(sources)]
        self.assertNotEqual(os.path.normcase(paths[0]), os.path.normcase(paths[1]))
        for path in paths:
            self.assertTrue(Path(path).is_relative_to(self.output))

    @unittest.skipUnless(os.name == 'nt', 'Windows drive path semantics')
    def test_cross_drive_io_data_and_pdf_paths_keep_distinct_sources(self):
        paths = [_build_io_paths_for_file(source, 'links', str(self.output), str(self.output))
                 for source in [r'C:\one\report.pdf', r'D:\two\report.pdf']]
        self.assertNotEqual(paths[0][0], paths[1][0])
        self.assertNotEqual(paths[0][1], paths[1][1])

    def test_filename_normalization_collision_blocks_batch_before_writing(self):
        sources = [self.pdf('A B.pdf'), self.pdf('a-b.pdf')]
        sentinel = self.output / 'a-b.pdf'
        sentinel.write_bytes(b'previous result')
        worker = ProcessWorker(sources, {'filename_ectd_format'}, str(self.output))
        errors = []
        worker.error.connect(errors.append)
        with patch.object(worker, '_run_serial', side_effect=lambda: sentinel.write_bytes(b'engine started')):
            worker.run()
        self.assertEqual(sentinel.read_bytes(), b'previous result')
        self.assertEqual(len(errors), 1)

    def test_output_must_not_replace_another_queued_input(self):
        sources = [self.pdf('source.pdf'), self.pdf('source_decrypted.pdf')]
        worker = DecryptWorker(list(enumerate(sources)), str(self.root))
        outcomes = []
        worker.finished_all.connect(lambda *args: outcomes.append(args))
        with patch('ratools_pdf.ui.dialogs.security_center_dialog.qpdf.decrypt_pdf') as engine:
            engine.return_value = SimpleNamespace(is_success=True)
            worker.run()
        engine.assert_not_called()  # Boundary contract: reject the whole plan before engine writes.
        self.assertEqual(outcomes[0][:2], (0, 2))

    def test_io_duplicate_targets_rejected_before_first_export(self):
        source = self.pdf('source.pdf')
        sentinel = self.output / 'source_bookmarks.csv'
        sentinel.write_bytes(b'previous export')
        worker = IOActionWorker('export_bookmarks', [source, source], str(self.output))
        errors = []
        worker.error_action.connect(errors.append)
        worker.run()
        self.assertEqual(sentinel.read_bytes(), b'previous export')
        self.assertEqual(len(errors), 1)

    def test_split_preserves_same_named_sources_in_different_folders(self):
        sources = [self.pdf('a/report.pdf', 'Source A'), self.pdf('b/report.pdf', 'Source B')]
        worker = PageAssemblyWorker('split', [qpdf.PageSpec(path) for path in sources], str(self.output))
        results = []
        worker.finished_signal.connect(lambda *args: results.append(args))
        worker.run()
        self.assertTrue(results[0][0], results[0][1])
        outputs = list(self.output.rglob('*.pdf'))
        self.assertEqual(len(outputs), 2, 'one source overwrote another split result')
        texts = []
        for path in outputs:
            with fitz.open(path) as doc:
                texts.append(doc[0].get_text().strip())
        self.assertEqual(sorted(texts), ['Source A', 'Source B'])

    def test_decryption_preserves_same_named_sources_in_different_folders(self):
        sources = [self.pdf('a/report.pdf', 'Source A'), self.pdf('b/report.pdf', 'Source B')]
        worker = DecryptWorker(list(enumerate(sources)), str(self.output))
        results = []
        worker.finished_all.connect(lambda *args: results.append(args))
        worker.run()
        self.assertEqual(results[0][:2], (2, 0))
        outputs = list(self.output.rglob('*.pdf'))
        self.assertEqual(len(outputs), 2, 'one source overwrote another decrypted result')
        texts = []
        for path in outputs:
            with fitz.open(path) as doc:
                texts.append(doc[0].get_text().strip())
        self.assertEqual(sorted(texts), ['Source A', 'Source B'])

    def test_split_pattern_must_not_replace_a_queued_input(self):
        sources = [self.pdf('report.pdf'), self.pdf('report_part_1.pdf')]
        worker = PageAssemblyWorker('split', [qpdf.PageSpec(path) for path in sources], str(self.root))
        results = []
        worker.finished_signal.connect(lambda *args: results.append(args))
        with patch('ratools_pdf.ui.dialogs.page_assembly_dialog.qpdf.split_pages') as engine:
            worker.run()
        engine.assert_not_called()
        self.assertFalse(results[0][0])

    def test_split_audits_existing_concrete_outputs_before_engine_runs(self):
        source = self.pdf('report.pdf')
        concrete = self.output / 'report_part_01-03.pdf'
        concrete.write_bytes(b'existing result')
        worker = PageAssemblyWorker('split', [qpdf.PageSpec(source)], str(self.output), split_pages_per_file=3)
        results = []
        worker.finished_signal.connect(lambda *args: results.append(args))

        def check_path(path):
            # Model a linked output without creating/moving a real link in the test.
            if os.path.normcase(os.path.abspath(path)) == os.path.normcase(str(concrete)):
                raise ValueError('output requires link review')
            return check_unlinked_path(path)

        with patch('ratools_pdf.services.output_paths.check_unlinked_path', side_effect=check_path), \
                patch('ratools_pdf.ui.dialogs.page_assembly_dialog.qpdf.split_pages') as engine:
            worker.run()
        engine.assert_not_called()
        self.assertFalse(results[0][0])
        self.assertEqual(concrete.read_bytes(), b'existing result')

    def test_case_alias_targets_rejected_before_export(self):
        if os.name != 'nt':
            self.skipTest('Windows case-insensitive paths')
        source = self.pdf('source.pdf')
        sentinel = self.output / 'source_links.json'
        sentinel.write_bytes(b'previous export')
        worker = IOActionWorker('export_links', [source, source.upper()], str(self.output))
        errors = []
        worker.error_action.connect(errors.append)
        worker.run()
        self.assertEqual(sentinel.read_bytes(), b'previous export')
        self.assertEqual(len(errors), 1)

    def test_equal_normalized_names_in_distinct_folders_remain_valid(self):
        sources = [self.pdf('a/A B.pdf', 'Source A'), self.pdf('b/a-b.pdf', 'Source B')]
        worker = ProcessWorker(sources, {'filename_ectd_format'}, str(self.output), processing_mode='force')
        errors = []
        worker.error.connect(errors.append)
        worker.run()
        self.assertEqual(errors, [])
        for folder, text in [('a', 'Source A'), ('b', 'Source B')]:
            with fitz.open(self.output / folder / 'a-b.pdf') as doc:
                self.assertIn(text, doc[0].get_text())

    def test_batch_rejects_output_that_is_another_input(self):
        sources = [self.pdf('report.pdf'), self.pdf('out/report.pdf')]
        original = (self.output / 'report.pdf').read_bytes()
        worker = ProcessWorker(sources, {'title_from_filename'}, str(self.output))
        errors = []
        worker.error.connect(errors.append)
        with patch.object(worker, '_run_serial', side_effect=lambda: (self.output / 'report.pdf').write_bytes(b'changed')):
            worker.run()
        self.assertEqual((self.output / 'report.pdf').read_bytes(), original)
        self.assertEqual(len(errors), 1)

    def test_io_same_named_sources_roundtrip_in_matching_subdirectories(self):
        sources = [self.pdf('a/report.pdf', 'Source A'), self.pdf('b/report.pdf', 'Source B')]
        exporter = IOActionWorker('export_bookmarks', sources, str(self.output))
        errors = []
        exporter.error_action.connect(errors.append)
        exporter.run()
        self.assertEqual(errors, [])
        self.assertEqual(len(list(self.output.rglob('*.csv'))), 2)
        for source in sources:
            with fitz.open(source) as doc:
                doc.set_toc([])
                data = doc.tobytes()
            Path(source).write_bytes(data)
        restored = self.root / 'restored'
        importer = IOActionWorker('import_bookmarks', sources, str(self.output), str(restored))
        importer.error_action.connect(errors.append)
        importer.run()
        self.assertEqual(errors, [])
        for folder, title in [('a', 'Source A'), ('b', 'Source B')]:
            with fitz.open(restored / folder / 'report.pdf') as doc:
                self.assertEqual(doc.get_toc()[0][1], title)

    def test_explicit_original_name_decryption_can_replace_its_own_input(self):
        source = self.pdf('report.pdf', 'Keep this text')
        worker = DecryptWorker([(0, source)], str(self.root), suffix='')
        results = []
        worker.finished_all.connect(lambda *args: results.append(args))
        worker.run()
        self.assertEqual(results[0][:2], (1, 0))
        with fitz.open(source) as doc:
            self.assertIn('Keep this text', doc[0].get_text())
        self.assertFalse(any(path.is_dir() and path.name.startswith('.ratools') for path in self.root.iterdir()))


if __name__ == '__main__':
    unittest.main()
