"""PD03 evidence boundaries: real worker facts, portable export and hostile bundles."""

import copy
import csv
import hashlib
import importlib
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

from PySide6.QtWidgets import QApplication, QWidget

from ratools_pdf.controllers.log_controller import LogController
from ratools_pdf.controllers.main_controller import MainController
from ratools_pdf.controllers.workers import ProcessWorker
from ratools_pdf.services.processing_plan import build_processing_plan
from test_processing_verification import write_pdf


class EvidenceTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.root = Path(self.directory.name)
        self.source = self.root / 'source.pdf'
        write_pdf(self.source, features=True)
        self.host = SimpleNamespace(process_log_starts={}, process_log_rows=[])
        self.plan = build_processing_plan([str(self.source)], {'cleanup_remove_annotations'},
                                          common_base=str(self.root))
        MainController._retain_processing_plan(self.host, self.plan, 'running')
        self.host.processing_evidence['started_at'] = self.plan['created_at']
        MainController._finish_processing_evidence(self.host, 'cancelled')

    def service(self):
        try:
            return importlib.import_module('ratools_pdf.services.audit_evidence')
        except ImportError:
            self.fail('PD03 audit export service is missing')

    def export(self, name='bundle', evidence=None):
        target = self.root / name
        self.service().export_evidence_package(evidence or self.host.processing_evidence, target)
        return target

    def read(self, target, name='evidence.json'):
        return json.loads((target / name).read_text(encoding='utf-8'))

    def resign_test_bundle(self, target):
        """Model an attacker changing schemas/paths and recomputing unkeyed hashes."""
        manifest = self.read(target, 'integrity.json')
        for item in manifest['files']:
            if item['path'] in {'evidence.json', 'summary.csv'}:
                content = (target / item['path']).read_bytes()
                item.update(sha256=hashlib.sha256(content).hexdigest(), bytes=len(content))
        (target / 'integrity.json').write_text(json.dumps(manifest), encoding='utf-8')
        marker = self.read(target, 'COMPLETE.json')
        marker['integrity_sha256'] = hashlib.sha256((target / 'integrity.json').read_bytes()).hexdigest()
        (target / 'COMPLETE.json').write_text(json.dumps(marker), encoding='utf-8')

    def test_real_mixed_batch_preserves_same_name_identities_actual_results_and_versions(self):
        service = self.service()
        second = self.root / 'other' / 'source.pdf'
        second.parent.mkdir()
        write_pdf(second, width=420, features=True)
        missing = self.root / 'missing.pdf'
        files = [str(self.source), str(second), str(missing)]
        plan = build_processing_plan(files, {'cleanup_remove_annotations'}, common_base=str(self.root))
        MainController._retain_processing_plan(self.host, plan, 'running')
        self.host.processing_evidence['started_at'] = plan['created_at']
        worker = ProcessWorker(files, {'cleanup_remove_annotations'}, str(self.root / 'out'),
                               common_base=str(self.root), processing_plan=plan)
        worker.structured_progress.connect(lambda event: MainController.record_process_log_event(self.host, event))
        worker.run()
        MainController._finish_processing_evidence(self.host, 'finished')
        target = self.export()
        data = self.read(target)
        rows = {item['input']['relative_path']: item for item in data['files']}
        self.assertEqual(set(rows), {'source.pdf', 'other/source.pdf', 'missing.pdf'})
        self.assertEqual(len({item['file_id'] for item in rows.values()}), 3)
        for relative in ('source.pdf', 'other/source.pdf'):
            item = rows[relative]
            self.assertEqual(item['outcome'], 'succeeded')
            self.assertTrue(item['published'])
            self.assertEqual(item['output']['relative_path'], relative)
            self.assertEqual(item['output']['sha256'], hashlib.sha256((self.root / 'out' / relative).read_bytes()).hexdigest())
            self.assertEqual(item['before']['annotations'], 1)
            self.assertEqual(item['after']['annotations'], 0)
            self.assertEqual(item['verification']['save_reopen'], 'passed')
            self.assertEqual(item['verification']['rule_verification'], 'not_verified')
            self.assertIs(item['rule_actual_changed'], True)
            self.assertTrue(item['rule_results'][0]['changed'])
            self.assertTrue(item['rule_results'][0]['verified'])
        self.assertEqual(rows['missing.pdf']['outcome'], 'planning_failed')
        self.assertIsNone(rows['missing.pdf']['output'])
        self.assertEqual(data['batch']['outcome'], 'finished_with_failures')
        self.assertEqual(data['schema_version'], 'pd03.audit.v1')
        self.assertEqual(data['source_schema_version'], 'pd02.batch.v1')
        self.assertEqual(data['plan']['schema_version'], 'pd02.plan.v1')
        self.assertEqual(data['configuration']['version'], 'pd02.engine.v1')
        self.assertTrue(data['runtime']['application']['version'])
        self.assertTrue(data['runtime']['pymupdf']['version'])
        self.assertTrue(data['runtime']['qpdf']['version'])
        self.assertEqual(data['runtime']['application']['source_build'], 'unknown')
        serialized = json.dumps(data, ensure_ascii=False)
        self.assertNotIn(str(self.root), serialized)
        self.assertNotIn('private', serialized)
        self.assertTrue(service.verify_evidence_package(target)['valid'])

    def test_cancelled_and_stopped_batches_remain_incomplete_processing_facts(self):
        for outcome in ('cancelled', 'stopped', 'failed'):
            with self.subTest(outcome=outcome):
                evidence = copy.deepcopy(self.host.processing_evidence)
                evidence['outcome'] = outcome
                data = self.read(self.export(outcome, evidence))
                self.assertEqual(data['batch']['outcome'], outcome)
                self.assertEqual(data['files'][0]['outcome'], 'not_started')
                self.assertIsNone(data['files'][0]['verification'])
                self.assertIn('output:not_published', data['files'][0]['missing_evidence'])

    def test_terminal_failures_export_stable_reason_codes_and_diagnostic_gap(self):
        self.service()
        for outcome, code in (('processing_failed', 'engine_failure'),
                              ('verification_failed', 'verification_failed'),
                              ('stale_plan', 'stale_plan'), ('stopped', 'stop_requested'),
                              ('planning_failed', 'planning_failed')):
            with self.subTest(outcome=outcome):
                evidence = copy.deepcopy(self.host.processing_evidence)
                evidence['files'][self.plan['files'][0]['file_id']]['outcome'] = outcome
                data = self.read(self.export(outcome, evidence))
                self.assertIn(code, data['files'][0].get('reason_codes', []))
                self.assertIn('engine_diagnostic:not_recorded', data['files'][0]['missing_evidence'])

    def test_actual_worker_stop_exports_stopped_without_success(self):
        self.service()
        from ratools_pdf.services.pdf_facts import verify_output
        MainController._retain_processing_plan(self.host, self.plan, 'running')
        worker = ProcessWorker([str(self.source)], {'cleanup_remove_annotations'}, str(self.root / 'out'),
                               processing_plan=self.plan)
        worker.structured_progress.connect(lambda event: MainController.record_process_log_event(self.host, event))
        def verify_then_stop(*args):
            result = verify_output(*args)
            worker.request_stop()
            return result
        with patch('ratools_pdf.controllers.workers.verify_output', side_effect=verify_then_stop):
            worker.run()
        MainController._finish_processing_evidence(self.host, 'finished')
        data = self.read(self.export())
        self.assertEqual(data['files'][0]['outcome'], 'stopped')
        self.assertFalse(data['files'][0]['published'])
        self.assertIsNone(data['files'][0]['output'])

    def test_new_batch_records_execution_metadata_and_old_batch_stays_unknown(self):
        self.service()
        data = self.read(self.export('current'))
        self.assertEqual(data['processing_runtime']['application']['version'], '0.8.1')
        self.assertTrue(data['processing_runtime']['pymupdf']['version'])
        self.assertEqual(data['processing_runtime']['qpdf']['version'], 'unknown')
        old = copy.deepcopy(self.host.processing_evidence)
        old.pop('execution_runtime', None)
        data = self.read(self.export('old', old))
        self.assertEqual(data['processing_runtime']['application']['version'], 'unknown')

    def test_running_batch_refuses_before_creating_directory(self):
        service = self.service()
        self.host.processing_evidence.update(ended_at=None, outcome='running')
        with self.assertRaises(ValueError):
            self.export()
        self.assertFalse((self.root / 'bundle').exists())

    def test_existing_directory_is_never_overwritten(self):
        self.service()
        target = self.root / 'bundle'
        target.mkdir()
        (target / 'keep.txt').write_text('keep', encoding='utf-8')
        with self.assertRaises(FileExistsError):
            self.export()
        self.assertEqual((target / 'keep.txt').read_text(), 'keep')
        self.assertEqual({item.name for item in target.iterdir()}, {'keep.txt'})

    def test_partial_write_retains_directory_and_verifier_rejects_it(self):
        service = self.service()
        real_open = Path.open
        def fail_csv(path, *args, **kwargs):
            if path.name == 'summary.csv':
                raise OSError('private absolute path must never be exported')
            return real_open(path, *args, **kwargs)
        with patch.object(Path, 'open', fail_csv):
            with self.assertRaises(service.AuditExportError) as caught:
                self.export()
        target = self.root / 'bundle'
        self.assertTrue(target.is_dir())
        self.assertFalse((target / 'COMPLETE.json').exists())
        self.assertNotIn('private', str(caught.exception))
        self.assertFalse(service.verify_evidence_package(target)['valid'])

    def test_csv_escapes_whitespace_formulas_apostrophes_and_roundtrips_unicode_newlines(self):
        service = self.service()
        for index, value in enumerate(('=SUM(1,2)', ' \t+1', '\t-1', '\r@SUM(1)', "'原始\n多行", '普通\n原因')):
            with self.subTest(value=value):
                evidence = copy.deepcopy(self.host.processing_evidence)
                evidence['plan']['files'][0]['rules'][0]['reason'] = value
                target = self.export(f'csv-{index}', evidence)
                with (target / 'summary.csv').open(encoding='utf-8-sig', newline='') as stream:
                    row = next(csv.DictReader(stream))
                cell = row['reasons']
                if value.lstrip().startswith(('=', '+', '-', '@')) or value.startswith("'"):
                    self.assertTrue(cell.startswith("'"))
                self.assertEqual(service.decode_csv_cell(cell), value)
                self.assertEqual(self.read(target)['files'][0]['reasons'], [value])

    def test_unknown_fields_do_not_export_document_content_passwords_or_logs(self):
        self.service()
        evidence = self.host.processing_evidence
        evidence['password'] = 'private secret'
        evidence['plan']['configuration']['password'] = 'private secret'
        evidence['files'][self.plan['files'][0]['file_id']]['message'] = 'private body'
        evidence['plan']['files'][0]['before']['metadata'] = {'author': 'private author'}
        data = self.read(self.export())
        self.assertNotIn('private', json.dumps(data))
        self.assertNotIn('password', json.dumps(data))

    def test_absolute_and_parent_identities_are_rejected_without_output(self):
        service = self.service()
        for index, relative in enumerate(('C:/private/source.pdf', '/private/source.pdf', '../escape.pdf', 'C:source.pdf')):
            with self.subTest(relative=relative):
                evidence = copy.deepcopy(self.host.processing_evidence)
                evidence['plan']['files'][0]['relative_path'] = relative
                evidence['files'][self.plan['files'][0]['file_id']]['relative_path'] = relative
                with self.assertRaises(ValueError):
                    self.export(f'bad-{index}', evidence)
                self.assertFalse((self.root / f'bad-{index}').exists())

    def test_unsupported_source_schemas_fail_before_writing(self):
        self.service()
        for index, field in enumerate(('batch', 'plan', 'snapshot')):
            evidence = copy.deepcopy(self.host.processing_evidence)
            obj = {'batch': evidence, 'plan': evidence['plan'], 'snapshot': evidence['plan']['files'][0]['before']}[field]
            obj['schema_version'] = 'future.v999'
            with self.assertRaises(ValueError):
                self.export(f'future-{index}', evidence)
            self.assertFalse((self.root / f'future-{index}').exists())

    def test_verifier_detects_tampered_missing_and_extra_payloads(self):
        service = self.service()
        for case in ('tampered', 'missing', 'extra'):
            target = self.export(case)
            if case == 'tampered':
                with (target / 'summary.csv').open('a', encoding='utf-8') as stream:
                    stream.write('tamper')
            elif case == 'missing':
                (target / 'summary.csv').rename(self.root / 'held.csv')
            else:
                (target / 'extra.txt').write_text('extra', encoding='utf-8')
            self.assertFalse(service.verify_evidence_package(target)['valid'], case)

    def test_rehashed_unknown_export_schema_and_manifest_escape_are_rejected(self):
        service = self.service()
        target = self.export('future')
        data = self.read(target)
        data['schema_version'] = 'future.v999'
        (target / 'evidence.json').write_text(json.dumps(data), encoding='utf-8')
        self.resign_test_bundle(target)
        self.assertFalse(service.verify_evidence_package(target)['valid'])

    def test_rehashed_missing_structure_and_nested_future_schema_are_rejected(self):
        service = self.service()
        for case in ('missing_files', 'snapshot_schema', 'event_schema'):
            target = self.export(case)
            data = self.read(target)
            if case == 'missing_files':
                data['files'] = []
            elif case == 'snapshot_schema':
                data['plan']['files'][0]['before']['schema_version'] = 'future.v999'
            else:
                data['events'].append({'batch_id': data['batch']['batch_id'],
                                       'file_id': data['files'][0]['file_id'],
                                       'before': None, 'after': None,
                                       'verification': {'schema_version': 'future.v999'}})
            (target / 'evidence.json').write_text(json.dumps(data), encoding='utf-8')
            self.resign_test_bundle(target)
            self.assertFalse(service.verify_evidence_package(target)['valid'], case)
        target = self.export('escape')
        manifest = self.read(target, 'integrity.json')
        manifest['files'][0]['path'] = '../outside.json'
        (target / 'integrity.json').write_text(json.dumps(manifest), encoding='utf-8')
        self.resign_test_bundle(target)
        self.assertFalse(service.verify_evidence_package(target)['valid'])

    def test_callable_cli_has_success_and_failure_exit_codes(self):
        self.service()
        target = self.export()
        args = [sys.executable, '-B', '-m', 'ratools_pdf.services.audit_evidence', 'verify', str(target)]
        result = subprocess.run(args, capture_output=True, text=True, encoding='utf-8')
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertTrue(json.loads(result.stdout)['valid'])
        (target / 'extra.txt').write_text('extra', encoding='utf-8')
        result = subprocess.run(args, capture_output=True, text=True, encoding='utf-8')
        self.assertNotEqual(result.returncode, 0)
        self.assertFalse(json.loads(result.stdout)['valid'])


class AuditUITests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.root = Path(self.directory.name)
        source = self.root / 'source.pdf'
        write_pdf(source)
        self.host = SimpleNamespace(process_log_starts={}, process_log_rows=[], process_logs='',
                                    loaded_files=[str(source)], last_precheck_results=[], last_output_dir='')
        plan = build_processing_plan([str(source)], {'compress_images'}, common_base=str(self.root))
        MainController._retain_processing_plan(self.host, plan, 'running')
        MainController._finish_processing_evidence(self.host, 'cancelled')
        self.messages = []
        self.view = QWidget()
        self.view.settings_dialog = SimpleNamespace(default_output_edit=SimpleNamespace(text=lambda: ''))
        for name in ('warning', 'error', 'success'):
            setattr(self.view, f'show_{name}_message', lambda *args, kind=name: self.messages.append((kind, args)))
        self.controller = LogController(self.host, self.view)
        self.addCleanup(self.view.close)

    def require_export(self):
        self.assertTrue(hasattr(self.controller, 'export_audit_evidence'), 'PD03 actual GUI export action is missing')

    def test_actual_log_dialog_button_exports_latest_batch_without_legacy_logs(self):
        self.require_export()
        self.controller.show_log_dialog()
        self.addCleanup(self.controller.log_dialog.close)
        self.assertTrue(self.controller.log_dialog.btn_export_audit.isEnabled())
        with patch('ratools_pdf.controllers.log_controller.QFileDialog.getExistingDirectory', return_value=str(self.root)):
            self.controller.log_dialog.btn_export_audit.click()
        bundles = [path for path in self.root.iterdir() if path.is_dir()]
        self.assertEqual(len(bundles), 1)
        self.assertTrue((bundles[0] / 'COMPLETE.json').is_file())
        self.assertEqual(self.messages[-1][0], 'success')

    def test_dialog_cancellation_creates_no_export_directory(self):
        self.require_export()
        with patch('ratools_pdf.controllers.log_controller.QFileDialog.getExistingDirectory', return_value=''):
            self.controller.export_audit_evidence()
        self.assertFalse(any(path.is_dir() for path in self.root.iterdir()))

    def test_batch_switch_during_modal_refuses_export(self):
        self.require_export()
        def switch(*args):
            self.host.processing_evidence = dict(self.host.processing_evidence, batch_id='different')
            return str(self.root)
        with patch('ratools_pdf.controllers.log_controller.QFileDialog.getExistingDirectory', side_effect=switch):
            self.controller.export_audit_evidence()
        self.assertFalse(any(path.is_dir() for path in self.root.iterdir()))
        self.assertEqual(self.messages[-1][0], 'warning')

    def test_export_freezes_selected_evidence_before_modal(self):
        self.require_export()
        def mutate(*args):
            self.host.processing_evidence['outcome'] = 'failed'
            return str(self.root)
        with patch('ratools_pdf.controllers.log_controller.QFileDialog.getExistingDirectory', side_effect=mutate):
            self.controller.export_audit_evidence()
        bundles = [path for path in self.root.iterdir() if path.is_dir()]
        data = json.loads((bundles[0] / 'evidence.json').read_text(encoding='utf-8'))
        self.assertEqual(data['batch']['outcome'], 'cancelled')


if __name__ == '__main__':
    unittest.main()
