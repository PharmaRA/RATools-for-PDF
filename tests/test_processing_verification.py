"""PD02: plans bind real inputs; only verified staged PDFs may be published."""

import copy
import hashlib
import json
import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

import fitz

from ratools_pdf.controllers.workers import ProcessWorker

try:
    from ratools_pdf.services import pdf_facts, processing_plan
except ImportError:
    pdf_facts = processing_plan = None


def write_pdf(path, *, width=300, height=400, rotation=0, features=False):
    with fitz.open() as doc:
        page = doc.new_page(width=width, height=height)
        page.insert_text((30, 40), 'Synthetic test document')
        page.set_rotation(rotation)
        if features:
            page.insert_link({'kind': fitz.LINK_URI, 'from': fitz.Rect(20, 20, 80, 40),
                              'uri': 'https://example.invalid'})
            page.add_text_annot((100, 100), 'private annotation text')
            widget = fitz.Widget()
            widget.field_name = 'private field name'
            widget.field_type = fitz.PDF_WIDGET_TYPE_TEXT
            widget.rect = fitz.Rect(30, 120, 100, 140)
            page.add_widget(widget)
            doc.set_toc([[1, 'private bookmark', 1]])
            doc.embfile_add('private attachment.txt', b'private attachment body')
            doc.set_metadata({'author': 'private author'})
        doc.save(path)


class ProcessingVerificationTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.root = Path(self.directory.name)
        self.source = self.root/'source.pdf'
        write_pdf(self.source, features=True)

    def require_services(self):
        self.assertIsNotNone(pdf_facts, 'PD02 fact inspection service is missing')
        self.assertIsNotNone(processing_plan, 'PD02 input-bound planning service is missing')

    def plan(self, options=None, mode='smart'):
        self.require_services()
        return processing_plan.build_processing_plan(
            [str(self.source)], options or {'compress_images'}, processing_mode=mode,
            compression_settings={'dpi': 150, 'quality': 80}, common_base=str(self.root))

    def test_snapshot_records_numeric_structure_without_document_content(self):
        self.require_services()
        fact = pdf_facts.capture_pdf_snapshot(self.source)
        self.assertEqual(fact['sha256'], hashlib.sha256(self.source.read_bytes()).hexdigest())
        self.assertEqual(fact['bytes'], self.source.stat().st_size)
        self.assertEqual(fact['page_count'], 1)
        self.assertEqual(fact['pages'], [{'width': 300.0, 'height': 400.0, 'rotation': 0}])
        for metric in ('bookmarks', 'links', 'annotations', 'attachments', 'forms'):
            self.assertEqual(fact[metric], 1, metric)
        self.assertEqual(fact['signature_fields'], 0)
        self.assertIs(fact['signature_present'], False)
        self.assertEqual(fact['reopen'], 'passed')
        self.assertRegex(fact['pdf_version'], r'^\d+\.\d+$')
        self.assertNotIn('private', json.dumps(fact))

    def test_unreadable_snapshot_keeps_unknown_counts_and_safe_error_codes(self):
        self.require_services()
        self.source.write_bytes(b'not a PDF containing private body')
        fact = pdf_facts.capture_pdf_snapshot(self.source)
        self.assertIsNone(fact['page_count'])
        self.assertIsNone(fact['annotations'])
        self.assertEqual(fact['reopen'], 'failed')
        self.assertTrue(fact['errors'])
        self.assertNotIn(str(self.root), json.dumps(fact))
        self.assertNotIn('private', json.dumps(fact))

    def test_plan_distinguishes_hits_skips_and_user_selected_unsupported_rules(self):
        plan = self.plan({'cleanup_remove_annotations', 'flatten_rotation', 'compress_images'})
        item = plan['files'][0]
        rules = {rule['option_id']: rule for rule in item['rules']}
        self.assertEqual(rules['cleanup_remove_annotations']['precheck'], 'hit')
        self.assertEqual(rules['cleanup_remove_annotations']['decision'], 'execute')
        self.assertEqual(rules['flatten_rotation']['decision'], 'skip')
        self.assertEqual(rules['compress_images']['precheck'], 'unsupported')
        self.assertEqual(rules['compress_images']['decision'], 'execute')
        self.assertEqual(rules['compress_images']['reason'], 'selected_by_user')
        self.assertIsNone(rules['compress_images']['actual_changed'])
        self.assertIn('forms_present', item['risks'])
        self.assertIn('attachments_present', item['risks'])
        self.assertEqual(item['relative_path'], 'source.pdf')
        self.assertNotIn('private', json.dumps(plan))

    def test_force_plan_still_records_precheck_and_executes_unmatched_rules(self):
        plan = self.plan({'flatten_rotation'}, mode='force')
        rule = plan['files'][0]['rules'][0]
        self.assertEqual(rule['precheck'], 'miss')
        self.assertEqual(rule['decision'], 'execute')
        self.assertEqual(rule['reason'], 'force_mode')

    def test_plan_fingerprint_rejects_changed_bytes_even_when_file_still_opens(self):
        plan = self.plan()
        with self.source.open('ab') as stream:
            stream.write(b'\n% document changed after preview\n')
        self.assertFalse(processing_plan.input_matches_plan(self.source, plan['files'][0]))

    def test_plan_configuration_binding_rejects_changed_settings_and_options(self):
        plan = self.plan()
        self.assertTrue(processing_plan.configuration_matches_plan(
            plan, {'compress_images'}, 'smart', {'dpi': 150, 'quality': 80}))
        self.assertFalse(processing_plan.configuration_matches_plan(
            plan, {'compress_images'}, 'smart', {'dpi': 300, 'quality': 80}))
        self.assertFalse(processing_plan.configuration_matches_plan(
            plan, {'cleanup_remove_annotations'}, 'smart', {'dpi': 150, 'quality': 80}))

    def test_failed_file_does_not_prevent_planning_other_files_or_leak_parser_error(self):
        self.require_services()
        missing = self.root/'missing.pdf'
        plan = processing_plan.build_processing_plan(
            [str(missing), str(self.source)], {'compress_images'}, common_base=str(self.root))
        self.assertEqual([item['status'] for item in plan['files']], ['failed', 'ready'])
        self.assertTrue(plan['files'][0]['errors'])
        self.assertNotIn('private', json.dumps(plan))

    def test_expected_annotation_removal_is_structural_change_not_damage(self):
        self.require_services()
        before = pdf_facts.capture_pdf_snapshot(self.source)
        output = self.root/'output.pdf'
        with fitz.open(self.source) as doc:
            for page in doc:
                for annot in list(page.annots() or []):
                    page.delete_annot(annot)
            doc.save(output)
        after = pdf_facts.capture_pdf_snapshot(output)
        result = pdf_facts.verify_output(before, after, {'cleanup_remove_annotations'})
        self.assertTrue(result['publish_allowed'])
        self.assertEqual(result['save_reopen'], 'passed')
        self.assertEqual(result['structural_expectations'], 'passed')
        self.assertEqual(result['changes']['annotations'], 'expected_change')
        self.assertEqual(result['rule_verification'], 'not_verified')

    def test_unexpected_geometry_change_prevents_publication(self):
        self.require_services()
        before = pdf_facts.capture_pdf_snapshot(self.source)
        output = self.root/'output.pdf'
        write_pdf(output, width=600, height=800, features=True)
        result = pdf_facts.verify_output(before, pdf_facts.capture_pdf_snapshot(output), set())
        self.assertFalse(result['publish_allowed'])
        self.assertEqual(result['changes']['pages'], 'unexpected_change')
        allowed = pdf_facts.verify_output(before, pdf_facts.capture_pdf_snapshot(output), {'page_size_a4'})
        self.assertTrue(allowed['publish_allowed'])

    def test_unknown_check_never_claims_verification_passed(self):
        self.require_services()
        before = pdf_facts.capture_pdf_snapshot(self.source)
        after = copy.deepcopy(before)
        after['links'] = None
        after['errors'] = ['links:RuntimeError']
        result = pdf_facts.verify_output(before, after, set())
        self.assertFalse(result['publish_allowed'])
        self.assertEqual(result['structural_expectations'], 'not_verified')
        self.assertEqual(result['changes']['links'], 'unknown')

    def test_bad_success_output_preserves_prior_published_pdf(self):
        output_dir = self.root/'out'
        output_dir.mkdir()
        published = output_dir/'source.pdf'
        published.write_bytes(b'previous result must survive')
        worker = ProcessWorker([str(self.source)], set(), str(output_dir))
        events = []
        worker.structured_progress.connect(events.append)
        with patch('ratools_pdf.controllers.workers.mp.Process'), \
                patch('ratools_pdf.controllers.workers.mp.Pipe', return_value=(MagicMock(), MagicMock())):
            task = worker._start_process_task(0, str(self.source), False)
        Path(task['work_path']).write_bytes(b'corrupt staged PDF')
        self.assertFalse(worker._finish_process_task(task, True, 'engine reported success'))
        self.assertEqual(published.read_bytes(), b'previous result must survive')
        self.assertEqual(events[-1]['verification']['save_reopen'], 'failed')
        self.assertFalse(events[-1]['published'])

    def test_stale_preview_is_refused_before_creating_staging_directory(self):
        plan = self.plan()
        self.source.write_bytes(b'changed after preview')
        output = self.root/'out'
        worker = ProcessWorker([str(self.source)], {'compress_images'}, str(output),
                               compression_settings={'dpi': 150, 'quality': 80}, processing_plan=plan)
        events = []
        worker.structured_progress.connect(events.append)
        worker.run()
        self.assertFalse(output.exists())
        self.assertEqual(events[-1]['outcome'], 'stale_plan')
        self.assertFalse(events[-1]['published'])

    def test_confirmed_plan_processes_real_pdf_and_retains_actual_structured_evidence(self):
        self.require_services()
        plan = processing_plan.build_processing_plan(
            [str(self.source)], {'cleanup_remove_annotations'}, common_base=str(self.root))
        worker = ProcessWorker([str(self.source)], {'cleanup_remove_annotations'}, str(self.root/'out'),
                               processing_plan=plan)
        events = []
        worker.structured_progress.connect(events.append)
        worker.run()
        terminal = [event for event in events if event['type'] == 'terminal'][0]
        self.assertEqual(terminal['batch_id'], plan['batch_id'])
        self.assertTrue(terminal['engine_success'])
        self.assertTrue(terminal['published'])
        self.assertEqual(terminal['outcome'], 'succeeded')
        self.assertEqual(terminal['before']['annotations'], 1)
        self.assertEqual(terminal['after']['annotations'], 0)
        self.assertEqual(terminal['verification']['changes']['annotations'], 'expected_change')
        self.assertEqual(terminal['rule_actual_changed'], 'unknown')
        with fitz.open(self.root/'out'/'source.pdf') as doc:
            self.assertEqual(sum(1 for _ in doc[0].annots() or []), 0)
        with fitz.open(self.source) as doc:
            self.assertEqual(sum(1 for _ in doc[0].annots() or []), 1)

    def test_changed_input_during_processing_preserves_prior_output(self):
        self.require_services()
        plan = processing_plan.build_processing_plan(
            [str(self.source)], {'compress_images'}, common_base=str(self.root))
        output_dir = self.root/'out'
        output_dir.mkdir()
        published = output_dir/'source.pdf'
        published.write_bytes(b'previous successful result')
        worker = ProcessWorker([str(self.source)], {'compress_images'}, str(output_dir), processing_plan=plan)
        events = []
        worker.structured_progress.connect(events.append)
        with patch('ratools_pdf.controllers.workers.mp.Process'), \
                patch('ratools_pdf.controllers.workers.mp.Pipe', return_value=(MagicMock(), MagicMock())):
            task = worker._start_process_task(0, str(self.source), False)
        Path(task['work_path']).write_bytes(self.source.read_bytes())
        with self.source.open('ab') as stream:
            stream.write(b'\n% changed while processing\n')
        self.assertFalse(worker._finish_process_task(task, True, 'success'))
        self.assertEqual(events[-1]['outcome'], 'stale_plan')
        self.assertEqual(published.read_bytes(), b'previous successful result')

    def test_failed_plan_file_does_not_block_valid_file_in_same_batch(self):
        self.require_services()
        missing = self.root/'missing.pdf'
        files = [str(missing), str(self.source)]
        plan = processing_plan.build_processing_plan(files, {'cleanup_remove_annotations'}, common_base=str(self.root))
        worker = ProcessWorker(files, {'cleanup_remove_annotations'}, str(self.root/'out'), processing_plan=plan)
        events = []
        worker.structured_progress.connect(events.append)
        worker.run()
        terminal = [event for event in events if event['type'] == 'terminal']
        self.assertEqual([event['outcome'] for event in terminal], ['planning_failed', 'succeeded'])
        self.assertEqual({event['batch_id'] for event in terminal}, {plan['batch_id']})
        self.assertTrue((self.root/'out'/'source.pdf').exists())

    def test_pdf_version_upgrade_by_normal_save_is_not_document_damage(self):
        self.require_services()
        original = self.source.read_bytes()
        self.source.write_bytes(original.replace(original.splitlines()[0], b'%PDF-1.4', 1))
        plan = processing_plan.build_processing_plan(
            [str(self.source)], {'title_from_filename'}, common_base=str(self.root))
        worker = ProcessWorker([str(self.source)], {'title_from_filename'}, str(self.root/'out'), processing_plan=plan)
        events = []
        worker.structured_progress.connect(events.append)
        worker.run()
        terminal = [event for event in events if event['type'] == 'terminal'][0]
        self.assertEqual(terminal['before']['pdf_version'], '1.4')
        self.assertTrue(terminal['published'])
        self.assertEqual(terminal['verification']['changes']['pdf_version'], 'expected_change')

    def check_request_during_verification(self, request):
        self.require_services()
        output = self.root/'out'
        output.mkdir()
        published = output/'source.pdf'
        published.write_bytes(b'prior result')
        second_source = self.root/'second.pdf'
        write_pdf(second_source)
        worker = ProcessWorker([str(self.source), str(second_source)], set(), str(output))
        events = []
        worker.structured_progress.connect(events.append)
        with patch('ratools_pdf.controllers.workers.mp.Process'), \
                patch('ratools_pdf.controllers.workers.mp.Pipe', return_value=(MagicMock(), MagicMock())):
            task = worker._start_process_task(0, str(self.source), False)
        Path(task['work_path']).write_bytes(self.source.read_bytes())
        real_verify = pdf_facts.verify_output
        def verify_then_request(*args):
            result = real_verify(*args)
            getattr(worker, request)()
            return result
        with patch('ratools_pdf.controllers.workers.verify_output', side_effect=verify_then_request):
            self.assertFalse(worker._finish_process_task(task, True, 'engine success'))
        self.assertEqual(published.read_bytes(), b'prior result')
        self.assertFalse(Path(task['staging_dir']).exists())
        self.assertFalse(events[-1]['published'])
        return events[-1]

    def test_stop_requested_during_verification_prevents_publication(self):
        event = self.check_request_during_verification('request_stop')
        self.assertEqual(event['outcome'], 'stopped')
        self.assertEqual(event['status'], '已停止')

    def test_current_skip_requested_during_verification_prevents_publication(self):
        event = self.check_request_during_verification('request_skip_current')
        self.assertEqual(event['outcome'], 'skipped')
        self.assertEqual(event['status'], '已跳过')

    def test_dynamic_annotation_removal_is_an_expected_structural_change(self):
        self.require_services()
        dynamic_source = self.root/'dynamic.pdf'
        with fitz.open() as doc:
            page = doc.new_page(width=300, height=400)
            page.insert_text((30, 40), 'dynamic test')
            appearance = doc.get_new_xref()
            doc.update_object(appearance,
                              '<< /Type /XObject /Subtype /Form /BBox [0 0 10 10] /Resources << >> >>')
            doc.update_stream(appearance, b'q Q')
            doc.xref_set_key(page.xref, 'Annots',
                             f'[<< /Subtype /RichMedia /Rect [0 0 10 10] '
                             f'/AP << /N {appearance} 0 R >> >>]')
            doc.save(dynamic_source)
        plan = processing_plan.build_processing_plan(
            [str(dynamic_source)], {'cleanup_remove_dynamic_content'}, common_base=str(self.root))
        self.assertEqual(plan['files'][0]['status'], 'ready')
        worker = ProcessWorker([str(dynamic_source)], {'cleanup_remove_dynamic_content'},
                               str(self.root/'dynamic-out'), processing_plan=plan)
        events = []
        worker.structured_progress.connect(events.append)
        worker.run()
        terminal = [event for event in events if event['type'] == 'terminal'][0]
        self.assertTrue(terminal['published'])
        self.assertEqual(terminal['verification']['changes']['annotations'], 'expected_change')
        self.assertEqual(terminal['after']['annotations'], 0)

    def test_dynamic_annotation_counter_handles_standard_spacing_and_unenumerated_objects(self):
        self.require_services()
        with fitz.open() as doc:
            page = doc.new_page()
            appearance = doc.get_new_xref()
            doc.update_object(appearance,
                              '<< /Type /XObject /Subtype /Form /BBox [0 0 10 10] /Resources << >> >>')
            doc.update_stream(appearance, b'q Q')
            dynamic_xref = doc.get_new_xref()
            doc.update_object(dynamic_xref,
                              f'<< /Subtype /RichMedia /Rect [0 0 10 10] /AP << /N {appearance} 0 R >> >>')
            doc.xref_set_key(page.xref, 'Annots', f'[{dynamic_xref} 0 R]')
            self.assertEqual(pdf_facts._count_dynamic_page_annotations(doc, page), 1)

    def test_dynamic_annotation_counter_deduplicates_objects_already_seen_by_page_annots(self):
        self.require_services()
        with fitz.open() as doc:
            page = doc.new_page()
            dynamic_xref = doc.get_new_xref()
            doc.update_object(dynamic_xref, '<< /Subtype /RichMedia /Rect [0 0 10 10] >>')
            doc.xref_set_key(page.xref, 'Annots', f'[{dynamic_xref} 0 R]')
            self.assertEqual(pdf_facts._count_dynamic_page_annotations(doc, page, {dynamic_xref}), 0)

    def test_failed_verification_consumes_current_skip_without_skipping_next_file(self):
        self.require_services()
        output = self.root/'out'
        second_source = self.root/'second.pdf'
        write_pdf(second_source)
        worker = ProcessWorker([str(self.source)], set(), str(output))
        events = []
        worker.structured_progress.connect(events.append)
        def task(path):
            with patch('ratools_pdf.controllers.workers.mp.Process'), \
                    patch('ratools_pdf.controllers.workers.mp.Pipe', return_value=(MagicMock(), MagicMock())):
                result = worker._start_process_task(0, path, False)
            Path(result['work_path']).write_bytes(Path(path).read_bytes())
            return result
        first = task(str(self.source))
        actual_verify = pdf_facts.verify_output
        def failed_verify(*args):
            result = actual_verify(*args)
            worker.request_skip_current()
            result['publish_allowed'] = False
            result['structural_expectations'] = 'not_verified'
            return result
        with patch('ratools_pdf.controllers.workers.verify_output', side_effect=failed_verify):
            self.assertFalse(worker._finish_process_task(first, True, 'done'))
        other = self.root/'other.pdf'
        worker.request_skip_file(str(other))
        self.assertTrue(worker._finish_process_task(task(str(second_source)), True, 'done'))
        self.assertEqual(events[-1]['outcome'], 'succeeded')
        self.assertIn(os.path.normcase(os.path.abspath(other)), worker._skip_requested_files)


if __name__ == '__main__':
    unittest.main()
