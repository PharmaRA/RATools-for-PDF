"""PD07 observable rule outcomes; fixtures contain synthetic, non-business PDF data."""

from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch
from types import SimpleNamespace

import fitz

from ratools_pdf.config import rules_catalog
from ratools_pdf.pdf import processor
from ratools_pdf.controllers.workers import ProcessWorker
from ratools_pdf.controllers.main_controller import MainController
from ratools_pdf.services import audit_evidence, processing_plan
from ratools_pdf.services.rule_results import merge_execution


class RuleResultsTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.root = Path(self.tmp.name)
        self.source = self.root / 'sample.pdf'
        self.output = self.root / 'result.pdf'

    def make_pdf(self, title='', uri=False, width=500, height=700):
        with fitz.open() as doc:
            page = doc.new_page(width=width, height=height)
            if title:
                doc.set_metadata({'title': title})
            if uri:
                page.insert_text((55, 65), 'synthetic link', color=(0, 0, 1))
                page.insert_link({'kind': fitz.LINK_URI, 'from': fitz.Rect(50, 50, 200, 72),
                                  'uri': 'https://example.invalid'})
            doc.save(self.source)

    def process(self, options, mode='force'):
        self.assertTrue(hasattr(processor, 'process_document_detailed'),
                        'PD07 must expose detailed outcomes without changing the public tuple')
        return processor.process_document_detailed(
            str(self.source), str(self.output), set(options), processing_mode=mode)

    def rule(self, result, rule_id):
        return next(item for item in result['rule_results'] if item['rule_id'] == rule_id)

    def test_title_records_real_change_and_independent_output_postcondition(self):
        # Losing the title mutation observation or independent reopen check must fail.
        self.make_pdf(title='old')
        result = self.process({'title_from_filename'})
        self.assertTrue(result['success'], result['message'])
        rule = self.rule(result, 'title_from_filename')
        self.assertEqual((rule['matched'], rule['planned'], rule['changed'], rule['verified']),
                         (None, True, True, True))
        with fitz.open(self.output) as doc:
            self.assertEqual(doc.metadata['title'], 'sample')
        self.assertTrue(any(event['rule_id'] == 'title_from_filename' and
                            event['phase'] == 'execute' and event['changed'] is True
                            for event in result['rule_events']))

    def test_title_already_equal_is_observed_unchanged(self):
        # Treating selected or called as changed must fail on the same-title PDF.
        self.make_pdf(title='sample')
        result = self.process({'title_from_filename'})
        self.assertEqual(self.rule(result, 'title_from_filename')['changed'], False)
        self.assertEqual(self.source.read_bytes(), self.output.read_bytes())

    def test_smart_miss_is_skipped_and_has_no_execute_event(self):
        self.make_pdf(title='sample')
        result = self.process({'title_from_filename'}, 'smart')
        rule = self.rule(result, 'title_from_filename')
        self.assertEqual((rule['matched'], rule['planned'], rule['executed'], rule['changed']),
                         (False, False, False, False))
        self.assertEqual(rule['skipped_reason'], 'precheck_miss')
        self.assertFalse(any(event['phase'] == 'execute' for event in result['rule_events']))

    def test_unsupported_precheck_remains_unknown_but_real_resize_is_observed(self):
        self.make_pdf()
        result = self.process({'page_size_a4'}, 'smart')
        rule = self.rule(result, 'page_size_a4')
        self.assertEqual((rule['matched'], rule['planned'], rule['changed'], rule['verified']),
                         (None, True, True, True))
        self.assertIn('precheck_unsupported', rule['warning'])
        with fitz.open(self.output) as doc:
            self.assertAlmostEqual(doc[0].rect.width, 210 / 25.4 * 72, delta=.5)

    def test_composite_rule_retains_its_id_and_does_not_claim_complete_postcheck(self):
        self.make_pdf(uri=True)
        result = self.process({'cleanup_remove_external_uri_and_text_black'}, 'smart')
        self.assertTrue(result['success'], result['message'])
        rule = self.rule(result, 'cleanup_remove_external_uri_and_text_black')
        self.assertEqual((rule['matched'], rule['planned'], rule['changed']), (True, True, True))
        self.assertIsNone(rule['verified'], 'URI removal alone cannot verify text recoloring')
        self.assertEqual([item['rule_id'] for item in result['rule_results']],
                         ['cleanup_remove_external_uri_and_text_black'])
        with fitz.open(self.output) as doc:
            self.assertEqual(doc[0].get_links(), [])
            traces = [item for item in doc[0].get_texttrace() if item.get('type') == 0]
            self.assertLess(traces[0]['color'][2], .1)

    def test_failure_preserves_observed_change_and_never_executes_pending_rule(self):
        self.make_pdf()
        with patch.object(processor, '_step_open_action', side_effect=RuntimeError('synthetic failure')):
            # Step identity is resolved when a run begins so the production call is replaced.
            with patch.object(processor, '_PIPELINE_STEPS',
                              (processor._step_document_title, processor._step_open_action,
                               processor._step_page_size)):
                result = self.process({'title_from_filename', 'open_page_first', 'page_size_a4'})
        self.assertFalse(result['success'])
        self.assertTrue(self.rule(result, 'title_from_filename')['changed'])
        self.assertIsNone(self.rule(result, 'title_from_filename')['verified'])
        self.assertEqual(self.rule(result, 'open_page_first')['status'], 'failed')
        pending = self.rule(result, 'page_size_a4')
        self.assertEqual((pending['executed'], pending['changed'], pending['skipped_reason']),
                         (False, None, 'pipeline_aborted'))
        self.assertTrue(any(event['phase'] == 'error' for event in result['rule_events']))

    def test_unobserved_rewrite_never_becomes_changed_or_verified_by_selection(self):
        self.make_pdf(title='sample')
        result = self.process({'generate_object_streams'})
        self.assertTrue(result['success'], result['message'])
        rule = self.rule(result, 'generate_object_streams')
        self.assertTrue(rule['executed'])
        self.assertIsNone(rule['changed'])
        self.assertIsNone(rule['verified'])
        self.assertIn('mutation_not_observed', rule['warning'])

    def test_catalog_covers_non_ui_engine_rule_and_aliases_share_precheck_identity(self):
        self.assertTrue(hasattr(rules_catalog, 'RULE_METADATA'), 'engine rules need stable metadata')
        self.assertIn('remove_pdf_restrictions', rules_catalog.RULE_METADATA)
        self.assertEqual(rules_catalog.equivalent_rule_ids('cleanup_remove_external_uri_and_text_black'),
                         {'cleanup_remove_external_uri', 'cleanup_remove_external_uri_and_text_black'})

    def test_worker_and_export_keep_actual_events_instead_of_preview_placeholders(self):
        # Dropping the detailed child result at pipe/host/export boundaries must fail.
        self.make_pdf(title='old')
        plan = processing_plan.build_processing_plan(
            [str(self.source)], {'title_from_filename'}, common_base=str(self.root))
        worker = ProcessWorker([str(self.source)], {'title_from_filename'}, str(self.root / 'out'),
                               processing_plan=plan)
        events = []
        worker.structured_progress.connect(events.append)
        worker.run()
        terminal = next(event for event in events if event['type'] == 'terminal')
        self.assertTrue(terminal['published'])
        self.assertTrue(terminal.get('rule_results'), 'worker discarded detailed child result')
        self.assertTrue(terminal['rule_results'][0]['changed'])
        self.assertTrue(terminal['rule_results'][0]['verified'])
        self.assertTrue(terminal['rule_events'])
        evidence = {'schema_version': 'pd02.batch.v1', 'batch_id': plan['batch_id'],
                    'plan': plan, 'ended_at': '2026-11-09T00:00:00Z', 'outcome': 'completed',
                    'files': {terminal['file_id']: {}}, 'events': []}
        host = SimpleNamespace(processing_evidence_history={plan['batch_id']: evidence},
                               process_log_starts={}, process_log_rows=[])
        for event in events:
            MainController.record_process_log_event(host, event)
        self.assertTrue(evidence['files'][terminal['file_id']]['rule_results'][0]['changed'])
        self.assertEqual(host.process_log_rows[0]['changes'], '同步文件名为标题')
        document = audit_evidence._build_document(evidence)
        self.assertTrue(document['files'][0]['rule_results'][0]['changed'])
        self.assertTrue(document['events'][-1]['rule_events'])
        self.assertTrue(document['files'][0]['rules'][0]['actual_changed'])

    def test_worker_preserves_smart_skipped_records_from_confirmed_plan(self):
        self.make_pdf(title='sample')
        plan = processing_plan.build_processing_plan(
            [str(self.source)], {'title_from_filename', 'page_size_a4'}, common_base=str(self.root))
        worker = ProcessWorker([str(self.source)], {'title_from_filename', 'page_size_a4'},
                               str(self.root / 'out'), processing_plan=plan)
        events = []
        worker.structured_progress.connect(events.append)
        worker.run()
        terminal = next(event for event in events if event['type'] == 'terminal')
        results = {item['rule_id']: item for item in terminal.get('rule_results', [])}
        self.assertIn('title_from_filename', results, 'frozen smart skips must survive execution')
        self.assertEqual((results['title_from_filename']['matched'],
                          results['title_from_filename']['executed']), (False, False))
        self.assertEqual(results['title_from_filename']['skipped_reason'], 'precheck_miss')

    def test_failure_before_open_emits_error_without_claiming_execution(self):
        self.source.write_bytes(b'synthetic invalid PDF')
        result = self.process({'title_from_filename'})
        self.assertFalse(result['success'])
        rule = self.rule(result, 'title_from_filename')
        self.assertEqual((rule['executed'], rule['changed']), (False, None))
        self.assertTrue(any(event['phase'] == 'error' for event in result['rule_events']))

    def test_a4_priority_does_not_claim_letter_was_executed(self):
        self.make_pdf()
        result = self.process({'page_size_a4', 'page_size_letter'})
        self.assertTrue(result['success'], result['message'])
        letter = self.rule(result, 'page_size_letter')
        self.assertFalse(letter['executed'])
        self.assertEqual(letter['skipped_reason'], 'superseded_by_page_size_a4')

    def test_uri_decolor_does_not_report_invalid_composite_changed(self):
        self.make_pdf(uri=True)
        result = self.process({'cleanup_remove_external_uri_and_text_black',
                               'cleanup_remove_invalid_links_and_text_black'})
        self.assertTrue(result['success'], result['message'])
        self.assertTrue(self.rule(result, 'cleanup_remove_external_uri_and_text_black')['changed'])
        self.assertIs(self.rule(result, 'cleanup_remove_invalid_links_and_text_black')['changed'], False)
        with fitz.open(self.output) as doc:
            self.assertEqual(doc[0].get_links(), [])
            traces = [item for item in doc[0].get_texttrace() if item.get('type') == 0]
            self.assertLess(traces[0]['color'][2], .1)

    def test_invalid_decolor_does_not_report_uri_composite_changed(self):
        self.make_pdf(uri=True)
        # PyMuPDF hides actionless annotations; emulate only its enumeration boundary.
        # Link deletion and text stream recoloring below still operate on a real PDF.
        get_links = fitz.Page.get_links

        def enumerate_invalid(page):
            return [{**link, 'kind': fitz.LINK_NONE} if link.get('kind') == fitz.LINK_URI else link
                    for link in get_links(page)]

        with patch.object(fitz.Page, 'get_links', enumerate_invalid):
            result = self.process({'cleanup_remove_external_uri_and_text_black',
                                   'cleanup_remove_invalid_links_and_text_black'})
        self.assertTrue(result['success'], result['message'])
        self.assertTrue(self.rule(result, 'cleanup_remove_invalid_links_and_text_black')['changed'])
        self.assertIs(self.rule(result, 'cleanup_remove_external_uri_and_text_black')['changed'], False)
        with fitz.open(self.output) as doc:
            self.assertEqual(doc[0].get_links(), [])
            traces = [item for item in doc[0].get_texttrace() if item.get('type') == 0]
            self.assertLess(traces[0]['color'][2], .1)

    def test_frozen_force_plan_keeps_runtime_paper_supersession_reason(self):
        self.make_pdf()
        selected = {'page_size_a4', 'page_size_letter'}
        plan = processing_plan.build_processing_plan(
            [str(self.source)], selected, 'force', common_base=str(self.root))
        worker = ProcessWorker([str(self.source)], selected, str(self.root / 'out'),
                               processing_mode='force', processing_plan=plan)
        events = []
        worker.structured_progress.connect(events.append)
        worker.run()
        terminal = next(event for event in events if event['type'] == 'terminal')
        letter = next(item for item in terminal['rule_results'] if item['rule_id'] == 'page_size_letter')
        self.assertFalse(letter['executed'])
        self.assertEqual(letter['skipped_reason'], 'superseded_by_page_size_a4')

    def test_frozen_plan_merge_keeps_actual_pipeline_abort_reason(self):
        self.make_pdf()
        selected = {'title_from_filename', 'open_page_first', 'page_size_a4'}
        plan = processing_plan.build_processing_plan(
            [str(self.source)], selected, 'force', common_base=str(self.root))
        with patch.object(processor, '_step_open_action', side_effect=RuntimeError('synthetic failure')):
            with patch.object(processor, '_PIPELINE_STEPS',
                              (processor._step_document_title, processor._step_open_action,
                               processor._step_page_size)):
                actual = self.process(selected)
        results, _ = merge_execution(plan['files'][0]['rule_results'],
                                     actual['rule_results'], actual['rule_events'])
        pending = next(item for item in results if item['rule_id'] == 'page_size_a4')
        self.assertFalse(pending['executed'])
        self.assertEqual(pending['skipped_reason'], 'pipeline_aborted')


if __name__ == '__main__':
    unittest.main()
