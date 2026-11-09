"""Content-free rule facts. Selection, execution, mutation and verification differ."""

from copy import deepcopy
from pathlib import Path

import fitz

from ratools_pdf.config.rules_catalog import equivalent_rule_ids


STEP_RULES = (
    ('title_from_filename', 'fast_web_view'),
    ('open_page_first', 'zoom_default'),
    ('page_layout_default',),
    ('initial_view_bookmarks_and_page',),
    ('collapse_all_bookmarks',),
    ('page_size_a4', 'page_size_letter'),
    ('bookmark_inherit_zoom', 'bookmark_open_new_window', 'bookmark_remove_external_links',
     'bookmark_remove_invalid', 'bookmark_remove_unknown_actions'),
    ('link_abs_to_rel_path', 'link_inherit_zoom', 'link_open_new_window', 'link_text_blue',
     'link_black_border', 'link_bordered_to_blue_border', 'link_unbordered_blue_to_blue_border',
     'link_remove_border'),
    ('cleanup_remove_external_uri', 'cleanup_remove_external_uri_and_text_black',
     'cleanup_remove_invalid_links', 'cleanup_remove_invalid_links_and_text_black',
     'cleanup_remove_unknown_action_links', 'cleanup_remove_dynamic_content',
     'cleanup_remove_attachments', 'cleanup_remove_tags', 'cleanup_remove_annotations',
     'cleanup_remove_metadata', 'cleanup_remove_all_links_bookmarks'),
    ('compress_images',),
)
FINAL_RULES = ('convert_pdf_version', 'fast_web_view', 'remove_pdf_restrictions',
               'flatten_rotation', 'flatten_annotations', 'generate_object_streams',
               'realloc_flate', 'compress_standard', 'compress_aggressive')


def plan_rule_results(selected, resolution, detectable):
    """Use the resolver's exact effective options, never reimplement smart selection."""
    effective = set(resolution.get('options', ()))
    suggested = set(resolution.get('suggested', ()))
    results = []
    for rule_id in sorted(set(selected)):
        supported = bool(equivalent_rule_ids(rule_id) & set(detectable))
        available = resolution.get('precheck_available')
        matched = (bool(equivalent_rule_ids(rule_id) & suggested)
                   if supported and available is True else None)
        planned = rule_id in effective
        skipped = None
        warning = []
        if rule_id == 'embed_nonstandard_fonts':
            planned, skipped = False, 'external_workflow_required'
        elif rule_id == 'filename_ectd_format':
            planned, skipped = False, 'output_naming_handled_by_worker'
        elif not planned:
            skipped = 'precheck_miss'
        if not supported:
            warning.append('precheck_unsupported')
        elif available is False:
            warning.append('precheck_unavailable')
        results.append({'rule_id': rule_id, 'matched': matched, 'planned': planned,
                        'executed': False, 'changed': False if not planned else None,
                        'verified': None, 'verification': 'not_verified',
                        'skipped_reason': skipped, 'warning': warning,
                        'status': 'pending' if planned else 'skipped'})
    return results


class RuleResultCollector:
    def __init__(self, selected, resolution, detectable):
        self.results = plan_rule_results(selected, resolution, detectable)
        self.by_id = {item['rule_id']: item for item in self.results}
        self.events = []
        self.active = []
        for item in self.results:
            self.emit(item, 'inspect')
            self.emit(item, 'plan')

    def emit(self, item, phase):
        self.events.append({'phase': phase, **deepcopy(item)})

    def begin(self, rule_ids):
        self.active = [self.by_id[key] for key in rule_ids
                       if key in self.by_id and self.by_id[key]['planned']]
        if ('page_size_a4' in rule_ids and
                self.by_id.get('page_size_a4', {}).get('planned') and
                self.by_id.get('page_size_letter', {}).get('planned')):
            letter = self.by_id['page_size_letter']
            letter.update(status='skipped', changed=False, skipped_reason='superseded_by_page_size_a4')
            self.active = [item for item in self.active if item['rule_id'] != 'page_size_letter']
        for item in self.active:
            item.update(executed=True, status='running')

    def observe(self, rule_id, changed):
        item = self.by_id.get(rule_id)
        if item is not None and item['planned']:
            item['changed'] = bool(item['changed'] or changed)

    def end(self):
        for item in self.active:
            item['status'] = 'executed'
            self.emit(item, 'execute')
        self.active = []

    def fail(self, error):
        for item in self.active:
            item['status'] = 'failed'
            item['warning'].append('execution_error:' + type(error).__name__)
            self.emit(item, 'error')
        self.active = []
        for item in self.results:
            if item['status'] == 'pending':
                item.update(status='skipped', skipped_reason='pipeline_aborted')
                item['warning'].append('pipeline_error:' + type(error).__name__)
                self.emit(item, 'error')

    def verify(self, output_path, input_path):
        """Independent output checks only for explicitly supported invariants."""
        try:
            with fitz.open(output_path) as doc:
                for item in self.results:
                    key = item['rule_id']
                    if not item['executed'] or item['status'] == 'failed':
                        continue
                    verified = None
                    if key == 'title_from_filename':
                        verified = (doc.metadata or {}).get('title') == Path(input_path).stem
                    elif key in {'page_size_a4', 'page_size_letter'}:
                        dimensions = ((210 / 25.4 * 72, 297 / 25.4 * 72)
                                      if key == 'page_size_a4' else (612, 792))
                        verified = all(abs(sorted((p.rect.width, p.rect.height))[0] - dimensions[0]) < .5
                                       and abs(sorted((p.rect.width, p.rect.height))[1] - dimensions[1]) < .5
                                       for p in doc)
                    elif key == 'cleanup_remove_external_uri':
                        verified = all(link.get('kind') != fitz.LINK_URI
                                       for page in doc for link in page.get_links())
                    elif key == 'cleanup_remove_annotations':
                        verified = all(not list(page.annots() or []) for page in doc)
                    elif key == 'cleanup_remove_attachments':
                        verified = doc.embfile_count() == 0
                    item['verified'] = verified
                    item['verification'] = ('passed' if verified else 'failed') if verified is not None else 'not_verified'
                    self.emit(item, 'verify')
        except Exception as error:
            for item in self.results:
                if item['executed']:
                    item['warning'].append('postcheck_error:' + type(error).__name__)

    def finish(self):
        for item in self.results:
            if item['executed'] and item['changed'] is None:
                item['warning'].append('mutation_not_observed')
            self.emit(item, 'result')
        return deepcopy(self.results), deepcopy(self.events)


def actual_changed(results):
    """Aggregate observations conservatively across planned rules."""
    if any(item.get('changed') is True for item in results):
        return True
    if results and all(item.get('changed') is False for item in results):
        return False
    return 'unknown'


def merge_execution(plan_results, results, events):
    """A frozen smart plan retains skipped rules when the engine executes force mode."""
    if not plan_results:
        return results, events
    observed = {item['rule_id']: item for item in results}
    merged = []
    for planned in plan_results:
        item = deepcopy(planned)
        actual = observed.get(item['rule_id'])
        if actual is not None:
            item.update(actual)
            for key in ('matched', 'planned'):
                item[key] = planned[key]
            if not planned['planned']:
                item['skipped_reason'] = planned['skipped_reason']
            item['warning'] = list(dict.fromkeys(planned['warning'] + actual['warning']))
        elif item['planned']:
            item.update(status='skipped', skipped_reason='execution_result_unavailable')
        merged.append(item)
    by_id = {item['rule_id']: item for item in plan_results}
    merged_events = []
    for event in events:
        entry = deepcopy(event)
        planned = by_id.get(entry['rule_id'])
        if planned:
            entry.update(matched=planned['matched'], planned=planned['planned'])
        merged_events.append(entry)
    for item in merged:
        if not item['planned']:
            merged_events.extend({'phase': phase, **deepcopy(item)}
                                 for phase in ('inspect', 'plan', 'result'))
    return merged, merged_events
