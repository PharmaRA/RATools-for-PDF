"""Input-bound, JSON-compatible processing plans. No text or metadata is retained."""

import hashlib
import json
import os
import uuid
from datetime import datetime, timezone
from pathlib import Path

from ratools_pdf.pdf import precheck
from ratools_pdf.services.pdf_facts import InspectionCancelled, capture_pdf_snapshot, fingerprint
from ratools_pdf.services import profiles
from ratools_pdf.services.rule_results import plan_rule_results
from ratools_pdf.config.rules_catalog import equivalent_rule_ids


CONFIG_VERSION = 'pd02.engine.v1'


def timestamp():
    return datetime.now(timezone.utc).isoformat()


def source_file_id(path):
    normalized = os.path.normcase(os.path.abspath(path))
    return hashlib.sha256(normalized.encode('utf-8')).hexdigest()[:24]


def configuration(options, processing_mode, compression_settings, profile=None):
    settings = compression_settings or {}
    result = {'version': CONFIG_VERSION, 'options': sorted(set(options or [])),
            'processing_mode': 'force' if processing_mode == 'force' else 'smart',
            'compression_settings': {key: settings[key] for key in ('dpi', 'quality') if key in settings}}
    if profile is not None:
        validated = profiles.validate_profile(profile)
        if set(validated['enabled_rules']) != set(result['options']) & set(profiles.rules_catalog.OPTION_TITLES):
            raise ValueError('Profile rules do not match processing options')
        expected = validated['rule_parameters'].get('compress_images', {})
        actual = result['compression_settings']
        if expected and expected != actual:
            raise ValueError('Profile parameters do not match processing settings')
        result['profile'] = validated
        result['profile_digest'] = profiles.profile_digest(validated)
    return result


def _configuration_digest(config):
    return hashlib.sha256(json.dumps(config, sort_keys=True).encode('utf-8')).hexdigest()


def configuration_matches_plan(plan, options, processing_mode, compression_settings, profile=None):
    # Workers consume the frozen snapshot; UI callers pass their current identity explicitly.
    if profile is None:
        profile = plan.get('configuration', {}).get('profile')
    try:
        current = configuration(options, processing_mode, compression_settings, profile)
        if profile is not None and plan.get('configuration', {}).get('profile_digest') != current['profile_digest']:
            return False
        return plan.get('configuration_digest') == _configuration_digest(current)
    except ValueError:
        return False


def input_matches_plan(path, file_plan, cancelled=None):
    try:
        actual = fingerprint(path, cancelled)
        return all(actual[key] == file_plan['before'].get(key) for key in ('sha256', 'bytes'))
    except (OSError, KeyError):
        return False


def build_processing_plan(files, options, processing_mode='smart', compression_settings=None,
                          common_base='', cancelled=None, progress=None, profile=None):
    config = configuration(options, processing_mode, compression_settings, profile)
    selected = set(config['options'])
    plan = {'schema_version': 'pd02.plan.v1', 'batch_id': uuid.uuid4().hex,
            'created_at': timestamp(), 'configuration': config,
            'configuration_digest': _configuration_digest(config), 'files': [], 'cancelled': False}
    for index, source in enumerate(files):
        if cancelled and cancelled():
            plan['cancelled'] = True
            break
        try:
            relative = Path(source).resolve().relative_to(Path(common_base).resolve()).as_posix()
        except (ValueError, OSError):
            relative = Path(source).name
        item = {'file_id': source_file_id(source), 'relative_path': relative,
                'status': 'failed', 'rules': [], 'effective_options': [], 'risks': [], 'errors': [], 'before': None}
        plan['files'].append(item)
        try:
            before = capture_pdf_snapshot(source, cancelled)
            item['before'] = before
            report = precheck.build_precheck_report(source, selected_options=selected)
            if before['reopen'] != 'passed' or before['errors'] or not report.get('available'):
                item['errors'] = list(before['errors']) or ['precheck:unavailable']
            else:
                item['status'] = 'ready'
            suggestions = {key for key, value in report.get('suggestions', {}).items()
                           if not value.get('report_only')}
            for option in sorted(selected):
                equivalent = equivalent_rule_ids(option)
                supported = bool(equivalent & precheck.PRECHECK_DETECTABLE_OPTIONS)
                hit = bool(equivalent & suggestions)
                detection = ('hit' if hit else 'miss') if supported else 'unsupported'
                if not report.get('available') and supported:
                    detection = 'unknown'
                if item['status'] != 'ready':
                    decision, reason = 'skip', 'inspection_failed'
                elif option == 'embed_nonstandard_fonts':
                    decision, reason = 'skip', 'external_workflow_required'
                elif config['processing_mode'] == 'force':
                    decision, reason = 'execute', 'force_mode'
                elif not supported:
                    decision, reason = 'execute', 'selected_by_user'
                elif hit:
                    decision, reason = 'execute', 'precheck_hit'
                else:
                    decision, reason = 'skip', 'precheck_miss'
                item['rules'].append({'option_id': option, 'precheck': detection, 'decision': decision,
                                      'reason': reason, 'actual_changed': None, 'verified': 'not_verified'})
                if decision == 'execute':
                    item['effective_options'].append(option)
            item['rule_results'] = plan_rule_results(selected, {
                'options': item['effective_options'], 'suggested': suggestions,
                'precheck_available': report.get('available') if config['processing_mode'] == 'smart' else None,
            }, precheck.PRECHECK_DETECTABLE_OPTIONS)
            for metric, risk in (('signature_present', 'signature_present'), ('forms', 'forms_present'),
                                 ('attachments', 'attachments_present')):
                if before.get(metric):
                    item['risks'].append(risk)
                elif before.get(metric) is None:
                    item['risks'].append(metric + '_unknown')
            if selected & {'page_size_a4', 'page_size_letter', 'flatten_rotation'}:
                item['risks'].append('geometry_may_change')
        except InspectionCancelled:
            plan['cancelled'] = True
            item['errors'] = ['planning:cancelled']
            break
        except Exception as exc:
            item['errors'] = [f'planning:{type(exc).__name__}']
        if progress:
            progress(index + 1, len(files), item)
    return plan
