"""Portable PD02 facts, unkeyed package integrity, and a callable verifier.

No PDF or localized log is read. COMPLETE.json is written last; incomplete
directories are deliberately retained and can never pass verification.
"""

import argparse
from copy import deepcopy
import csv
from datetime import datetime, timezone
import hashlib
import json
from pathlib import Path, PurePosixPath, PureWindowsPath
import re

from ratools_pdf.config.version import APP_VERSION_STR
from ratools_pdf.services.file_safety import check_unlinked_path


AUDIT_SCHEMA = 'pd03.audit.v1'
INTEGRITY_SCHEMA = 'pd03.integrity.v1'
COMPLETE_SCHEMA = 'pd03.complete.v1'
PAYLOADS = {'evidence.json', 'summary.csv'}
PACKAGE_FILES = PAYLOADS | {'integrity.json', 'COMPLETE.json'}
CSV_FIELDS = ('batch_id', 'file_id', 'input_relative_path', 'input_sha256',
              'output_relative_path', 'output_sha256', 'outcome', 'engine_success',
              'published', 'save_reopen', 'structural_expectations',
              'rule_verification', 'reason_codes', 'reasons', 'warnings', 'missing_evidence')
_OUTCOME_REASONS = {'processing_failed': 'engine_failure',
                    'verification_failed': 'verification_failed',
                    'planning_failed': 'planning_failed', 'stale_plan': 'stale_plan',
                    'publish_failed': 'publish_failed', 'stopped': 'stop_requested',
                    'skipped': 'skip_requested', 'not_started': 'not_started',
                    'timed_out': 'file_timeout', 'resource_preflight_failed': 'resource_preflight_failed'}
_METRICS = ('sha256', 'bytes', 'page_count', 'pages', 'bookmarks', 'links',
            'annotations', 'attachments', 'forms', 'signature_fields',
            'signature_present', 'pdf_version', 'reopen', 'errors')
_VERIFICATION = ('save_reopen', 'structural_expectations', 'rule_verification',
                 'changes', 'publish_allowed', 'signature_validity')


class AuditExportError(RuntimeError):
    """A new export directory remains incomplete; raw exception text is private."""


def _schema(value, expected):
    if not isinstance(value, dict) or value.get('schema_version') != expected:
        raise ValueError('Unsupported or missing evidence schema')


def _identity(value):
    if not isinstance(value, str) or not value or '\\' in value or '\x00' in value:
        raise ValueError('Invalid relative file identity')
    posix, windows = PurePosixPath(value), PureWindowsPath(value)
    if (posix.is_absolute() or windows.drive or windows.root or
            any(part in {'', '.', '..'} for part in value.split('/'))):
        raise ValueError('Invalid relative file identity')
    return value


def _identifier(value):
    if not isinstance(value, str) or not re.fullmatch(r'[A-Za-z0-9_-]{1,128}', value):
        raise ValueError('Invalid evidence identifier')
    return value


def _snapshot(value):
    if value is None:
        return None
    _schema(value, 'pd02.snapshot.v1')
    result = {key: deepcopy(value.get(key)) for key in _METRICS}
    result['schema_version'] = 'pd02.snapshot.v1'
    # Structure only: never copy metadata, text, annotations or widget bodies.
    if result['pages'] is not None:
        result['pages'] = [{key: page.get(key) for key in ('width', 'height', 'rotation')}
                           for page in result['pages']]
    return result


def _verification(value):
    if value is None:
        return None
    _schema(value, 'pd02.verification.v1')
    result = {key: deepcopy(value.get(key)) for key in _VERIFICATION}
    result['schema_version'] = 'pd02.verification.v1'
    if isinstance(result['changes'], dict):
        result['changes'] = {key: val for key, val in result['changes'].items()
                             if key in _METRICS}
    return result


def _runtime_metadata():
    import pymupdf
    return {
        'observation': 'export_time_not_processing_time',
        'application': {'version': APP_VERSION_STR, 'source_build': 'unknown'},
        'pymupdf': {'version': str(pymupdf.VersionBind)},
        'qpdf': {'version': 'unknown', 'reason': 'not_recorded_by_pd02'},
        'processing_runtime': 'unknown_not_recorded_by_pd02',
    }


def _execution_runtime(evidence):
    value = evidence.get('execution_runtime')
    if not isinstance(value, dict):
        return {
            'application': {'version': 'unknown', 'source_build': 'unknown'},
            'pymupdf': {'version': 'unknown'},
            'qpdf': {'version': 'unknown', 'reason': 'not_recorded_by_pd02'},
        }
    return {
        'application': {'version': value.get('application', {}).get('version') or 'unknown',
                        'source_build': value.get('application', {}).get('source_build') or 'unknown'},
        'pymupdf': {'version': value.get('pymupdf', {}).get('version') or 'unknown'},
        'qpdf': {'version': value.get('qpdf', {}).get('version') or 'unknown',
                 'reason': value.get('qpdf', {}).get('reason') or 'not_recorded_by_pd02'},
    }


def _build_document(evidence):
    evidence = deepcopy(evidence)
    _schema(evidence, 'pd02.batch.v1')
    if not evidence.get('ended_at') or evidence.get('outcome') in {'running', 'previewed'}:
        raise ValueError('Select a terminal processing batch')
    batch_id = _identifier(evidence.get('batch_id'))
    plan = evidence.get('plan')
    _schema(plan, 'pd02.plan.v1')
    if plan.get('batch_id') != batch_id:
        raise ValueError('Batch and plan identities differ')
    config = plan.get('configuration', {})
    if config.get('version') != 'pd02.engine.v1':
        raise ValueError('Unsupported processing configuration')
    configuration = {key: deepcopy(config.get(key)) for key in
                     ('version', 'options', 'processing_mode')}
    configuration['options'] = [_identifier(option) for option in configuration['options']]
    configuration['compression_settings'] = {key: config.get('compression_settings', {}).get(key)
                                              for key in ('dpi', 'quality')
                                              if key in config.get('compression_settings', {})}
    if 'profile' in config or 'profile_digest' in config:
        from ratools_pdf.services.processing_plan import configuration_matches_plan
        if not configuration_matches_plan(plan, config.get('options'), config.get('processing_mode'),
                                          config.get('compression_settings'), profile=config.get('profile')):
            raise ValueError('Invalid profile snapshot or configuration digest')
        configuration['profile'] = deepcopy(config['profile'])
        configuration['profile_digest'] = config['profile_digest']
    files, plans = [], []
    for source in plan.get('files', []):
        file_id = _identifier(source.get('file_id'))
        relative = _identity(source.get('relative_path'))
        fact = evidence.get('files', {}).get(file_id, {})
        if fact.get('relative_path', relative) != relative:
            raise ValueError('Conflicting file identity')
        before = _snapshot(fact.get('before') or source.get('before'))
        after = _snapshot(fact.get('after'))
        verification = _verification(fact.get('verification'))
        rules = [{key: deepcopy(rule.get(key)) for key in
                  ('option_id', 'precheck', 'decision', 'reason', 'actual_changed', 'verified')}
                 for rule in source.get('rules', [])]
        for rule in rules:
            _identifier(rule['option_id'])
        rules_plan = {'file_id': file_id, 'relative_path': relative,
                      'status': source.get('status'), 'rules': rules,
                      'effective_options': [_identifier(option) for option in source.get('effective_options', [])],
                      'risks': list(source.get('risks', [])), 'errors': list(source.get('errors', [])),
                      'before': _snapshot(source.get('before'))}
        plans.append(rules_plan)
        published = fact.get('published') is True
        output_relative = fact.get('output_relative_path')
        if output_relative is not None:
            output_relative = _identity(output_relative)
        missing = []
        outcome = fact.get('outcome', 'unknown')
        reason_codes = [_OUTCOME_REASONS[outcome]] if outcome in _OUTCOME_REASONS else []
        if outcome != 'succeeded':
            missing.append('engine_diagnostic:not_recorded')
        if before is None or not before.get('sha256'):
            missing.append('input:fingerprint_missing')
        if verification is None:
            missing.append('verification:not_recorded')
        if not published:
            missing.append('output:not_published')
        elif output_relative is None or after is None or not after.get('sha256'):
            missing.append('output:identity_or_fingerprint_missing')
        reasons = list(source.get('errors', [])) + [rule['reason'] for rule in rules if rule.get('reason')]
        files.append({
            'file_id': file_id,
            'input': {'relative_path': relative, 'sha256': (before or {}).get('sha256'),
                      'bytes': (before or {}).get('bytes')},
            'output': ({'relative_path': output_relative, 'sha256': (after or {}).get('sha256'),
                        'bytes': (after or {}).get('bytes')} if published else None),
            'before': before, 'after': after, 'verification': verification,
            'outcome': outcome, 'engine_success': fact.get('engine_success'),
            'published': published, 'rule_actual_changed': fact.get('rule_actual_changed', 'unknown'),
            'started_at': fact.get('started_at'), 'ended_at': fact.get('ended_at'),
            'rules': rules, 'reason_codes': reason_codes, 'reasons': reasons,
            'warnings': list(source.get('risks', [])),
            'missing_evidence': missing,
        })
    if len({item['file_id'] for item in files}) != len(files):
        raise ValueError('Duplicate file identities')
    events = []
    for event in evidence.get('events', []):
        if event.get('batch_id') != batch_id or event.get('file_id') not in {item['file_id'] for item in files}:
            raise ValueError('Event belongs to another batch')
        entry = {key: deepcopy(event.get(key)) for key in
                 ('type', 'batch_id', 'file_id', 'time', 'started_at', 'ended_at',
                  'engine_success', 'published', 'outcome', 'rule_actual_changed') if key in event}
        entry.update(before=_snapshot(event.get('before')), after=_snapshot(event.get('after')),
                     verification=_verification(event.get('verification')))
        if event.get('output_relative_path') is not None:
            entry['output_relative_path'] = _identity(event['output_relative_path'])
        events.append(entry)
    limitations = ['hashes_are_not_signatures', 'rule_changes_unknown_until_pd07',
                   'signature_validity_not_verified', 'pdf_payloads_not_included',
                   'engine_warnings_not_recorded_by_pd02']
    if not isinstance(evidence.get('execution_runtime'), dict):
        limitations.append('processing_runtime_not_recorded')
    else:
        limitations.append('qpdf_execution_version_not_recorded')
    return {
        'schema_version': AUDIT_SCHEMA, 'source_schema_version': 'pd02.batch.v1',
        'exported_at': datetime.now(timezone.utc).isoformat(),
        'batch': {key: evidence.get(key) for key in
                  ('batch_id', 'created_at', 'started_at', 'ended_at', 'outcome')},
        'runtime': _runtime_metadata(), 'processing_runtime': _execution_runtime(evidence),
        'configuration': configuration,
        'plan': {'schema_version': 'pd02.plan.v1', 'batch_id': batch_id,
                 'configuration_digest': plan.get('configuration_digest'),
                 'cancelled': plan.get('cancelled'), 'files': plans},
        'files': files, 'events': events,
        'csv_encoding': 'utf-8-sig; apostrophe-prefix-v1; list fields joined by LF; JSON preserves list boundaries',
        'limitations': limitations,
    }


def _csv_cell(value):
    text = '' if value is None else str(value)
    if text.startswith("'") or text.lstrip().startswith(('=', '+', '-', '@')):
        return "'" + text
    return text


def decode_csv_cell(text):
    """Undo exactly one safety prefix; use only on cells from this CSV schema."""
    if text.startswith("'") and (text[1:].startswith("'") or
                                 text[1:].lstrip().startswith(('=', '+', '-', '@'))):
        return text[1:]
    return text


def _csv_rows(document):
    for item in document['files']:
        original, output = item['input'], item['output'] or {}
        verification = item['verification'] or {}
        row = {'batch_id': document['batch']['batch_id'], 'file_id': item['file_id'],
               'input_relative_path': original['relative_path'], 'input_sha256': original['sha256'],
               'output_relative_path': output.get('relative_path'), 'output_sha256': output.get('sha256'),
               **{key: item[key] for key in ('outcome', 'engine_success', 'published')},
               **{key: verification.get(key, 'not_verified') for key in
                  ('save_reopen', 'structural_expectations', 'rule_verification')},
               **{key: '\n'.join(item[key]) for key in ('reason_codes', 'reasons', 'warnings', 'missing_evidence')}}
        yield {key: _csv_cell(row.get(key)) for key in CSV_FIELDS}


def _write_json(path, value):
    with path.open('x', encoding='utf-8', newline='\n') as stream:
        json.dump(value, stream, ensure_ascii=False, sort_keys=True, indent=2, allow_nan=False)
        stream.write('\n')


def _hash(path):
    digest = hashlib.sha256()
    with path.open('rb') as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b''):
            digest.update(chunk)
    return digest.hexdigest()


def export_evidence_package(evidence, destination):
    """Export to a NEW directory. Failures retain a visibly incomplete package."""
    document = _build_document(evidence)
    target = check_unlinked_path(destination)
    target.mkdir()  # existing destinations are never opened or reused
    try:
        _write_json(target / 'evidence.json', document)
        with (target / 'summary.csv').open('x', encoding='utf-8-sig', newline='') as stream:
            writer = csv.DictWriter(stream, fieldnames=CSV_FIELDS)
            writer.writeheader()
            writer.writerows(_csv_rows(document))
        manifest = {'schema_version': INTEGRITY_SCHEMA, 'algorithm': 'sha256',
                    'files': [{'path': name, 'sha256': _hash(target / name),
                               'bytes': (target / name).stat().st_size} for name in sorted(PAYLOADS)]}
        _write_json(target / 'integrity.json', manifest)
        _write_json(target / 'COMPLETE.json', {'schema_version': COMPLETE_SCHEMA, 'status': 'complete',
                                              'batch_id': document['batch']['batch_id'],
                                              'integrity_sha256': _hash(target / 'integrity.json')})
    except Exception:
        raise AuditExportError('Evidence export failed; the new directory is incomplete and must not be archived.') from None
    return target


def _validate_document(document):
    _schema(document, AUDIT_SCHEMA)
    if document.get('source_schema_version') != 'pd02.batch.v1':
        raise ValueError('Unsupported source schema')
    _schema(document.get('plan'), 'pd02.plan.v1')
    if document.get('configuration', {}).get('version') != 'pd02.engine.v1':
        raise ValueError('Unsupported configuration schema')
    config = document['configuration']
    if 'profile' in config or 'profile_digest' in config:
        from ratools_pdf.services.processing_plan import configuration_matches_plan
        binding = {'configuration': config, 'configuration_digest': document['plan'].get('configuration_digest')}
        if not configuration_matches_plan(binding, config.get('options'), config.get('processing_mode'),
                                          config.get('compression_settings'), profile=config.get('profile')):
            raise ValueError('Invalid profile snapshot or configuration digest')
    batch = document['batch']
    _identifier(batch['batch_id'])
    if not batch.get('ended_at') or batch.get('outcome') in {'running', 'previewed'}:
        raise ValueError('Nonterminal evidence')
    if document['plan']['batch_id'] != batch['batch_id']:
        raise ValueError('Batch mismatch')
    for component in ('application', 'qpdf', 'pymupdf'):
        if not document['runtime'][component]['version']:
            raise ValueError('Missing runtime version')
    for item in document['files']:
        _identifier(item['file_id'])
        _identity(item['input']['relative_path'])
        if item.get('output') and item['output'].get('relative_path') is not None:
            _identity(item['output']['relative_path'])
        _snapshot(item['before'])
        _snapshot(item['after'])
        _verification(item['verification'])
    for item in document['plan']['files']:
        _identity(item['relative_path'])
        _snapshot(item['before'])
    for item in document['events']:
        _snapshot(item['before'])
        _snapshot(item['after'])
        _verification(item['verification'])
        if item.get('output_relative_path') is not None:
            _identity(item['output_relative_path'])


def verify_evidence_package(directory):
    """Reject incomplete/tampered/extra/escaping/link payloads; never follow links."""
    try:
        root = check_unlinked_path(directory)
        entries = list(root.iterdir())
        if {entry.name for entry in entries} != PACKAGE_FILES:
            raise ValueError('Package is incomplete or has unexpected payloads')
        for entry in entries:
            check_unlinked_path(entry)
            if not entry.is_file():
                raise ValueError('Unexpected payload type')
        marker = json.loads((root / 'COMPLETE.json').read_text(encoding='utf-8'))
        _schema(marker, COMPLETE_SCHEMA)
        if marker.get('status') != 'complete' or marker.get('integrity_sha256') != _hash(root / 'integrity.json'):
            raise ValueError('Invalid completion marker')
        manifest = json.loads((root / 'integrity.json').read_text(encoding='utf-8'))
        _schema(manifest, INTEGRITY_SCHEMA)
        if manifest.get('algorithm') != 'sha256' or len(manifest['files']) != len(PAYLOADS):
            raise ValueError('Invalid integrity manifest')
        names = []
        for entry in manifest['files']:
            name = _identity(entry['path'])
            if name not in PAYLOADS or name in names:
                raise ValueError('Unexpected integrity path')
            names.append(name)
            if entry['sha256'] != _hash(root / name) or entry['bytes'] != (root / name).stat().st_size:
                raise ValueError('Payload hash or size mismatch')
        document = json.loads((root / 'evidence.json').read_text(encoding='utf-8'))
        _validate_document(document)
        if marker.get('batch_id') != document['batch']['batch_id']:
            raise ValueError('Completion identity mismatch')
        with (root / 'summary.csv').open(encoding='utf-8-sig', newline='') as stream:
            reader = csv.DictReader(stream)
            if tuple(reader.fieldnames or []) != CSV_FIELDS or list(reader) != list(_csv_rows(document)):
                raise ValueError('CSV and JSON facts differ')
        return {'schema_version': 'pd03.verify.v1', 'valid': True, 'errors': [],
                'authentication': 'not_signed'}
    except Exception:
        return {'schema_version': 'pd03.verify.v1', 'valid': False,
                'errors': ['incomplete_invalid_or_tampered_package'], 'authentication': 'not_signed'}


def main(argv=None):
    parser = argparse.ArgumentParser(description='Verify a RATools audit evidence package (integrity, not authenticity).')
    parser.add_argument('action', choices=['verify'])
    parser.add_argument('directory')
    args = parser.parse_args(argv)
    result = verify_evidence_package(args.directory)
    print(json.dumps(result))
    return 0 if result['valid'] else 1


if __name__ == '__main__':
    raise SystemExit(main())
