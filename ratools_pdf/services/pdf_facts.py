"""Content-free PDF structure facts and the staged-output publication gate."""

import hashlib
import re
from pathlib import Path

import fitz

from ratools_pdf.pdf import inspect as pdf_inspect


STRUCTURE_FIELDS = ('page_count', 'pages', 'bookmarks', 'links', 'annotations',
                    'attachments', 'forms', 'signature_fields', 'signature_present', 'pdf_version')


class InspectionCancelled(Exception):
    pass


_DYNAMIC_SUBTYPES = ('RichMedia', 'Movie', 'Sound', '3D')


def _count_dynamic_page_annotations(doc, page, seen_xrefs=None):
    """Count dynamic page annotations not already enumerated by Page.annots()."""
    seen_xrefs = set(seen_xrefs or ())
    kind, value = doc.xref_get_key(page.xref, 'Annots')
    if kind != 'array':
        return 0
    count = 0
    # Direct annotation dictionaries are uncommon but may be embedded in the
    # page array; count each dynamic subtype token once.
    for subtype in _DYNAMIC_SUBTYPES:
        count += len(re.findall(r'/Subtype\s*/' + re.escape(subtype) + r'\b', value))
    for ref in re.findall(r'(\d+)\s+0\s+R', value):
        xref = int(ref)
        if xref in seen_xrefs:
            continue
        try:
            obj = doc.xref_object(xref, compressed=False)
        except Exception:
            continue
        if any(re.search(r'/Subtype\s*/' + re.escape(subtype) + r'\b', obj)
               for subtype in _DYNAMIC_SUBTYPES):
            count += 1
    return count


def fingerprint(path, cancelled=None):
    digest = hashlib.sha256()
    size = 0
    with open(path, 'rb') as stream:
        while True:
            if cancelled and cancelled():
                raise InspectionCancelled()
            chunk = stream.read(1024 * 1024)
            if not chunk:
                break
            digest.update(chunk)
            size += len(chunk)
    return {'sha256': digest.hexdigest(), 'bytes': size}


def capture_pdf_snapshot(path, cancelled=None):
    """Read metrics only. Failed inspection is null, never an invented zero."""
    result = dict.fromkeys(('sha256', 'bytes') + STRUCTURE_FIELDS)
    result.update(schema_version='pd02.snapshot.v1', reopen='not_verified', errors=[])

    def read(name, reader):
        try:
            result[name] = reader()
        except InspectionCancelled:
            raise
        except Exception as exc:
            result['errors'].append(f'{name}:{type(exc).__name__}')

    try:
        initial = Path(path).stat()
        result.update(fingerprint(path, cancelled))
        with fitz.open(path) as doc:
            if doc.needs_pass:
                result['reopen'] = 'failed'
                result['errors'].append('open:encrypted')
                return result
            result['reopen'] = 'passed'

            def pages_metric(metric):
                values = []
                for page in doc:
                    if cancelled and cancelled():
                        raise InspectionCancelled()
                    values.append(metric(page))
                return values

            read('page_count', lambda: doc.page_count)
            read('pages', lambda: pages_metric(lambda page: {
                'width': round(page.rect.width, 4), 'height': round(page.rect.height, 4),
                'rotation': page.rotation}))
            read('bookmarks', lambda: len(doc.get_toc()))
            read('links', lambda: sum(pages_metric(lambda page: len(page.get_links()))))
            def annotation_count(page):
                enumerated = list(page.annots() or [])
                seen = {getattr(item, 'xref', None) for item in enumerated}
                seen.discard(None)
                return len(enumerated) + _count_dynamic_page_annotations(doc, page, seen)
            read('annotations', lambda: sum(pages_metric(annotation_count)))
            read('attachments', doc.embfile_count)
            read('forms', lambda: sum(pages_metric(lambda page: sum(1 for _ in page.widgets() or []))))
            read('signature_fields', lambda: sum(pages_metric(lambda page: sum(
                1 for widget in page.widgets() or [] if widget.field_type == fitz.PDF_WIDGET_TYPE_SIGNATURE))))

            def has_signature():
                if doc.get_sigflags() > 0:
                    return True
                return any(pages_metric(lambda page: any(
                    widget.field_type == fitz.PDF_WIDGET_TYPE_SIGNATURE and
                    doc.xref_get_key(widget.xref, 'V')[0] != 'null'
                    for widget in page.widgets() or [])))

            read('signature_present', has_signature)
            read('pdf_version', lambda: pdf_inspect.read_pdf_header_version(path) or None)
        final = Path(path).stat()
        if (initial.st_size, initial.st_mtime_ns) != (final.st_size, final.st_mtime_ns):
            result['errors'].append('input:changed_during_inspection')
    except InspectionCancelled:
        raise
    except Exception as exc:
        result['reopen'] = 'failed'
        result['errors'].append(f'open:{type(exc).__name__}')
    return result


def verify_output(before, after, options):
    """Compare structural expectations separately from reopen and rule qualification."""
    selected = set(options)
    allowed = {
        'pages': {'page_size_a4', 'page_size_letter', 'flatten_rotation'},
        'bookmarks': {'bookmark_remove_external_links', 'bookmark_remove_invalid',
                      'bookmark_remove_unknown_actions', 'cleanup_remove_all_links_bookmarks',
                      'page_size_a4', 'page_size_letter'},
        'links': {'cleanup_remove_external_uri', 'cleanup_remove_external_uri_and_text_black',
                  'cleanup_remove_invalid_links', 'cleanup_remove_invalid_links_and_text_black',
                  'cleanup_remove_unknown_action_links', 'cleanup_remove_all_links_bookmarks',
                  'page_size_a4', 'page_size_letter'},
        'annotations': {'cleanup_remove_annotations', 'cleanup_remove_dynamic_content',
                        'flatten_annotations', 'page_size_a4', 'page_size_letter'},
        'attachments': {'cleanup_remove_attachments'},
        'forms': {'page_size_a4', 'page_size_letter'},
        'signature_fields': {'page_size_a4', 'page_size_letter'},
        'signature_present': {'page_size_a4', 'page_size_letter'},
        'pdf_version': {'convert_pdf_version', 'generate_object_streams', 'compress_standard',
                        'compress_aggressive', 'compress_images'},
    }
    changes = {}
    for field in STRUCTURE_FIELDS:
        original, output = before.get(field), after.get(field)
        if original is None or output is None:
            changes[field] = 'unknown'
        elif original == output:
            changes[field] = 'unchanged'
        elif field == 'pdf_version':
            # MuPDF/qpdf may upgrade the header while serializing an otherwise
            # valid document; the version transition alone is not corruption.
            changes[field] = 'expected_change'
        elif selected & allowed.get(field, set()):
            changes[field] = 'expected_change'
        else:
            changes[field] = 'unexpected_change'
    if 'unexpected_change' in changes.values() or after.get('page_count') == 0:
        expectation = 'failed'
    elif 'unknown' in changes.values() or before.get('errors') or after.get('errors'):
        expectation = 'not_verified'
    else:
        expectation = 'passed'
    reopen = after.get('reopen', 'not_verified')
    return {
        'schema_version': 'pd02.verification.v1',
        'save_reopen': reopen,
        'structural_expectations': expectation,
        'rule_verification': 'not_verified',
        'changes': changes,
        'publish_allowed': reopen == 'passed' and expectation == 'passed',
        'signature_validity': 'not_verified',
    }
