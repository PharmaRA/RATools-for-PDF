"""Run unittest with temporary files isolated and cleanup sent to the recycle bin.

Usage: python -B scripts/run_tests_safely.py [test_filename_pattern] [--layer LAYER]
The audit hook rejects any physical Python deletion that bypasses the adapters.
"""

import argparse
import json
import os
from pathlib import Path
import shutil
import sys
import tempfile
import unittest
import uuid
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.dont_write_bytecode = True
os.environ['QT_QPA_PLATFORM'] = 'offscreen'
from ratools_pdf.services.file_safety import recycle_owned_path
from scripts.test_layers import (
    LayerResult, empty_summary, finish_summary, install_unit_guard,
    load_classification, select_suite, summarize_tests,
)


def install_cleanup_guard(run_dir):
    from contextlib import ExitStack

    run_dir = run_dir.resolve()
    expected_parent = (ROOT / 'backups' / 'test-runs').resolve()
    if run_dir.parent != expected_parent or not run_dir.is_dir():
        raise RuntimeError('Invalid test-run directory')
    failures = []

    def recycle(path, *args, **kwargs):
        candidate = Path(os.fsdecode(path)).absolute()
        resolved = candidate.resolve()
        if not resolved.is_relative_to(run_dir):
            raise RuntimeError('Refusing cleanup outside the current test run')
        if not candidate.exists():
            return
        try:
            recycle_owned_path(candidate, run_dir)
        except Exception as exc:
            failures.append(type(exc).__name__)
            raise

    def audit(event, args):
        if event in {'os.remove', 'os.rmdir', 'shutil.rmtree'}:
            raise RuntimeError('Physical deletion is disabled for this test run')

    sys.addaudithook(audit)
    tempfile.tempdir = str(run_dir)
    stack = ExitStack()
    for name in ('remove', 'unlink', 'rmdir'):
        stack.enter_context(patch.object(os, name, recycle))
    stack.enter_context(patch.object(shutil, 'rmtree', recycle))
    return stack, failures, recycle


def shutdown_test_qt_application():
    """Release test-owned Qt objects while Python modules are still intact."""
    # Pure unit tests must not import Qt just to perform cleanup.
    widgets = sys.modules.get('PySide6.QtWidgets')
    if widgets is None:
        return
    app = widgets.QApplication.instance()
    if app is None:
        return
    print('QT_CLEANUP starting', flush=True)
    from PySide6.QtCore import QCoreApplication, QEvent
    # Tests do not run app.exec(); deleteLater() therefore needs an explicit
    # drain before QApplication tears down its platform resources.
    for widget in app.topLevelWidgets():
        widget.deleteLater()
    QCoreApplication.sendPostedEvents(None, QEvent.Type.DeferredDelete)
    print('QT_CLEANUP widgets_released', flush=True)
    app.shutdown()
    print('QT_CLEANUP completed', flush=True)


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('pattern', nargs='?', default='test_*.py')
    parser.add_argument('--layer', choices=('all', 'unit', 'integration', 'qualification'), default='all',
        help='all runs the full ordinary unit/integration regression; qualification is explicit')
    parser.add_argument('--tests-dir', type=Path, default=ROOT / 'tests')
    parser.add_argument('--classification', type=Path, default=ROOT / 'tests/test-layers-v1.json')
    parser.add_argument('--qualification-manifest', type=Path, default=ROOT / 'docs/qualification/pd01-manifest-v1.json')
    parser.add_argument('--qualification-output', type=Path, help='New evidence directory; required for qualification')
    parser.add_argument('--samples-root', type=Path)
    parser.add_argument('--skip-qpdf', action='store_true')
    args = parser.parse_args(argv)
    if os.name != 'nt':
        raise SystemExit('This safety runner currently requires Windows recycle-bin support.')
    run_dir = ROOT / 'backups' / 'test-runs' / uuid.uuid4().hex
    run_dir.mkdir(parents=True)
    previous_temp = tempfile.tempdir
    # Windows spawn re-executes this script as __mp_main__; install guards there too.
    os.environ['RATOOLS_SAFE_TEST_ROOT'] = str(run_dir.resolve())
    stack, failures, recycle = install_cleanup_guard(run_dir)
    code = 2
    summary = empty_summary()
    try:
        with stack:
            try:
                assignments = load_classification(args.classification, args.tests_dir)
                if args.layer == 'qualification':
                    if args.qualification_output is None:
                        raise ValueError('qualification requires --qualification-output (new directory)')
                    from ratools_pdf.services.qualification import run_qualification
                    report = run_qualification(args.qualification_manifest, args.qualification_output,
                        samples_root=args.samples_root, qpdf_enabled=not args.skip_qpdf)
                    item = summary['layers']['qualification']
                    item.update(status=report['status'], **report['summary'])
                    # Blocked samples are missing prerequisites; matrix rows are not added again.
                    item['pending'] += item.pop('blocked', 0)
                    item['matrix_statuses'] = {status: sum(row['status'] == status for row in report['matrix'])
                        for status in ('passed', 'failed', 'pending', 'blocked')}
                    finish_summary(summary)
                    code = 0 if report['status'] == 'passed' else 1 if report['status'] == 'failed' else 2
                else:
                    selected = ('unit', 'integration') if args.layer == 'all' else (args.layer,)
                    if args.layer == 'unit':
                        install_unit_guard()
                    suite, test_layers = select_suite(args.tests_dir, assignments, args.pattern, selected)
                    result = unittest.TextTestRunner(verbosity=2, resultclass=LayerResult).run(suite)
                    summarize_tests(summary, result, test_layers, selected)
                    code = 0 if result.wasSuccessful() else 1
            except (ValueError, OSError, json.JSONDecodeError) as exc:
                print(f'Classification/selection/qualification error: {exc}', file=sys.stderr)
                code = 2
    finally:
        try:
            shutdown_test_qt_application()
        except Exception as exc:
            failures.append(type(exc).__name__)
            print(f'Qt cleanup failed: {type(exc).__name__}: {exc}', file=sys.stderr, flush=True)
        tempfile.tempdir = previous_temp
        os.environ.pop('RATOOLS_SAFE_TEST_ROOT', None)
        recycle(run_dir)
    print('LAYER_SUMMARY ' + json.dumps(summary, sort_keys=True))
    return 1 if failures else code


if __name__ == '__main__':
    raise SystemExit(main())
elif __name__ == '__mp_main__' and os.environ.get('RATOOLS_SAFE_TEST_ROOT'):
    _child_guard = install_cleanup_guard(Path(os.environ['RATOOLS_SAFE_TEST_ROOT']))
