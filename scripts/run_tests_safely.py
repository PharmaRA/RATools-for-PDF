"""Run unittest with temporary files isolated and cleanup sent to the recycle bin.

Usage: python -B scripts/run_tests_safely.py [test_filename_pattern]
The audit hook rejects any physical Python deletion that bypasses the adapters.
"""

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


def main():
    if os.name != 'nt':
        raise SystemExit('This safety runner currently requires Windows recycle-bin support.')
    run_dir = ROOT / 'backups' / 'test-runs' / uuid.uuid4().hex
    run_dir.mkdir(parents=True)
    previous_temp = tempfile.tempdir
    # Windows spawn re-executes this script as __mp_main__; install guards there too.
    os.environ['RATOOLS_SAFE_TEST_ROOT'] = str(run_dir.resolve())
    stack, failures, recycle = install_cleanup_guard(run_dir)
    result = None
    try:
        with stack:
            suite = unittest.defaultTestLoader.discover(
                str(ROOT / 'tests'), pattern=sys.argv[1] if len(sys.argv) > 1 else 'test_*.py',
            )
            result = unittest.TextTestRunner(verbosity=2).run(suite)
    finally:
        tempfile.tempdir = previous_temp
        os.environ.pop('RATOOLS_SAFE_TEST_ROOT', None)
        recycle(run_dir)
    return 0 if result and result.wasSuccessful() and not failures else 1


if __name__ == '__main__':
    raise SystemExit(main())
elif __name__ == '__mp_main__' and os.environ.get('RATOOLS_SAFE_TEST_ROOT'):
    _child_guard = install_cleanup_guard(Path(os.environ['RATOOLS_SAFE_TEST_ROOT']))
