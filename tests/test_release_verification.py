"""Exercise real source startup/PDF/qpdf/codec through the public release CLI."""

import json
import importlib.util
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest


ROOT = Path(__file__).resolve().parents[1]


class ReleaseVerificationTests(unittest.TestCase):
    def test_both_entrypoints_emit_complete_isolated_release_checks(self):
        self.assertIsNotNone(importlib.util.find_spec('ratools_pdf.services.release_check'),
                             'Unified release selfcheck is not implemented')
        with tempfile.TemporaryDirectory() as directory:
            for entry, variant in [('main.py', 'normal'), ('main_no_update.py', 'no-update')]:
                with self.subTest(variant=variant):
                    target = Path(directory)/variant
                    env = dict(os.environ, QT_QPA_PLATFORM='offscreen', RATOOLS_ENABLE_UPDATE_CHECK='1')
                    process = subprocess.run([sys.executable, '-B', str(ROOT/entry), '--verify-release', str(target)],
                        cwd=ROOT, env=env, capture_output=True, timeout=90,
                        creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0)
                    self.assertTrue((target/'release-verification.json').exists(),
                                    'Public release selfcheck did not produce its report')
                    self.assertEqual(process.returncode, 0, process.stderr.decode('utf-8', errors='replace')[-1500:])
                    report = json.loads((target/'release-verification.json').read_text(encoding='utf-8'))
                    self.assertTrue(report['ok'], report)
                    self.assertEqual(report['variant'], variant)
                    for name in ('gui', 'core_pdf', 'qpdf', 'image_compression', 'update_boundary'):
                        self.assertEqual(report['checks'][name]['status'], 'passed', report)
                    self.assertEqual(report['checks']['gui']['settings'], 'settings/settings.ini')
                    self.assertTrue((target/'settings/settings.ini').exists())
                    boundary = report['checks']['update_boundary']
                    self.assertEqual(boundary['network_attempts'], 0)
                    self.assertEqual(boundary['update_entry'], variant == 'normal')
                    if variant == 'no-update':
                        self.assertFalse(boundary['updater_imported'])
                        self.assertFalse(boundary['worker_created'])
                    original = (target/'release-verification.json').read_bytes()
                    repeated = subprocess.run([sys.executable, '-B', str(ROOT/entry), '--verify-release', str(target)],
                        cwd=ROOT, env=env, capture_output=True, timeout=30,
                        creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0)
                    self.assertNotEqual(repeated.returncode, 0)
                    self.assertEqual((target/'release-verification.json').read_bytes(), original)
