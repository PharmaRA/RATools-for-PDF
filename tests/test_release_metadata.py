"""Release boundaries: tampering, incomplete locks and unrecorded dirty sources."""

import copy
import hashlib
import importlib.util
import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch
import sys
import importlib.util
import marshal
import zipfile
import subprocess


ROOT = Path(__file__).resolve().parents[1]


class ReleaseMetadataTests(unittest.TestCase):
    def api(self):
        self.assertIsNotNone(importlib.util.find_spec('scripts.release_metadata'),
                             'Release metadata/lock verification is not implemented')
        from scripts import release_metadata
        return release_metadata

    def test_lock_rejects_invalid_digest_and_missing_transitive_dependency(self):
        api = self.api()
        lock = api.load_lock(ROOT / 'requirements-release-win311.lock.json')
        broken = copy.deepcopy(lock)
        broken['packages'][0]['sha256'] = 'bad'
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'lock.json'
            path.write_text(json.dumps(broken), encoding='utf-8')
            with self.assertRaises(ValueError):
                api.load_lock(path)
        incomplete = copy.deepcopy(lock)
        incomplete['packages'] = [p for p in lock['packages'] if p['name'] != 'shiboken6']
        with self.assertRaisesRegex(ValueError, 'shiboken6'):
            api.validate_dependency_closure(incomplete)

    def test_installed_version_mismatch_fails_closed(self):
        api = self.api()
        lock = api.load_lock(ROOT / 'requirements-release-win311.lock.json')
        lock['packages'] = [p for p in lock['packages'] if p['name'] == 'Pillow']
        lock['packages'][0]['version'] = '0.0.0'
        with self.assertRaisesRegex(ValueError, 'Pillow'):
            api.verify_environment(lock)

    def test_source_provenance_includes_actual_dirty_theme_bytes(self):
        api = self.api()
        report = api.source_provenance(ROOT)
        self.assertRegex(report['commit'], r'^[0-9a-f]{40}$')
        self.assertEqual(report['sources']['ratools_pdf/ui/theme.py'],
                         hashlib.sha256((ROOT/'ratools_pdf/ui/theme.py').read_bytes()).hexdigest())
        self.assertIsInstance(report['dirty'], bool)
        self.assertIn('ratools_pdf/ui/theme.py', report['dirty_paths'])

    def test_artifact_verification_detects_changed_bytes_and_extra_files(self):
        api = self.api()
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root/'app.exe').write_bytes(b'artifact bytes')
            api.write_hash_manifest(root)
            self.assertTrue(api.verify_hash_manifest(root))
            (root/'app.exe').write_bytes(b'tampered bytes')
            self.assertFalse(api.verify_hash_manifest(root))
            (root/'app.exe').write_bytes(b'artifact bytes')
            (root/'unexpected.dll').write_bytes(b'unrecorded')
            self.assertFalse(api.verify_hash_manifest(root))

    def test_signing_plan_is_unsigned_until_explicit_execution(self):
        api = self.api()
        self.assertEqual(api.signing_plan()['status'], 'unsigned')
        plan = api.signing_plan(certificate_env='RATOOLS_RELEASE_CERT_SHA1')
        self.assertEqual(plan['status'], 'pending')
        self.assertEqual(plan['certificate_environment'], 'RATOOLS_RELEASE_CERT_SHA1')
        with self.assertRaises(ValueError):
            api.signing_plan(certificate_env='BAD;NAME')

    def test_missing_empty_or_unhashed_installed_record_fails_closed(self):
        api = self.api()
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            info = root/'release_fixture-1.0.dist-info'
            info.mkdir()
            (info/'METADATA').write_text('Name: release-fixture\nVersion: 1.0\n',encoding='utf-8')
            (root/'release_fixture.py').write_text('value = 1\n',encoding='utf-8')
            lock = {'packages':[{'name':'release-fixture','version':'1.0','requires_dist':[],'role':'runtime'}]}
            with patch.object(sys,'path',[str(root),*sys.path]):
                with self.assertRaises(ValueError):
                    api.verify_environment(lock,root,verify_records=True)
                (info/'RECORD').write_text('',encoding='utf-8')
                with self.assertRaises(ValueError):
                    api.verify_environment(lock,root,verify_records=True)
                (info/'RECORD').write_text('release_fixture.py,,\nrelease_fixture-1.0.dist-info/RECORD,,\n',encoding='utf-8')
                with self.assertRaises(ValueError):
                    api.verify_environment(lock,root,verify_records=True)

    def test_loaded_module_outside_locked_target_is_rejected(self):
        api = self.api()
        self.assertTrue(hasattr(api,'verify_import_origins'),'Locked import origin gate is missing')
        import packaging
        with tempfile.TemporaryDirectory() as directory:
            with self.assertRaises(ValueError):
                api.verify_import_origins(Path(directory))

    def test_hidden_build_child_preserves_diagnostics_and_failure_exit(self):
        from scripts import build_pyinstaller_safely as builder
        self.assertTrue(hasattr(builder,'run_locked_child'),'Hidden build child output adapter is missing')
        code = ("import sys; from scripts.build_pyinstaller_safely import run_locked_child; "
                "raise SystemExit(run_locked_child([sys.executable,'-B','-c',\"print('LOCKED_CHILD_DIAGNOSTIC',flush=True); raise SystemExit(7)\"]))")
        result = subprocess.run([sys.executable,'-B','-c',code],cwd=ROOT,capture_output=True,
            creationflags=subprocess.CREATE_NO_WINDOW)
        self.assertEqual(result.returncode,7)
        self.assertIn(b'LOCKED_CHILD_DIAGNOSTIC',result.stdout)

    def test_locked_child_resolves_pythonwin_from_locked_target(self):
        from scripts import build_pyinstaller_safely as builder
        self.assertTrue(hasattr(builder,'locked_python_paths'),'Complete locked Python path is missing')
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            target = root/'dependencies'
            package = target/'Pythonwin/pywin'
            package.mkdir(parents=True)
            (package/'__init__.py').write_text('locked_value = 73\n',encoding='utf-8')
            import os
            env = dict(os.environ,PYTHONPATH=os.pathsep.join(builder.locked_python_paths(root/'bootstrap',target,ROOT)))
            result = subprocess.run([sys.executable,'-B','-c',"import pywin; print(pywin.locked_value)"],cwd=ROOT,env=env,capture_output=True,creationflags=subprocess.CREATE_NO_WINDOW)
            self.assertEqual(result.returncode,0)
            self.assertEqual(result.stdout.strip(),b'73')

    def test_variant_analysis_uses_its_unique_actual_counter_output(self):
        from scripts import build_pyinstaller_safely as builder
        self.assertTrue(hasattr(builder,'resolve_analysis_toc'),'Variant Analysis resolver is missing')
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            first, second = root/'normal',root/'no-update'
            first.mkdir()
            second.mkdir()
            (first/'Analysis-00.toc').write_text('()',encoding='utf-8')
            (second/'Analysis-01.toc').write_text('()',encoding='utf-8')
            self.assertEqual(builder.resolve_analysis_toc(second),second/'Analysis-01.toc')
            (second/'Analysis-02.toc').write_text('()',encoding='utf-8')
            with self.assertRaises(ValueError):
                builder.resolve_analysis_toc(second)

    def test_stdlib_comparison_rejects_boolean_replaced_by_integer(self):
        api = self.api()
        for source,constants in [('value = True',(1,None)),('value = (True,)',((1,),None))]:
            with self.subTest(source=source):
                original = compile(source,'sample.py','exec')
                self.assertNotEqual(api._code_structure(original),
                                    api._code_structure(original.replace(co_consts=constants)))

    def test_stdlib_comparison_rejects_changed_float_sign_bit(self):
        api = self.api()
        original = compile('value = 0.0','sample.py','exec')
        self.assertNotEqual(api._code_structure(original),
                            api._code_structure(original.replace(co_consts=(-0.0,None))))

    def test_unowned_external_analysis_input_is_rejected(self):
        api = self.api()
        self.assertTrue(hasattr(api,'classify_bundle_inputs'),'Actual input ownership gate is missing')
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            unknown = root/'external/unlocked.dll'
            unknown.parent.mkdir()
            unknown.write_bytes(b'unlocked')
            with self.assertRaises(ValueError):
                api.classify_bundle_inputs([('unlocked.dll',unknown,'BINARY')],{}, {}, root/'app')

    def test_generated_stdlib_archive_requires_matching_stdlib_bytecode(self):
        api = self.api()
        self.assertTrue(hasattr(api,'generated_stdlib_archive'),'Generated stdlib archive validation is missing')
        with tempfile.TemporaryDirectory() as directory:
            work = Path(directory)
            archive = work/'normal/app/base_library.zip'
            archive.parent.mkdir(parents=True)
            source = Path(sys.base_prefix)/'Lib/abc.py'
            code = compile(source.read_bytes(),str(source),'exec',dont_inherit=True,optimize=0)
            with zipfile.ZipFile(archive,'w') as output:
                output.writestr('abc.pyc',importlib.util.MAGIC_NUMBER+b'\0'*12+marshal.dumps(code))
            evidence = api.generated_stdlib_archive(archive,work)
            self.assertEqual(evidence['verified_members'],1)
            # Same source/code, different sharing of the filename string in the
            # marshalled object graph must not reject a genuine stdlib module.
            equivalent = code.replace(co_filename=code.co_filename.encode().decode())
            with zipfile.ZipFile(archive,'w') as output:
                output.writestr('abc.pyc',importlib.util.MAGIC_NUMBER+b'\0'*12+marshal.dumps(equivalent))
            try:
                equivalent_evidence = api.generated_stdlib_archive(archive,work)
            except ValueError:
                self.fail('Equivalent stdlib code is rejected because marshal reference sharing differs')
            self.assertEqual(equivalent_evidence['verified_members'],1)
            with zipfile.ZipFile(archive,'w') as output:
                output.writestr('abc.pyc',importlib.util.MAGIC_NUMBER+b'\0'*12+marshal.dumps(compile('bad = True',str(source),'exec')))
            with self.assertRaises(ValueError):
                api.generated_stdlib_archive(archive,work)

    def test_signing_success_and_failure_update_all_records_and_integrity(self):
        api = self.api()
        self.assertTrue(hasattr(api,'record_signing_result'),'Consistent signing evidence update is missing')
        with tempfile.TemporaryDirectory() as directory:
            dist = Path(directory)
            bundle = dist/'app'
            (bundle/'release').mkdir(parents=True)
            exe = bundle/'app.exe'
            exe.write_bytes(b'unsigned')
            provenance = bundle/'release/build-provenance.json'
            provenance.write_text(json.dumps({'signing':{'status':'unsigned'}}),encoding='utf-8')
            (dist/'build.json').write_text(json.dumps({'signing':{'status':'unsigned'},'executables':[str(exe)]}),encoding='utf-8')
            api.write_hash_manifest(bundle)
            exe.write_bytes(b'signed')
            api.record_signing_result(exe,api.signing_plan(certificate_env='PUBLIC_CERT_REF'),True)
            self.assertTrue(api.verify_hash_manifest(bundle))
            self.assertEqual(json.loads(provenance.read_text())['signing']['status'],'signed')
            build = json.loads((dist/'build.json').read_text())
            self.assertEqual(build['signing']['status'],'signed')
            exe.write_bytes(b'failed signing changed bytes')
            api.record_signing_result(exe,api.signing_plan(certificate_env='PUBLIC_CERT_REF'),False)
            self.assertTrue(api.verify_hash_manifest(bundle))
            self.assertEqual(json.loads(provenance.read_text())['signing']['status'],'failed')
            self.assertEqual(json.loads((dist/'build.json').read_text())['signing']['status'],'failed')

    def test_signing_cli_preflights_all_evidence_before_signer_and_records_real_outcome(self):
        api = self.api()
        from scripts import sign_release
        with tempfile.TemporaryDirectory() as directory:
            dist = Path(directory)
            bundle = dist/'app'
            (bundle/'release').mkdir(parents=True)
            exe = bundle/'app.exe'
            exe.write_bytes(b'unsigned')
            provenance = bundle/'release/build-provenance.json'
            provenance.write_text(json.dumps({'signing':{'status':'unsigned'}}),encoding='utf-8')
            (dist/'build.json').write_text(json.dumps({'signing':{'status':'unsigned'},'executables':[str(exe)]}),encoding='utf-8')
            api.write_hash_manifest(bundle)
            marker = dist/'signer-invoked'
            def signer(*_args,**_kwargs):
                marker.write_text('invoked',encoding='utf-8')
                exe.write_bytes(b'signed boundary output')
                return subprocess.CompletedProcess([],0)
            real_check = sign_release.check_unlinked_path
            def reject_provenance(path):
                if Path(path).resolve() == provenance.resolve():
                    raise ValueError('Synthetic linked evidence target')
                return real_check(path)
            argv = [str(exe),'--certificate-store-sha1','a'*40,'--execute']
            with patch.object(sign_release,'check_unlinked_path',side_effect=reject_provenance), patch.object(sign_release.subprocess,'run',side_effect=signer):
                with self.assertRaises(ValueError):
                    sign_release.main(argv)
            self.assertFalse(marker.exists())
            with patch.object(sign_release.subprocess,'run',side_effect=signer):
                self.assertEqual(sign_release.main(argv),0)
            self.assertTrue(marker.exists())
            self.assertTrue(api.verify_hash_manifest(bundle))
            self.assertEqual(json.loads(provenance.read_text())['signing']['status'],'signed')
            def failed_signer(*_args,**_kwargs):
                exe.write_bytes(b'failed boundary output')
                return subprocess.CompletedProcess([],3)
            with patch.object(sign_release.subprocess,'run',side_effect=failed_signer):
                self.assertEqual(sign_release.main(argv),1)
            self.assertTrue(api.verify_hash_manifest(bundle))
            self.assertEqual(json.loads(provenance.read_text())['signing']['status'],'failed')
