import hashlib
import json
from pathlib import Path
import stat
import tempfile
import unittest
from unittest.mock import patch
import zipfile

from scripts import package_nuitka_release as package


class NuitkaPackagingTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.root = Path(self.temporary.name)

    def write(self, path, data):
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(data)
        return path

    def runtime(self):
        source = self.root / 'source' / 'pymupdf'
        self.write(source / '__init__.py', b'# runtime\n')
        self.write(source / '_mupdf.pyd', b'native runtime')
        self.write(source / 'layout' / '__init__.py', b'# optional layout\n')
        self.write(source / 'layout/resources/onnx/layout.onnx', b'optional model' * 200)
        self.write(source / 'layout/resources/onnx/config.json', b'{}')
        self.write(source / '__pycache__/__init__.cpython-311.pyc', b'bytecode')
        self.write(source / 'other.onnx', b'unrelated model')
        return source

    def test_stage_preserves_code_native_libraries_and_non_model_resources(self):
        source = self.runtime()
        build = self.root / 'build'
        destination = build / 'pymupdf'
        result = package.stage_pymupdf(source, destination, build)
        self.assertEqual('passed', result['status'])
        for member in ('__init__.py', '_mupdf.pyd', 'layout/__init__.py',
                       'layout/resources/onnx/config.json', 'other.onnx'):
            self.assertEqual((source / member).read_bytes(), (destination / member).read_bytes())
        self.assertFalse((destination / 'layout/resources/onnx/layout.onnx').exists())
        self.assertFalse((destination / '__pycache__').exists())

    def test_source_link_failure_happens_before_existing_runtime_is_recycled(self):
        source = self.runtime()
        build = self.root / 'build'
        destination = build / 'pymupdf'
        sentinel = self.write(destination / 'keep.txt', b'old runtime')
        original_check = package.check_unlinked_path

        def reject_native(path):
            if Path(path) == source / '_mupdf.pyd':
                raise ValueError('simulated linked input')
            return original_check(path)

        with patch.object(package, 'check_unlinked_path', side_effect=reject_native):
            with patch.object(package, 'recycle_owned_path') as recycle:
                with self.assertRaises(ValueError):
                    package.stage_pymupdf(source, destination, build)
                recycle.assert_not_called()
        self.assertEqual(b'old runtime', sentinel.read_bytes())

    def test_stage_rejects_destination_outside_owned_build(self):
        source = self.runtime()
        with self.assertRaises(ValueError):
            package.stage_pymupdf(source, self.root / 'outside/pymupdf', self.root / 'build')
        self.assertFalse((self.root / 'outside').exists())

    def test_zip_and_repack_preserve_all_retained_file_bytes(self):
        source = self.root / 'RATools-for-PDF_v0.8.1'
        native = b'MZ' + bytes(range(256)) * 500
        self.write(source / 'RATools-for-PDF.exe', native)
        self.write(source / 'pymupdf/__init__.py', b'# runtime\n')
        self.write(source / 'pymupdf/layout/resources/onnx/layout.onnx', b'model' * 1000)
        self.write(source / 'plugins/other.onnx', b'keep this model')
        original = self.root / 'original.zip'
        with zipfile.ZipFile(original, 'w') as archive:
            for path in sorted(source.rglob('*')):
                archive.write(path, source.name + '/' + path.relative_to(source).as_posix())
        source_hash = hashlib.sha256(original.read_bytes()).hexdigest()
        folder_zip, repacked_zip = self.root / 'folder.zip', self.root / 'repacked.zip'
        folder_result = package.create_zip(source, folder_zip)
        repack_result = package.repack_zip(original, repacked_zip)
        self.assertEqual('deflate-9', folder_result['compression'])
        self.assertEqual(source_hash, hashlib.sha256(original.read_bytes()).hexdigest())
        for path in (folder_zip, repacked_zip):
            with zipfile.ZipFile(path) as archive:
                self.assertEqual(native, archive.read(source.name + '/RATools-for-PDF.exe'))
                self.assertEqual(b'keep this model', archive.read(source.name + '/plugins/other.onnx'))
                self.assertNotIn(source.name + '/pymupdf/layout/resources/onnx/layout.onnx', archive.namelist())
                self.assertIsNone(archive.testzip())
                self.assertTrue(all(info.compress_type == zipfile.ZIP_DEFLATED
                                    for info in archive.infolist() if not info.is_dir()))
        self.assertEqual(3, repack_result['retained_payloads_verified'])

    def test_existing_zip_is_preserved_and_output_inside_source_is_rejected(self):
        source = self.runtime()
        output = self.write(self.root / 'existing.zip', b'keep existing output')
        with self.assertRaises(FileExistsError):
            package.create_zip(source, output)
        self.assertEqual(b'keep existing output', output.read_bytes())
        with self.assertRaises(ValueError):
            package.create_zip(source, source / 'nested.zip')
        self.assertFalse((source / 'nested.zip').exists())

    def test_unsafe_and_linked_archive_members_are_rejected_before_output(self):
        for name, mode in (('../escape.dll', 0), ('bundle/link.dll', stat.S_IFLNK | 0o777)):
            with self.subTest(name=name):
                archive_path = self.root / ('unsafe-' + str(mode) + '.zip')
                info = zipfile.ZipInfo(name)
                info.create_system = 3
                info.external_attr = mode << 16
                with zipfile.ZipFile(archive_path, 'w') as archive:
                    archive.writestr(info, b'unsafe')
                output = self.root / ('output-' + str(mode) + '.zip')
                with self.assertRaises(ValueError):
                    package.repack_zip(archive_path, output)
                self.assertFalse(output.exists())

    def test_duplicate_windows_member_names_are_rejected(self):
        source = self.root / 'duplicates.zip'
        with zipfile.ZipFile(source, 'w') as archive:
            archive.writestr('bundle/core.dll', b'first')
            archive.writestr('bundle/Core.dll', b'second')
        with self.assertRaisesRegex(ValueError, 'Duplicate ZIP member path'):
            package.repack_zip(source, self.root / 'output.zip')

    def test_recycling_refuses_owned_root_and_outside_paths(self):
        owner = self.root / 'build'
        with patch.object(package, 'recycle_owned_path') as recycle:
            for target in (owner, self.root / 'outside'):
                with self.assertRaises(ValueError):
                    package.recycle_target(target, owner)
            recycle.assert_not_called()

    def test_workflow_and_runtime_version_constraint_use_packaging_contract(self):
        root = Path(__file__).resolve().parents[1]
        batch = (root / 'build_nuitka.bat').read_text(encoding='utf-8')
        workflow = (root / '.github/workflows/build.yml').read_text(encoding='utf-8')
        constraints = (root / 'requirements-nuitka.constraints.txt').read_text(encoding='utf-8')
        lock = json.loads((root / 'requirements-release-win311.lock.json').read_text(encoding='utf-8'))
        version = next(item['version'] for item in lock['packages'] if item['name'] == 'PyMuPDF')
        self.assertIn('PyMuPDF==' + version, constraints.splitlines())
        self.assertIn('-c requirements-nuitka.constraints.txt', workflow)
        self.assertIn('package_nuitka_release.py zip', workflow)
        self.assertIn('scripts/run_tests_safely.py', workflow)
        self.assertIn('stage-pymupdf', batch)
        self.assertNotRegex(batch, r'(?im)^\s*(?:if [^\r\n]*? )?(?:rmdir|del)\s')


if __name__ == '__main__':
    unittest.main()
