import builtins
import ast
import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import fitz

from ratools_pdf.pdf.processor import process_document
from ratools_pdf.services.compression_check import verify_image_compression


class CompressionCapabilityTests(unittest.TestCase):
    def test_missing_pillow_fails_requested_compression_without_publishing(self):
        with tempfile.TemporaryDirectory() as directory:
            source, output = Path(directory)/'in.pdf', Path(directory)/'out.pdf'
            with fitz.open() as doc:
                doc.new_page()
                doc.save(source)
            output.write_bytes(b'previous output')
            import_module = builtins.__import__
            def without_pillow(name, *args, **kwargs):
                if name == 'PIL' or name.startswith('PIL.'):
                    raise ImportError('synthetic missing Pillow')
                return import_module(name,*args,**kwargs)
            with patch('builtins.__import__', side_effect=without_pillow):
                ok, message = process_document(str(source),str(output),{'compress_images'},processing_mode='force')
            self.assertFalse(ok)
            self.assertIn('Pillow',message)
            self.assertEqual(output.read_bytes(),b'previous output')

    def test_build_config_keeps_pillow(self):
        root=Path(__file__).resolve().parents[1]
        script=(root/'build_pyinstaller.bat').read_text(encoding='utf-8')
        self.assertNotIn('--exclude-module PIL',script)
        self.assertIn('build_pyinstaller_safely.py',script)
        tree=ast.parse((root/'scripts/build_pyinstaller_safely.py').read_text(encoding='utf-8'))
        excluded=[]
        for node in ast.walk(tree):
            if isinstance(node,ast.For) and isinstance(node.target,ast.Name) and node.target.id=='module':
                excluded.extend(ast.literal_eval(node.iter))
        self.assertTrue(excluded)
        self.assertNotIn('PIL',excluded)

    def test_release_check_runs_real_compression_and_refuses_reusing_directory(self):
        with tempfile.TemporaryDirectory() as directory:
            target=Path(directory)/'verification'
            self.assertEqual(verify_image_compression(target),0)
            report=json.loads((target/'verification.json').read_text(encoding='utf-8'))
            self.assertTrue(report['ok'])
            self.assertLess(report['output_bytes'],report['input_bytes'])
            before=(target/'compressed.pdf').read_bytes()
            with self.assertRaises(FileExistsError):
                verify_image_compression(target)
            self.assertEqual((target/'compressed.pdf').read_bytes(),before)
