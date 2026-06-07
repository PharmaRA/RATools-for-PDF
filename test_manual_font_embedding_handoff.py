import os
import importlib.util
import sys
import tempfile
import types
import unittest
from types import SimpleNamespace
from unittest import mock

ROOT_DIR = os.path.dirname(os.path.abspath(__file__))


class _DummySignal:
    def __init__(self, *_args, **_kwargs):
        pass

    def connect(self, *_args, **_kwargs):
        pass

    def emit(self, *_args, **_kwargs):
        pass


class _DummyQObject:
    def __init__(self, *_args, **_kwargs):
        pass


class _DummyQThread(_DummyQObject):
    def start(self):
        pass

    def isRunning(self):
        return False


class _DummyQTimer:
    def __init__(self, *_args, **_kwargs):
        self.timeout = _DummySignal()

    def setInterval(self, *_args, **_kwargs):
        pass


def _load_module_with_stubs(module_name, file_name, stubs):
    old_modules = {}
    missing = object()
    for name, module in stubs.items():
        old_modules[name] = sys.modules.get(name, missing)
        sys.modules[name] = module
    try:
        spec = importlib.util.spec_from_file_location(module_name, os.path.join(ROOT_DIR, file_name))
        module = importlib.util.module_from_spec(spec)
        sys.modules[module_name] = module
        spec.loader.exec_module(module)
        return module
    finally:
        for name, old_module in old_modules.items():
            if old_module is missing:
                sys.modules.pop(name, None)
            else:
                sys.modules[name] = old_module


def _load_controller():
    pyside = types.ModuleType("PySide6")
    qt_widgets = types.ModuleType("PySide6.QtWidgets")
    qt_widgets.QFileDialog = object
    qt_widgets.QTreeWidgetItem = object
    qt_widgets.QMenu = object
    qt_core = types.ModuleType("PySide6.QtCore")
    qt_core.QObject = _DummyQObject
    qt_core.QThread = _DummyQThread
    qt_core.Signal = _DummySignal
    qt_core.Qt = SimpleNamespace()
    qt_core.QTimer = _DummyQTimer
    qt_core.QCoreApplication = SimpleNamespace(instance=lambda: None)
    qt_gui = types.ModuleType("PySide6.QtGui")
    qt_gui.QColor = object
    pdf_processor = types.ModuleType("pdf_processor")
    pdf_processor.PDFProcessor = object
    view = types.ModuleType("view")
    view.LogDialog = object
    return _load_module_with_stubs(
        "controller_manual_font_handoff_under_test",
        "controller.py",
        {
            "PySide6": pyside,
            "PySide6.QtWidgets": qt_widgets,
            "PySide6.QtCore": qt_core,
            "PySide6.QtGui": qt_gui,
            "pdf_processor": pdf_processor,
            "view": view,
        },
    )


def _load_pdf_processor():
    class FakeDoc:
        needs_pass = False
        page_count = 1
        metadata = {}

        def pdf_catalog(self):
            return 1

        def close(self):
            pass

    fake_fitz = types.ModuleType("fitz")
    fake_fitz.open = lambda _path: FakeDoc()
    fake_fitz.LINK_GOTOR = 5
    return _load_module_with_stubs(
        "pdf_processor_manual_font_handoff_under_test",
        "pdf_processor.py",
        {"fitz": fake_fitz},
    )


class ManualFontEmbeddingHandoffTests(unittest.TestCase):
    def test_extracts_acrobat_path_from_quoted_and_plain_registry_commands(self):
        MainController = _load_controller().MainController

        self.assertEqual(
            MainController._extract_executable_from_open_command(
                r'"C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe" "%1"'
            ),
            r"C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe",
        )
        self.assertEqual(
            MainController._extract_executable_from_open_command(
                r"C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe"
            ),
            r"C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe",
        )

    def test_warns_when_no_file_selected_for_manual_font_embedding(self):
        MainController = _load_controller().MainController
        warnings = []
        controller = MainController.__new__(MainController)
        controller.view = SimpleNamespace(
            tree=SimpleNamespace(selectedItems=lambda: []),
            show_warning_message=lambda title, message: warnings.append((title, message)),
        )

        controller.open_selected_files_in_acrobat_for_font_embedding()

        self.assertEqual(len(warnings), 1)
        self.assertIn("未选择", warnings[0][0])

    def test_shows_manual_font_embedding_dialog_before_opening_acrobat(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            pdf_a = os.path.join(temp_dir, "a.pdf")
            pdf_b = os.path.join(temp_dir, "b.pdf")
            txt = os.path.join(temp_dir, "note.txt")
            for path in (pdf_a, pdf_b):
                with open(path, "wb") as f:
                    f.write(b"%PDF-1.7\n%%EOF\n")
            with open(txt, "w", encoding="utf-8") as f:
                f.write("not a pdf")

            class FakeItem:
                def __init__(self, path):
                    self.path = path

                def text(self, column):
                    return ["name", self.path, "等待处理"][column]

            opened = []
            dialogs = []
            MainController = _load_controller().MainController
            controller = MainController.__new__(MainController)
            controller.view = SimpleNamespace(
                tree=SimpleNamespace(selectedItems=lambda: [
                    FakeItem(pdf_a),
                    FakeItem(txt),
                    FakeItem(pdf_b),
                    FakeItem(pdf_a),
                ]),
                show_warning_message=lambda *_args: None,
                show_manual_font_embedding_dialog=lambda paths, callback: dialogs.append((paths, callback)),
            )
            controller._open_pdf_for_manual_font_embedding = lambda path, acrobat_path=None: opened.append(path)
            controller._find_acrobat_executable = lambda: r"C:\Acrobat\Acrobat.exe"

            controller.open_selected_files_in_acrobat_for_font_embedding()

            self.assertEqual(opened, [])
            self.assertEqual(len(dialogs), 1)
            self.assertEqual(dialogs[0][0], [pdf_a, pdf_b])

            ok, message = dialogs[0][1]()

            self.assertTrue(ok, message)
            self.assertEqual(opened, [pdf_a, pdf_b])
            self.assertIn("已打开 2 个 PDF", message)

    def test_embed_nonstandard_fonts_option_is_noop_in_batch_processor(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            src = os.path.join(temp_dir, "src.pdf")
            out = os.path.join(temp_dir, "out.pdf")
            with open(src, "wb") as f:
                f.write(b"%PDF-1.7\n%%EOF\n")

            PDFProcessor = _load_pdf_processor().PDFProcessor
            with mock.patch.object(PDFProcessor, "_run_font_embedding_workflow") as workflow:
                ok, msg = PDFProcessor.process_document(src, out, {"embed_nonstandard_fonts"})

            self.assertTrue(ok, msg)
            self.assertTrue(os.path.exists(out))
            workflow.assert_not_called()


if __name__ == "__main__":
    unittest.main()
