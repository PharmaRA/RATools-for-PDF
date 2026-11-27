"""Windows stay alive until their threads and PDF subprocesses have completed."""

import multiprocessing as mp
import tempfile
import threading
import time
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, patch

import fitz
from PySide6.QtCore import QCoreApplication, QEvent, QObject, QThread, Qt
from PySide6.QtTest import QTest
from PySide6.QtWidgets import QApplication, QLabel, QMainWindow
from shiboken6 import isValid

from ratools_pdf.controllers.main_controller import MainController
from ratools_pdf.controllers.workers import ProcessWorker
from ratools_pdf.ui.dialogs.overlay_dialog import OverlayUnderlayDialog
from ratools_pdf.ui.dialogs.page_assembly_dialog import PageAssemblyDialog
from ratools_pdf.ui.dialogs.pdf_inspector_dialog import PdfInspectorDialog
from ratools_pdf.ui.dialogs.security_center_dialog import SecurityCenterDialog
from ratools_pdf.ui.main_window import MainWindow


class HeldThread(QThread):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.release = threading.Event()

    def run(self):
        self.release.wait(10)


class WindowHarness(MainWindow):
    def __init__(self):
        # Exercise the real closeEvent without touching user QSettings.
        QMainWindow.__init__(self)
        self.processing_hint_label = QLabel(self)
        self.persist_all_settings = Mock()


def held_pdf_child(source, output, options, mode, settings, conn):
    settings['started'].set()
    settings['release'].wait(15)
    try:
        with fitz.open() as doc:
            doc.new_page().insert_text((72, 72), 'Safe final output')
            doc.save(output)
        conn.send((True, 'done'))
    finally:
        conn.close()


class WindowLifecycleTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def setUp(self):
        # Keep signal/close-guard reference cycles alive until explicit GUI-thread
        # destruction. Otherwise a later ProcessWorker allocation can trigger GC
        # of old widgets on that worker, stopping Qt timers from the wrong thread.
        self._owned_objects = []

    def own(self, obj):
        self._owned_objects.append(obj)
        return obj

    def tearDown(self):
        self.assertEqual(QThread.currentThread(), self.app.thread())
        for obj in self._owned_objects:
            if isValid(obj) and isinstance(obj, QThread):
                if isinstance(obj, HeldThread):
                    obj.release.set()
                elif obj.isRunning():
                    obj.request_stop()
                self.assertTrue(obj.wait(15000), 'Test-owned thread did not finish')
        # Close guards still need live controller/thread wrappers at this point.
        for obj in self._owned_objects:
            if isValid(obj) and isinstance(obj, (QMainWindow, PdfInspectorDialog,
                    OverlayUnderlayDialog, PageAssemblyDialog, SecurityCenterDialog)):
                obj.close()
        for obj in reversed(self._owned_objects):
            if isValid(obj):
                obj.deleteLater()
                QCoreApplication.sendPostedEvents(obj, QEvent.DeferredDelete)
                self.assertFalse(isValid(obj), 'Test-owned QObject survived cleanup')
        self._owned_objects.clear()

    def spin_until(self, condition):
        deadline = time.monotonic() + 5
        while not condition() and time.monotonic() < deadline:
            self.app.processEvents()
            QThread.msleep(5)
        self.assertTrue(condition())

    def controller(self, view):
        controller = MainController.__new__(MainController)
        QObject.__init__(controller)
        self.own(controller)
        controller.view = view
        controller.worker = None
        controller.precheck = SimpleNamespace(precheck_worker=None)
        controller.detection = SimpleNamespace(detection_worker=None)
        controller.io = SimpleNamespace(io_worker=None)
        controller.updates = SimpleNamespace(update_worker=None)
        view.close_guard = controller.can_close
        return controller

    def test_all_workbenches_reject_title_escape_system_and_programmatic_close(self):
        for cls in (OverlayUnderlayDialog, PageAssemblyDialog, SecurityCenterDialog, PdfInspectorDialog):
            for action in ('title', 'escape', 'system', 'accept', 'done'):
                with self.subTest(dialog=cls.__name__, action=action):
                    dialog = self.own(cls())
                    # Inspector shadows btn_close with its footer button. Find
                    # the actual title button by object name.
                    title = dialog.findChild(type(dialog.btn_close), 'dialogCloseBtn')
                    worker = self.own(HeldThread(dialog))
                    worker.start()
                    dialog.show()
                    self.spin_until(worker.isRunning)
                    try:
                        if action == 'title':
                            title.click()
                        elif action == 'escape':
                            QTest.keyClick(dialog, Qt.Key_Escape)
                        elif action == 'system':
                            dialog.close()
                        elif action == 'accept':
                            dialog.accept()
                        else:
                            dialog.done(0)
                        self.assertTrue(dialog.isVisible())
                        self.assertTrue(worker.isRunning())
                    finally:
                        worker.release.set()
                        worker.wait(2000)
                        dialog.close()
                    self.assertFalse(dialog.isVisible())

    def test_main_close_waits_for_every_controller_worker_type(self):
        for kind in ('processing','precheck','detection','io','updates'):
            with self.subTest(kind=kind):
                view = self.own(WindowHarness())
                controller = self.controller(view)
                worker = self.own(HeldThread())
                if kind == 'processing':
                    controller.worker = worker
                else:
                    attribute = {'precheck':'precheck_worker','detection':'detection_worker',
                                 'io':'io_worker','updates':'update_worker'}[kind]
                    setattr(getattr(controller, kind), attribute, worker)
                worker.start()
                view.show()
                self.spin_until(worker.isRunning)
                try:
                    view.close()
                    self.assertTrue(view.isVisible())
                    view.persist_all_settings.assert_not_called()
                finally:
                    worker.release.set()
                    worker.wait(2000)
                    view.close()
                self.assertFalse(view.isVisible())
                view.persist_all_settings.assert_called_once()

    def test_main_close_finds_dialog_threads_even_when_latest_reference_changed(self):
        view = self.own(WindowHarness())
        controller = self.controller(view)
        dialog = self.own(PdfInspectorDialog(parent=view))
        old_worker = self.own(HeldThread(dialog))
        latest_worker = self.own(HeldThread(dialog))
        dialog.worker = latest_worker
        old_worker.start()
        view.show()
        self.spin_until(old_worker.isRunning)
        try:
            view.close()
            self.assertTrue(view.isVisible())
            self.assertFalse(controller.can_close())
        finally:
            old_worker.release.set()
            old_worker.wait(2000)
            view.close()
        self.assertFalse(view.isVisible())

    def test_shutdown_waits_for_all_workers_and_retained_dialog_children(self):
        view = self.own(WindowHarness())
        controller = self.controller(view)
        workers = [self.own(HeldThread()) for _ in range(5)] + [self.own(HeldThread(view))]
        controller.worker = workers[0]
        controller.precheck.precheck_worker = workers[1]
        controller.detection.detection_worker = workers[2]
        controller.io.io_worker = workers[3]
        controller.updates.update_worker = workers[4]
        for worker in workers:
            worker.start()
        release = threading.Timer(0.1, lambda: [worker.release.set() for worker in workers])
        release.start()
        try:
            controller.shutdown()
            self.assertTrue(all(worker.wait(0) for worker in workers))
        finally:
            release.join()
            for worker in workers:
                worker.release.set()
                worker.wait(2000)
            view.close()

    def test_main_close_preserves_live_pdf_subprocess_until_output_is_published(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root/'input.pdf'
            with fitz.open() as doc:
                doc.new_page().insert_text((72, 72), 'Synthetic input')
                doc.save(source)
            started, release = mp.get_context('spawn').Event(), mp.get_context('spawn').Event()
            worker = self.own(ProcessWorker([str(source)], set(), str(root/'out'),
                                   compression_settings={'started':started, 'release':release}))
            view = self.own(WindowHarness())
            controller = self.controller(view)
            controller.worker = worker
            view.show()
            children = []
            try:
                with patch('ratools_pdf.controllers.workers._process_document_task_pipe', held_pdf_child):
                    worker.start()
                    self.assertTrue(started.wait(10))
                    children = mp.active_children()
                    self.assertTrue(children)
                    view.close()
                    self.assertTrue(view.isVisible())
                    self.assertTrue(any(child.is_alive() for child in children))
                    release.set()
                    self.assertTrue(worker.wait(10000))
                self.assertTrue(all(not child.is_alive() for child in children))
                with fitz.open(root/'out'/'input.pdf') as doc:
                    self.assertIn('Safe final output', doc[0].get_text())
                view.close()
                self.assertFalse(view.isVisible())
            finally:
                release.set()
                worker.wait(15000)
                for child in children:
                    child.join(2)
                view.close()


if __name__ == '__main__':
    unittest.main()
