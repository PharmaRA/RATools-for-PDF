"""IO owns the busy state until the worker thread has actually finished."""

import threading
import time
import unittest
from types import MethodType, SimpleNamespace
from unittest.mock import Mock, patch

from PySide6.QtCore import QObject, QThread, Signal
from PySide6.QtWidgets import QApplication, QLabel, QPushButton

from ratools_pdf.controllers.detection_controller import DetectionController
from ratools_pdf.controllers.io_controller import IOController
from ratools_pdf.controllers.main_controller import MainController
from ratools_pdf.controllers.precheck_controller import PrecheckController
from ratools_pdf.services.processing_plan import build_processing_plan
from ratools_pdf.ui.main_window import MainWindow


class ControlledIO(QThread):
    progress = Signal(str, str, str)
    finished_action = Signal(str)
    error_action = Signal(str)

    def __init__(self, result='success'):
        super().__init__()
        self.release = threading.Event()
        self.result = result

    def run(self):
        if self.result == 'success':
            self.finished_action.emit('synthetic success')
        else:
            self.error_action.emit('synthetic failure')
        self.release.wait(5)


class ControlledRead(ControlledIO):
    result_ready = Signal(dict)
    finished_precheck = Signal(str)
    error_precheck = Signal(str)
    finished_detection = Signal(str, list)
    error_detection = Signal(str)

    def __init__(self, kind, result):
        super().__init__(result)
        self.kind = kind

    def run(self):
        if self.result == 'failure':
            getattr(self, f'error_{self.kind}').emit('synthetic failure')
        elif self.kind == 'precheck':
            self.finished_precheck.emit('synthetic success')
        else:
            self.finished_detection.emit('synthetic success', [])
        self.release.wait(5)


class ControlledProcessing(ControlledIO):
    structured_progress = Signal(object)
    finished_all = Signal(str)
    error = Signal(str)

    def run(self):
        if self.result == 'success':
            self.finished_all.emit('synthetic success')
        else:
            self.error.emit('synthetic failure')
        self.release.wait(5)


class ControllerHost(QObject):
    """A real Qt receiver without constructing the settings-backed main window."""


class IOBusyTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def setUp(self):
        self.view = SimpleNamespace(
            current_file_count=1, active_preset_key=None, GENERAL_SETTING_OPTIONS=set(), PRESET_OPTIONS={},
            all_checkboxes={}, get_selected_options=lambda: {'title_from_filename'},
            show_warning_message=Mock(), show_success_message=Mock(), show_error_message=Mock(),
            show_info_message=Mock(), style=self.app.style,
        )
        for name in ('btn_start','btn_precheck','btn_retry_failed','btn_apply_precheck',
                     'btn_process_precheck_suggested','btn_bookmark_io_wizard','btn_link_io_wizard',
                     'btn_detect_annotations','btn_detect_broken_refs','btn_add_files','add_folder_btn',
                     'btn_clear','drop_zone','btn_skip_current'):
            setattr(self.view, name, QPushButton())
        for name in ('selection_summary_label','preset_summary_label','info_label','risk_hint_label','processing_hint_label'):
            setattr(self.view, name, QLabel())
        self.view.refresh_selection_summary = lambda: MainWindow.refresh_selection_summary(self.view)
        self.view.set_task_busy = lambda kind, busy: MainWindow.set_task_busy(self.view, kind, busy)
        self.host = ControllerHost()
        self.host.__dict__.update(
            worker=None, loaded_files=['C:/synthetic/report.pdf'], process_logs='',
            update_progress=Mock(), _mark_precheck_stale=Mock(),
            folder_nodes={}, file_nodes={}, view=self.view,
        )
        self.host.precheck = PrecheckController(self.host, self.view)
        self.host.detection = DetectionController(self.host, self.view)
        self.host.io = IOController(self.host, self.view)
        self.host._is_precheck_running = self.host.precheck.is_running
        self.host._is_detection_running = self.host.detection.is_running
        self.host.ensure_idle = MethodType(MainController.ensure_idle, self.host)

    def prepare_processing(self):
        self.view.settings_dialog = SimpleNamespace(default_output_edit=SimpleNamespace(text=lambda: ''))
        self.view.get_processing_mode = lambda: 'smart'
        self.view.get_processing_mode_label = lambda: 'smart'
        self.view.get_compression_settings = lambda: None
        self.host._prompt_skip_signed_files = lambda files: files
        self.host._get_processing_worker_count = lambda: 1
        self.host._get_file_timeout_seconds = MethodType(MainController._get_file_timeout_seconds, self.host)
        self.host._refresh_processing_hint = Mock()
        self.host.processing_timer = Mock()
        self.host.record_process_log_event = Mock()
        def preview(_files, options, mode, settings, base):
            # The real asynchronous modal is covered in test_processing_plan_ui.
            plan = build_processing_plan([], options, mode, settings, base)
            MainController._retain_processing_plan(self.host, plan, 'previewed')
            return plan
        self.host._preview_processing_plan = preview
        self.host.last_failed_files = []
        for method in ('_on_processing_thread_finished','_reset_processing_ui',
                       '_build_batch_result_summary','processing_finished','processing_error'):
            setattr(self.host, method, MethodType(getattr(MainController, method), self.host))

    def spin_until(self, condition):
        deadline = time.monotonic() + 2
        while not condition() and time.monotonic() < deadline:
            self.app.processEvents()
            QThread.msleep(5)
        self.assertTrue(condition())

    def test_io_is_blocked_by_processing_precheck_and_detection(self):
        for kind in ('worker','precheck','detection'):
            with self.subTest(kind=kind):
                active = Mock()
                active.isRunning.return_value = True
                if kind == 'worker':
                    self.host.worker = active
                elif kind == 'precheck':
                    self.host.precheck.precheck_worker = active
                else:
                    self.host.detection.detection_worker = active
                with patch('ratools_pdf.controllers.io_controller.IOActionWorker') as factory:
                    self.host.io.handle_io_action('export_bookmarks', dir_path='C:/synthetic/export')
                    factory.assert_not_called()
                self.host.worker = None
                self.host.precheck.precheck_worker = None
                self.host.detection.detection_worker = None

    def test_io_blocks_other_tasks_and_queue_mutations(self):
        active = Mock()
        active.isRunning.return_value = True
        self.host.io.io_worker = active
        self.assertFalse(self.host.ensure_idle('清空队列'))
        with patch('ratools_pdf.controllers.main_controller.ProcessWorker') as process, \
                patch('ratools_pdf.controllers.precheck_controller.PreCheckWorker') as precheck, \
                patch('ratools_pdf.controllers.detection_controller.DetectionWorker') as detection:
            MainController.start_processing(self.host)
            self.host.precheck.start_precheck()
            self.host.detection.start_detection('annotations')
            MainController.clear_list(self.host)
            MainController.remove_selected_items(self.host, [])
            MainController.add_files(self.host, [])
            process.assert_not_called()
            precheck.assert_not_called()
            detection.assert_not_called()
        self.assertEqual(self.host.loaded_files, ['C:/synthetic/report.pdf'])
        self.assertEqual(self.view.show_warning_message.call_count, 7)

    def test_success_and_failure_keep_busy_until_real_thread_finishes(self):
        for result in ('success','failure'):
            with self.subTest(result=result):
                worker = ControlledIO(result)
                try:
                    with patch('ratools_pdf.controllers.io_controller.IOActionWorker', return_value=worker) as factory:
                        self.host.io.handle_io_action('export_bookmarks', dir_path='C:/synthetic/export')
                        self.spin_until(lambda: 'synthetic' in self.host.process_logs)
                        self.assertTrue(worker.isRunning())
                        self.assertIs(self.host.io.io_worker, worker)
                        self.view.refresh_selection_summary()
                        self.assertFalse(self.view.btn_start.isEnabled())
                        self.assertFalse(self.view.btn_bookmark_io_wizard.isEnabled())
                        self.assertFalse(self.view.btn_clear.isEnabled())
                        self.host.io.handle_io_action('export_bookmarks', dir_path='C:/synthetic/export')
                        self.assertEqual(factory.call_count, 1)
                        worker.release.set()
                        worker.wait(2000)
                        self.spin_until(lambda: self.host.io.io_worker is None)
                        self.assertTrue(self.view.btn_start.isEnabled())
                        self.assertTrue(self.view.btn_bookmark_io_wizard.isEnabled())
                        self.assertTrue(self.host.ensure_idle('继续'))
                finally:
                    worker.release.set()
                    # deleteLater may already have disposed the finished wrapper.
                    try:
                        worker.wait(2000)
                    except RuntimeError:
                        pass
                    self.host.process_logs = ''

    def test_start_failure_releases_handle_and_restores_controls(self):
        worker = Mock()
        worker.start.side_effect = RuntimeError('cannot start')
        with patch('ratools_pdf.controllers.io_controller.IOActionWorker', return_value=worker):
            self.host.io.handle_io_action('export_bookmarks', dir_path='C:/synthetic/export')
        self.assertIsNone(self.host.io.io_worker)
        self.assertTrue(self.view.btn_bookmark_io_wizard.isEnabled())
        self.view.show_error_message.assert_called_once()

    def test_wizard_cannot_open_while_an_io_thread_is_owned(self):
        self.host.io.io_worker = Mock()
        with patch('ratools_pdf.controllers.io_controller.IODataWizardDialog') as dialog:
            self.host.io.handle_io_wizard('bookmarks')
            dialog.assert_not_called()

    def test_read_workers_block_io_until_actual_thread_completion(self):
        for kind, factory_path in (
            ('precheck', 'ratools_pdf.controllers.precheck_controller.PreCheckWorker'),
            ('detection', 'ratools_pdf.controllers.detection_controller.DetectionWorker'),
        ):
            for result in ('success', 'failure'):
                with self.subTest(kind=kind, result=result):
                    controller = getattr(self.host, kind)
                    worker = ControlledRead(kind, result)
                    try:
                        with patch(factory_path, return_value=worker):
                            if kind == 'precheck':
                                controller.start_precheck()
                            else:
                                controller.start_detection('annotations')
                        self.spin_until(lambda: 'synthetic' in self.host.process_logs)
                        self.assertTrue(controller.is_running())
                        self.assertFalse(self.view.btn_bookmark_io_wizard.isEnabled())
                        with patch('ratools_pdf.controllers.io_controller.IOActionWorker') as factory:
                            self.host.io.handle_io_action('export_bookmarks', dir_path='C:/synthetic/export')
                            factory.assert_not_called()
                        worker.release.set()
                        worker.wait(2000)
                        self.spin_until(lambda: not controller.is_running())
                        self.assertTrue(self.view.btn_bookmark_io_wizard.isEnabled())
                    finally:
                        worker.release.set()
                        try:
                            worker.wait(2000)
                        except RuntimeError:
                            pass
                        self.host.process_logs = ''

    def test_precheck_start_failure_releases_busy_ownership(self):
        self.assert_read_start_failure('precheck', 'ratools_pdf.controllers.precheck_controller.PreCheckWorker')

    def test_detection_start_failure_releases_busy_ownership(self):
        self.assert_read_start_failure('detection', 'ratools_pdf.controllers.detection_controller.DetectionWorker')

    def assert_read_start_failure(self, kind, factory_path):
        worker = Mock()
        worker.start.side_effect = RuntimeError('cannot start')
        controller = getattr(self.host, kind)
        with patch(factory_path, return_value=worker):
            if kind == 'precheck':
                controller.start_precheck()
            else:
                controller.start_detection('annotations')
        self.assertFalse(controller.is_running())
        self.assertTrue(self.view.btn_bookmark_io_wizard.isEnabled())

    def test_directory_picker_reentry_is_checked_before_launch(self):
        def choose(*_):
            self.host.worker = Mock()
            self.host.worker.isRunning.return_value = True
            return 'C:/synthetic/export'
        with patch('ratools_pdf.controllers.io_controller.QFileDialog.getExistingDirectory', side_effect=choose), \
                patch('ratools_pdf.controllers.io_controller.IOActionWorker') as factory:
            self.host.io.handle_io_action('export_bookmarks')
            factory.assert_not_called()

    def test_import_directory_failure_reports_error_and_stays_idle(self):
        with patch('ratools_pdf.controllers.io_controller.Path.mkdir', side_effect=PermissionError('denied')):
            self.host.io.handle_io_action('import_bookmarks', dir_path='C:/synthetic/import')
        self.assertFalse(self.host.io.is_running())
        self.assertTrue(self.view.btn_bookmark_io_wizard.isEnabled())
        self.view.show_error_message.assert_called_once()

    def test_processing_directory_picker_reentry_is_checked_before_launch(self):
        self.prepare_processing()
        owned = Mock()
        def choose(*_):
            self.host.io.io_worker = owned
            return 'C:/synthetic/out'
        with patch('ratools_pdf.controllers.main_controller.QFileDialog.getExistingDirectory', side_effect=choose), \
                patch('ratools_pdf.controllers.main_controller.ProcessWorker') as factory:
            MainController.start_processing(self.host)
            factory.assert_not_called()
        self.assertIs(self.host.io.io_worker, owned)
        self.assertIsNone(self.host.worker)

    def test_processing_success_and_failure_wait_for_actual_thread_completion(self):
        self.prepare_processing()
        for result in ('success', 'failure'):
            with self.subTest(result=result):
                worker = ControlledProcessing(result)
                try:
                    with patch('ratools_pdf.controllers.main_controller.QFileDialog.getExistingDirectory', return_value='C:/synthetic/out'), \
                            patch('ratools_pdf.controllers.main_controller.ProcessWorker', return_value=worker):
                        MainController.start_processing(self.host)
                    self.spin_until(lambda: 'synthetic' in self.host.process_logs)
                    self.assertIs(self.host.worker, worker)
                    self.assertTrue(worker.isRunning())
                    self.assertFalse(self.host.ensure_idle('导出'))
                    self.view.refresh_selection_summary()
                    self.assertFalse(self.view.btn_start.isEnabled())
                    self.assertFalse(self.view.btn_bookmark_io_wizard.isEnabled())
                    worker.release.set()
                    worker.wait(2000)
                    self.spin_until(lambda: self.host.worker is None)
                    self.assertTrue(self.view.btn_start.isEnabled())
                    self.assertTrue(self.view.btn_bookmark_io_wizard.isEnabled())
                finally:
                    worker.release.set()
                    try:
                        worker.wait(2000)
                    except RuntimeError:
                        pass
                    self.host.process_logs = ''

    def test_processing_start_failure_restores_busy_state(self):
        self.prepare_processing()
        worker = Mock()
        worker.start.side_effect = RuntimeError('cannot start')
        with patch('ratools_pdf.controllers.main_controller.QFileDialog.getExistingDirectory', return_value='C:/synthetic/out'), \
                patch('ratools_pdf.controllers.main_controller.ProcessWorker', return_value=worker):
            MainController.start_processing(self.host)
        self.assertIsNone(self.host.worker)
        self.assertTrue(self.view.btn_bookmark_io_wizard.isEnabled())
        self.view.show_error_message.assert_called_once()


if __name__ == '__main__':
    unittest.main()
