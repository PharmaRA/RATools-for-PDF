"""Only successful current precheck results can change rules or start processing."""

import tempfile
import unittest
from pathlib import Path
from types import MethodType, SimpleNamespace
from unittest.mock import Mock

from PySide6.QtWidgets import QApplication, QCheckBox, QPushButton, QTreeWidget

from ratools_pdf.controllers.main_controller import MainController
from ratools_pdf.controllers.precheck_controller import PrecheckController


class PrecheckStalenessTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.root = Path(self.directory.name)
        self.source = self.root/'source.pdf'
        self.source.write_bytes(b'synthetic source')
        self.option = QCheckBox()
        self.view = SimpleNamespace(
            tree=QTreeWidget(), all_checkboxes={'title_from_filename':self.option},
            refresh_selection_summary=Mock(), persist_all_settings=Mock(),
            _set_preset_button_state=Mock(), get_selected_options=lambda: {'title_from_filename'},
            show_warning_message=Mock(), show_success_message=Mock(), show_info_message=Mock(),
            show_error_message=Mock(), update_counters_ui=Mock(),
        )
        self.view.tree.setColumnCount(3)
        for name in ('btn_start','btn_precheck','btn_apply_precheck','btn_process_precheck_suggested'):
            setattr(self.view, name, QPushButton())
        self.host = SimpleNamespace(
            view=self.view, loaded_files=[], file_nodes={}, folder_nodes={}, process_logs='',
            io=SimpleNamespace(is_running=lambda:False), _is_precheck_running=lambda:False,
            _is_detection_running=lambda:False, ensure_idle=Mock(return_value=True),
            start_processing=Mock(), _reset_processing_ui=Mock(), _build_batch_result_summary=lambda text:text,
        )
        self.precheck = PrecheckController(self.host, self.view)
        self.host.precheck = self.precheck
        self.host._mark_precheck_stale = self.precheck.mark_stale
        self.host._cleanup_empty_folders = MethodType(MainController._cleanup_empty_folders, self.host)
        MainController.add_files(self.host, [str(self.source)])
        self.set_current_result()

    def set_current_result(self):
        self.precheck.last_precheck_results = []
        self.precheck.last_precheck_suggested_files = []
        self.precheck._record_precheck_result(dict(file_path=str(self.source), status='建议处理',
                                                 suggestion_ids='title_from_filename'))
        self.precheck.precheck_finished('synthetic current result')

    def assert_stale_cannot_execute(self):
        self.precheck.apply_precheck_suggestions()
        self.precheck.start_precheck_suggested_processing()
        self.assertFalse(self.option.isChecked())
        self.host.start_processing.assert_not_called()
        self.view.persist_all_settings.assert_not_called()
        self.assertFalse(self.view.btn_apply_precheck.property('hasPrecheckSuggestions'))
        self.assertFalse(self.view.btn_process_precheck_suggested.property('hasPrecheckSuggestedFiles'))

    def test_adding_file_invalidates_old_rules_and_suggested_processing(self):
        added = self.root/'new.pdf'
        added.write_bytes(b'new synthetic source')
        MainController.add_files(self.host, [str(added)])
        self.assert_stale_cannot_execute()

    def test_removing_file_invalidates_old_rules_even_if_result_rows_remain(self):
        MainController.remove_selected_items(self.host, [self.host.file_nodes[str(self.source)]])
        self.assert_stale_cannot_execute()

    def test_processing_success_and_failure_invalidate_precheck(self):
        for handler in (MainController.processing_finished, MainController.processing_error):
            with self.subTest(handler=handler.__name__):
                self.set_current_result()
                handler(self.host, 'synthetic processing result')
                self.assert_stale_cannot_execute()

    def test_failed_precheck_does_not_enable_partial_suggestions(self):
        self.precheck.precheck_result_current = False
        self.precheck.precheck_error('failed after first row')
        self.assert_stale_cannot_execute()

    def test_rechecking_restores_valid_suggestions_and_processing(self):
        self.precheck.mark_stale()
        self.set_current_result()
        self.precheck.apply_precheck_suggestions()
        self.precheck.start_precheck_suggested_processing()
        self.assertTrue(self.option.isChecked())
        self.host.start_processing.assert_called_once_with(processing_files=[str(self.source)])
        self.view.persist_all_settings.assert_called_once()

    def test_current_suggestions_cannot_execute_while_another_task_is_busy(self):
        self.host.ensure_idle.return_value = False
        self.precheck.apply_precheck_suggestions()
        self.precheck.start_precheck_suggested_processing()
        self.assertFalse(self.option.isChecked())
        self.host.start_processing.assert_not_called()


if __name__ == '__main__':
    unittest.main()
