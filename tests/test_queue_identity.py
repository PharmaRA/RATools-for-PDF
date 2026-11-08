"""Queue entries are unique and running jobs have independent identities."""

import os
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, patch

from PySide6.QtWidgets import QApplication, QTreeWidget

from ratools_pdf.controllers.main_controller import MainController
from ratools_pdf.controllers.workers import ProcessWorker


class QueueIdentityTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.root = Path(self.directory.name)
        self.source = self.root/'report.pdf'
        self.source.write_bytes(b'synthetic PDF queue entry')
        self.tree = QTreeWidget()
        self.tree.setColumnCount(3)
        self.addCleanup(self.tree.close)
        self.host = SimpleNamespace(
            loaded_files=[], folder_nodes={}, file_nodes={},
            _is_precheck_running=lambda:False, _is_detection_running=lambda:False,
            _mark_precheck_stale=Mock(),
            view=SimpleNamespace(tree=self.tree,show_info_message=Mock(),show_warning_message=Mock(),
                                 update_counters_ui=Mock()),
        )

    def test_file_and_parent_folder_are_added_once(self):
        MainController.add_files(self.host,[str(self.source),str(self.root)])
        self.assertEqual(self.host.loaded_files,[str(self.source)])
        self.assertEqual(len(self.host.file_nodes),1)
        self.assertEqual(self.tree.topLevelItem(0).childCount(),1)

    def test_relative_and_absolute_aliases_use_one_absolute_entry(self):
        relative = os.path.relpath(self.source)
        MainController.add_files(self.host,[relative,str(self.source)])
        self.assertEqual(self.host.loaded_files,[str(self.source)])
        self.assertEqual(list(self.host.file_nodes),[str(self.source)])

    @unittest.skipUnless(os.name=='nt','Windows case-insensitive paths')
    def test_case_alias_of_existing_file_does_not_create_new_entry(self):
        MainController.add_files(self.host,[str(self.source)])
        MainController.add_files(self.host,[str(self.source).upper()])
        self.assertEqual(self.host.loaded_files,[str(self.source)])
        self.assertEqual(len(self.host.file_nodes),1)

    def test_worker_keeps_a_snapshot_when_source_queue_is_extended(self):
        files = [str(self.source)]
        worker = ProcessWorker(files,set(),str(self.root/'out'))
        files.append(str(self.root/'later.pdf'))
        self.assertEqual(worker.files,[str(self.source)])

    def test_same_filename_in_different_directories_remains_two_entries(self):
        second = self.root/'other'/'report.pdf'
        second.parent.mkdir()
        second.write_bytes(b'second document')
        MainController.add_files(self.host,[str(self.source),str(second)])
        self.assertEqual(self.host.loaded_files,[str(self.source),str(second)])
        self.assertEqual(len(self.host.file_nodes),2)

    def make_task(self,index):
        stage = self.root/f'stage{index}'
        stage.mkdir()
        work = stage/'result.pdf'
        work.write_bytes(f'result {index}'.encode())
        proc = Mock()
        proc.is_alive.return_value = False
        conn = Mock()
        conn.poll.return_value = True
        conn.recv.return_value = (True,'done')
        return {'index':index,'file_path':str(self.source),'base_name':self.source.name,
                'out_path':str(self.root/f'published{index}.pdf'),'work_path':str(work),
                'staging_dir':str(stage),'proc':proc,'parent_conn':conn}

    def test_parallel_registry_does_not_replace_an_earlier_task_with_same_path(self):
        # The public planner rejects duplicate outputs. This also guards the lower
        # scheduler so future callers cannot silently orphan an already-started job.
        worker = ProcessWorker([str(self.source)]*2,set(),str(self.root),max_workers=2)
        tasks = [self.make_task(0),self.make_task(1)]
        events,errors = [],[]
        worker.structured_progress.connect(events.append)
        worker.error.connect(errors.append)
        with patch.object(worker,'_start_process_task',side_effect=tasks):
            worker._run_parallel()
        self.assertEqual(errors,[])
        self.assertEqual(len(events),2)
        for index,task in enumerate(tasks):
            self.assertEqual(Path(task['out_path']).read_bytes(),f'result {index}'.encode())
            task['proc'].join.assert_called()
            task['parent_conn'].close.assert_called_once()

    def test_skip_by_path_reaps_each_matching_task(self):
        worker = ProcessWorker([str(self.source)]*2,set(),str(self.root),max_workers=2)
        tasks = [self.make_task(0),self.make_task(1)]
        worker.request_skip_file(os.path.relpath(self.source))
        events = []
        worker.structured_progress.connect(events.append)
        with patch.object(worker,'_start_process_task',side_effect=tasks):
            worker._run_parallel()
        self.assertEqual([event['status'] for event in events],['已跳过','已跳过'])
        for task in tasks:
            self.assertFalse(Path(task['out_path']).exists())
            self.assertFalse(Path(task['staging_dir']).exists())
            task['proc'].join.assert_called()
            task['parent_conn'].close.assert_called_once()

    def test_stop_reaps_all_registered_jobs_even_if_paths_match(self):
        worker = ProcessWorker([str(self.source)]*2,set(),str(self.root),max_workers=2)
        tasks = [self.make_task(0),self.make_task(1)]
        def start(index,*_):
            if index==1:
                worker.request_stop()
            return tasks[index]
        events = []
        worker.structured_progress.connect(events.append)
        with patch.object(worker,'_start_process_task',side_effect=start):
            worker._run_parallel()
        self.assertEqual([event['status'] for event in events],['已停止','已停止'])
        for task in tasks:
            self.assertFalse(Path(task['staging_dir']).exists())
            self.assertFalse(Path(task['out_path']).exists())
            task['proc'].join.assert_called()
            task['parent_conn'].close.assert_called_once()


if __name__=='__main__':
    unittest.main()
