import unittest

import os
import tempfile
from unittest.mock import patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import QApplication
from ratools_pdf.controllers.io_actions import (
    _build_io_paths_for_file,
    _build_io_preview_rows,
    _collect_ectd_rename_plan,
    _io_action_metadata,
    _normalized_ectd_name,
)
from ratools_pdf.controllers.log_export import (
    _render_logs_as_csv_rows,
    _select_log_rows_for_export,
    _structured_log_row_from_event,
)
from ratools_pdf.controllers.workers import IOActionWorker
from ratools_pdf.ui.dialogs import IODataWizardDialog


class ControllerGuardTests(unittest.TestCase):
    def test_controller_worker_package_class_is_available(self):
        self.assertEqual(IOActionWorker.__name__, "IOActionWorker")

    def test_controller_helpers_are_available_from_package_modules(self):
        self.assertTrue(callable(_build_io_paths_for_file))
        self.assertTrue(callable(_collect_ectd_rename_plan))
        self.assertTrue(callable(_normalized_ectd_name))
        self.assertTrue(callable(_render_logs_as_csv_rows))
        self.assertTrue(callable(_select_log_rows_for_export))

    def test_io_action_metadata_describes_bookmark_export(self):
        meta = _io_action_metadata("export_bookmarks")

        self.assertEqual(meta["data_kind"], "bookmarks")
        self.assertEqual(meta["data_label"], "书签")
        self.assertEqual(meta["data_type"], "CSV")
        self.assertTrue(meta["is_export"])
        self.assertEqual(meta["action_name"], "导出")

    def test_io_preview_rows_report_import_matches_and_missing_files(self):
        with tempfile.TemporaryDirectory() as tmp:
            source_dir = os.path.join(tmp, "source")
            data_dir = os.path.join(tmp, "data")
            os.makedirs(os.path.join(source_dir, "a"))
            os.makedirs(os.path.join(source_dir, "b"))
            os.makedirs(os.path.join(data_dir, "a"))

            file_a = os.path.join(source_dir, "a", "report.pdf")
            file_b = os.path.join(source_dir, "b", "report.pdf")
            open(file_a, "wb").close()
            open(file_b, "wb").close()
            open(os.path.join(data_dir, "a", "report_bookmarks.csv"), "w", encoding="utf-8").close()

            rows = _build_io_preview_rows(
                [file_a, file_b],
                "import_bookmarks",
                data_dir,
                common_base=source_dir,
            )

            self.assertEqual(rows[0]["status"], "已匹配")
            self.assertTrue(rows[0]["data_path"].endswith(os.path.join("a", "report_bookmarks.csv")))
            self.assertEqual(rows[1]["status"], "未找到")
            self.assertTrue(rows[1]["data_path"].endswith(os.path.join("b", "report_bookmarks.csv")))

    def test_io_preview_rows_support_bookmarks_and_links_together(self):
        with tempfile.TemporaryDirectory() as tmp:
            file_path = os.path.join(tmp, "report.pdf")
            open(file_path, "wb").close()

            rows = _build_io_preview_rows(
                [file_path],
                ["export_bookmarks", "export_links"],
                tmp,
                common_base=tmp,
            )

            self.assertEqual([row["data_label"] for row in rows], ["书签", "链接"])
            self.assertTrue(rows[0]["data_path"].endswith("report_bookmarks.csv"))
            self.assertTrue(rows[1]["data_path"].endswith("report_links.json"))

    def test_io_worker_exports_bookmarks_and_links_together(self):
        with tempfile.TemporaryDirectory() as tmp:
            file_path = os.path.join(tmp, "report.pdf")
            open(file_path, "wb").close()

            worker = IOActionWorker(
                ["export_bookmarks", "export_links"],
                [file_path],
                tmp,
                common_base=tmp,
            )

            with patch("ratools_pdf.controllers.workers.PDFProcessor.export_bookmarks") as export_bookmarks, \
                    patch("ratools_pdf.controllers.workers.PDFProcessor.export_links") as export_links:
                worker.run()

            export_bookmarks.assert_called_once_with(file_path, os.path.join(tmp, "report_bookmarks.csv"))
            export_links.assert_called_once_with(file_path, os.path.join(tmp, "report_links.json"), scope="all")

    def test_io_worker_imports_bookmarks_and_links_as_pipeline(self):
        with tempfile.TemporaryDirectory() as tmp:
            source_dir = os.path.join(tmp, "source")
            data_dir = os.path.join(tmp, "data")
            output_dir = os.path.join(tmp, "out")
            os.makedirs(source_dir)
            os.makedirs(data_dir)
            file_path = os.path.join(source_dir, "report.pdf")
            csv_path = os.path.join(data_dir, "report_bookmarks.csv")
            json_path = os.path.join(data_dir, "report_links.json")
            output_pdf = os.path.join(output_dir, "report.pdf")
            open(file_path, "wb").close()
            open(csv_path, "w", encoding="utf-8").close()
            open(json_path, "w", encoding="utf-8").close()

            worker = IOActionWorker(
                ["import_bookmarks", "import_links"],
                [file_path],
                data_dir,
                output_dir=output_dir,
                common_base=source_dir,
            )

            with patch("ratools_pdf.controllers.workers.PDFProcessor.import_bookmarks") as import_bookmarks, \
                    patch("ratools_pdf.controllers.workers.PDFProcessor.import_links") as import_links:
                worker.run()

            import_bookmarks.assert_called_once()
            import_links.assert_called_once()
            self.assertEqual(import_bookmarks.call_args.args[0], file_path)
            self.assertEqual(import_bookmarks.call_args.args[1], csv_path)
            self.assertNotEqual(import_bookmarks.call_args.args[2], output_pdf)
            self.assertEqual(import_links.call_args.args[0], import_bookmarks.call_args.args[2])
            self.assertEqual(import_links.call_args.args[1], json_path)
            self.assertEqual(import_links.call_args.args[2], output_pdf)

    def test_io_paths_do_not_escape_target_dir_when_file_is_outside_common_base(self):
        data_path, output_path = _build_io_paths_for_file(
            r"C:\other\report.pdf",
            "links",
            r"D:\target",
            output_dir=r"E:\out",
            common_base=r"C:\base",
        )

        self.assertNotIn("..", data_path)
        self.assertNotIn("..", output_path)
        self.assertIn("_external", data_path)
        self.assertIn("_external", output_path)

    def test_io_paths_keep_relative_structure_for_files_inside_common_base(self):
        data_path, output_path = _build_io_paths_for_file(
            r"C:\base\a\report.pdf",
            "bookmarks",
            r"D:\target",
            output_dir=r"E:\out",
            common_base=r"C:\base",
        )

        self.assertIn("a", data_path)
        self.assertIn("a", output_path)
        self.assertNotIn("_external", data_path)
        self.assertNotIn("_external", output_path)

    def test_normalized_ectd_names_detect_collisions(self):
        self.assertEqual(_normalized_ectd_name("A B.pdf", 1), _normalized_ectd_name("a-b.pdf", 2))

    def test_collect_ectd_rename_plan_reports_collisions(self):
        rename_pairs, collisions = _collect_ectd_rename_plan([
            r"C:\docs\A B.pdf",
            r"C:\docs\a-b.pdf",
        ])

        self.assertTrue(rename_pairs)
        self.assertIn("a-b.pdf", collisions)
        self.assertEqual(len(collisions["a-b.pdf"]), 2)

    def test_render_logs_as_csv_rows_parses_known_log_format(self):
        rows = _render_logs_as_csv_rows(
            "[10:00:00] 开始处理: C:/a.pdf\n"
            "    输出文件: C:/out/a.pdf\n"
            "[10:00:05] C:/a.pdf\n"
            "    状态: 处理完成\n"
            "    结果: ✅ 处理成功；修改项：Foo、Bar\n"
        )

        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["file_original"], "C:/a.pdf")
        self.assertEqual(rows[0]["file_output"], "C:/out/a.pdf")
        self.assertEqual(rows[0]["status"], "处理完成")
        self.assertEqual(rows[0]["changes"], "Foo、Bar")

    def test_select_log_rows_for_export_prefers_structured_rows(self):
        structured_rows = [{
            "time": "10:00:05",
            "file_original": "C:/structured.pdf",
            "file_output": "C:/out/structured.pdf",
            "status": "处理完成",
            "success": "true",
            "duration_sec": 5,
            "changes": "Structured",
        }]
        text_log = (
            "[10:00:00] 开始处理: C:/parsed.pdf\n"
            "    输出文件: C:/out/parsed.pdf\n"
            "[10:00:05] C:/parsed.pdf\n"
            "    状态: 处理完成\n"
            "    结果: ✅ 处理成功；修改项：Parsed\n"
        )

        rows = _select_log_rows_for_export(structured_rows, text_log)

        self.assertEqual(rows, structured_rows)

    def test_select_log_rows_for_export_falls_back_to_text_logs(self):
        rows = _select_log_rows_for_export(
            [],
            "[10:00:00] 开始处理: C:/parsed.pdf\n"
            "    输出文件: C:/out/parsed.pdf\n"
            "[10:00:05] C:/parsed.pdf\n"
            "    状态: 处理完成\n"
            "    结果: ✅ 处理成功；修改项：Parsed\n",
        )

        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["file_original"], "C:/parsed.pdf")
        self.assertEqual(rows[0]["changes"], "Parsed")

    def test_structured_log_row_from_event_extracts_terminal_event(self):
        row = _structured_log_row_from_event(
            {
                "time": "10:00:05",
                "file_path": "C:/a.pdf",
                "out_path": "C:/out/a.pdf",
                "status": "处理完成",
                "message": "✅ 处理成功；修改项：Foo、Bar",
            },
            {"C:/a.pdf": {"time": "10:00:00"}},
        )

        self.assertEqual(row["time"], "10:00:05")
        self.assertEqual(row["file_original"], "C:/a.pdf")
        self.assertEqual(row["file_output"], "C:/out/a.pdf")
        self.assertEqual(row["status"], "处理完成")
        self.assertEqual(row["success"], "true")
        self.assertEqual(row["duration_sec"], 5)
        self.assertEqual(row["changes"], "Foo、Bar")

    def test_structured_log_row_from_event_includes_stopped_terminal_event(self):
        row = _structured_log_row_from_event(
            {
                "time": "10:00:05",
                "file_path": "C:/a.pdf",
                "out_path": "C:/out/a.pdf",
                "status": "已停止",
                "message": "用户手动停止处理",
            },
            {"C:/a.pdf": {"time": "10:00:00"}},
        )

        self.assertEqual(row["file_original"], "C:/a.pdf")
        self.assertEqual(row["status"], "已停止")
        self.assertEqual(row["success"], "false")
        self.assertEqual(row["duration_sec"], 5)

    def test_structured_log_row_from_event_handles_midnight_duration(self):
        row = _structured_log_row_from_event(
            {
                "time": "00:00:03",
                "file_path": "C:/a.pdf",
                "out_path": "C:/out/a.pdf",
                "status": "处理失败",
                "message": "failed",
            },
            {"C:/a.pdf": {"time": "23:59:58"}},
        )

        self.assertEqual(row["duration_sec"], 5)

    def test_io_wizard_empty_state_replaces_preview_table_until_directory_is_selected(self):
        app = QApplication.instance() or QApplication([])
        dialog = IODataWizardDialog("bookmarks", 1, lambda _action, _path: [])
        dialog.show()
        app.processEvents()

        self.assertEqual(dialog.data_kind, "bookmarks")
        self.assertEqual(dialog.radio_export.text(), "导出数据")
        self.assertEqual(dialog.btn_browse.text(), "选择目录")
        self.assertEqual(dialog.btn_confirm.text(), "开始导出")
        self.assertFalse(dialog.preview_table.isVisible())
        self.assertTrue(dialog.preview_empty_label.isVisible())
        self.assertIn("选择数据目录后", dialog.preview_empty_label.text())

    def test_io_wizard_import_preview_updates_summary_and_primary_button(self):
        rows = [
            {"file_name": "matched.pdf", "status": "已匹配", "data_path": "D:/data/matched_bookmarks.csv"},
            {"file_name": "missing.pdf", "status": "未找到", "data_path": "D:/data/missing_bookmarks.csv"},
        ]
        app = QApplication.instance() or QApplication([])
        dialog = IODataWizardDialog("bookmarks", 2, lambda _action, _path: rows)
        dialog.show()

        dialog.radio_import.setChecked(True)
        dialog.dir_edit.setText("D:/data")
        app.processEvents()

        self.assertEqual(dialog.btn_confirm.text(), "开始导入")
        self.assertTrue(dialog.preview_table.isVisible())
        self.assertFalse(dialog.preview_empty_label.isVisible())
        self.assertIn("已匹配 1 / 2", dialog.summary_label.text())
        self.assertTrue(dialog.btn_confirm.isEnabled())

    def test_io_wizard_links_dialog_only_handles_links(self):
        app = QApplication.instance() or QApplication([])
        dialog = IODataWizardDialog("links", 1, lambda _actions, _path: [
            {"file_name": "report.pdf", "data_label": "链接", "status": "将生成", "data_path": "D:/data/report_links.json"},
        ])
        dialog.show()

        dialog.dir_edit.setText("D:/data")
        app.processEvents()

        self.assertEqual(dialog.data_kind, "links")
        self.assertEqual(dialog.data_type, "JSON")
        self.assertEqual(dialog.get_action_types(), ["export_links"])
        self.assertEqual(dialog.preview_table.rowCount(), 1)
        self.assertIn("1 个数据文件", dialog.summary_label.text())


if __name__ == "__main__":
    unittest.main()
