import unittest

from controller import (
    _build_io_paths_for_file,
    _collect_ectd_rename_plan,
    _normalized_ectd_name,
    _render_logs_as_csv_rows,
    _select_log_rows_for_export,
    _structured_log_row_from_event,
)


class ControllerGuardTests(unittest.TestCase):
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


if __name__ == "__main__":
    unittest.main()
