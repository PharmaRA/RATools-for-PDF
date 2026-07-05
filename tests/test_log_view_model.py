import unittest

from ratools_pdf.ui.log_view_model import (
    build_log_summary_items,
    filter_log_summary_items,
    format_duration,
)


class LogViewModelTests(unittest.TestCase):
    def test_build_log_summary_items_prefers_structured_rows_and_preserves_raw_detail(self):
        raw_text = (
            "[10:00:00] 开始处理: C:/docs/a.pdf\n"
            "    输出文件: C:/out/a.pdf\n"
            "[10:00:05] C:/docs/a.pdf\n"
            "    状态: 处理完成\n"
            "    输出文件: C:/out/a.pdf\n"
            "    结果: ✅ 处理成功；修改项：书签、链接\n"
        )
        rows = [{
            "time": "10:00:05",
            "file_original": "C:/docs/a.pdf",
            "file_output": "C:/out/a.pdf",
            "status": "处理完成",
            "success": "true",
            "duration_sec": 5,
            "changes": "书签、链接",
        }]

        items = build_log_summary_items(raw_text, rows)

        self.assertEqual(len(items), 1)
        self.assertEqual(items[0]["file_name"], "a.pdf")
        self.assertEqual(items[0]["duration_text"], "5s")
        self.assertIn("success", items[0]["tags"])
        self.assertIn("书签、链接", items[0]["search_text"])
        self.assertIn("开始处理: C:/docs/a.pdf", items[0]["detail"])

    def test_build_log_summary_items_can_fall_back_to_rendered_log_blocks(self):
        raw_text = (
            "[10:00:00] 开始处理: C:/docs/b.pdf\n"
            "    输出文件: C:/out/b.pdf\n"
            "[10:00:08] C:/docs/b.pdf\n"
            "    状态: 处理失败\n"
            "    结果: qpdf failed\n"
        )

        items = build_log_summary_items(raw_text, [])

        self.assertEqual(len(items), 1)
        self.assertEqual(items[0]["file_name"], "b.pdf")
        self.assertEqual(items[0]["status"], "处理失败")
        self.assertIn("failure", items[0]["tags"])
        self.assertIn("qpdf failed", items[0]["result"])

    def test_filter_log_summary_items_combines_status_tags_and_search_text(self):
        items = [
            {"tags": {"success"}, "search_text": "C:/docs/a.pdf 处理完成"},
            {"tags": {"failure"}, "search_text": "C:/docs/b.pdf 处理失败 qpdf"},
            {"tags": {"skip"}, "search_text": "C:/docs/c.pdf 已停止"},
        ]

        filtered = filter_log_summary_items(items, "failure", "qpdf")

        self.assertEqual(filtered, [items[1]])

    def test_format_duration_handles_blank_seconds_and_minute_rollup(self):
        self.assertEqual(format_duration(""), "")
        self.assertEqual(format_duration(None), "")
        self.assertEqual(format_duration(65), "1m 5s")


if __name__ == "__main__":
    unittest.main()
