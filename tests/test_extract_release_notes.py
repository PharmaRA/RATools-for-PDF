import unittest

from scripts.extract_release_notes import extract_release_notes


CHANGELOG = """# 更新日志

本文件记录值得关注的版本变更。

## [0.7.1] - 2026-07-13

### 新增

- 处理前检测已签名文件

## [0.7.0] - 2026-07-13

### 变更

- 应用版本号提升至 `0.7.0`
- 项目结构重构

## [0.6.3] - 2026-07-05

### 新增

- 新增集中式 UI 主题系统
"""


class ExtractReleaseNotesTests(unittest.TestCase):
    def test_extracts_section_for_matching_version(self):
        notes = extract_release_notes(CHANGELOG, "0.7.1")

        self.assertEqual(notes, "### 新增\n\n- 处理前检测已签名文件")

    def test_extracts_middle_section_up_to_next_header(self):
        notes = extract_release_notes(CHANGELOG, "0.7.0")

        self.assertIn("- 应用版本号提升至 `0.7.0`", notes)
        self.assertIn("- 项目结构重构", notes)
        self.assertNotIn("集中式 UI 主题系统", notes)

    def test_extracts_last_section_up_to_end_of_file(self):
        notes = extract_release_notes(CHANGELOG, "0.6.3")

        self.assertEqual(notes, "### 新增\n\n- 新增集中式 UI 主题系统")

    def test_returns_none_when_version_missing(self):
        self.assertIsNone(extract_release_notes(CHANGELOG, "9.9.9"))

    def test_does_not_match_version_as_substring(self):
        # 0.7.1 不应被 0.7.10 之类的查询匹配到，也不应反向误匹配
        self.assertIsNone(extract_release_notes(CHANGELOG, "0.7"))


if __name__ == "__main__":
    unittest.main()
