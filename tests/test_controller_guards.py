import unittest

from controller import _build_io_paths_for_file, _collect_ectd_rename_plan, _normalized_ectd_name


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


if __name__ == "__main__":
    unittest.main()
