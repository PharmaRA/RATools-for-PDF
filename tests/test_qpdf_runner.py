"""qpdf 底层引擎与命令构造器单元测试。"""

import os
import tempfile
import unittest

import fitz

from ratools_pdf.pdf import qpdf


class QpdfRunnerTests(unittest.TestCase):
    def test_get_qpdf_path_returns_existing_path(self):
        path = qpdf.get_qpdf_path()
        self.assertTrue(path, "未能返回 qpdf 路径")
        self.assertTrue(os.path.exists(path), f"返回的 qpdf 路径不存在: {path}")

    def test_command_builder_args(self):
        builder = (
            qpdf.QpdfCommandBuilder(input_pdf="in.pdf", output_pdf="out.pdf")
            .set_linearize(True)
            .set_force_version("1.7")
            .set_decrypt_restrictions(True)
            .set_flatten_rotation(True)
            .set_flatten_annotations("all")
            .set_object_streams("generate")
            .set_recompress_flate(True, compression_level=9)
            .set_stream_data("compress")
        )
        args = builder.build_args()

        self.assertIn("--linearize", args)
        self.assertIn("--force-version=1.7", args)
        self.assertIn("--decrypt", args)
        self.assertIn("--flatten-rotation", args)
        self.assertIn("--flatten-annotations=all", args)
        self.assertIn("--object-streams=generate", args)
        self.assertIn("--recompress-flate", args)
        self.assertIn("--compression-level=9", args)
        self.assertIn("--stream-data=compress", args)
        self.assertEqual(args[-2:], ["in.pdf", "out.pdf"])

    def test_check_pdf_syntax_healthy_file(self):
        with tempfile.TemporaryDirectory() as tmp:
            pdf_path = os.path.join(tmp, "healthy.pdf")
            doc = fitz.open()
            doc.new_page()
            doc.save(pdf_path)
            doc.close()

            report = qpdf.check_pdf_syntax(pdf_path)
            self.assertTrue(report["is_healthy"])
            self.assertEqual(report["returncode"], 0)

    def test_inspect_pdf_json(self):
        with tempfile.TemporaryDirectory() as tmp:
            pdf_path = os.path.join(tmp, "sample.pdf")
            doc = fitz.open()
            page = doc.new_page()
            page.insert_text((50, 50), "Test qpdf json")
            doc.save(pdf_path)
            doc.close()

            ast = qpdf.inspect_pdf_json(pdf_path)
            self.assertIsInstance(ast, dict)
            self.assertIn("pages", ast)
            self.assertEqual(len(ast["pages"]), 1)

    def test_format_qpdf_error(self):
        err1 = RuntimeError("qpdf 执行失败: invalid password")
        self.assertIn("需要密码", qpdf.format_qpdf_error(err1))

        err2 = RuntimeError("qpdf 执行失败: broken xref")
        self.assertIn("未能移除PDF权限限制", qpdf.format_qpdf_error(err2))


if __name__ == "__main__":
    unittest.main()
