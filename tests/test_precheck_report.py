"""build_precheck_report 检查项的特征测试。

固化各检查项的触发/不触发条件，为 precheck 检查器注册表化重构提供安全网。
现有覆盖（tests/test_precheck_annotation_broken_reference.py）只涉及批注与
失效引用两个 report_only 项，这里补齐常规选项检查。
"""

import os
import tempfile
import unittest

import fitz

from ratools_pdf.pdf.precheck import build_precheck_report


def _report(path, selected=None):
    return build_precheck_report(path, selected_options=selected)


def _suggested_ids(report):
    return set(report.get("suggestions", {}).keys())


class PrecheckDocumentPropertyTests(unittest.TestCase):
    def test_title_mismatch_triggers_suggestion(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "目标文件名.pdf")
            doc = fitz.open()
            doc.new_page()
            doc.set_metadata({"title": "别的标题"})
            doc.save(path)
            doc.close()

            report = _report(path)

            self.assertTrue(report["available"])
            self.assertIn("title_from_filename", _suggested_ids(report))

    def test_title_match_no_suggestion(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "match.pdf")
            doc = fitz.open()
            doc.new_page()
            doc.set_metadata({"title": "match"})
            doc.save(path)
            doc.close()

            report = _report(path)

            self.assertNotIn("title_from_filename", _suggested_ids(report))

    def test_bookmarks_without_useoutlines_triggers_initial_view(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "s.pdf")
            doc = fitz.open()
            doc.new_page()
            doc.new_page()
            doc.set_toc([[1, "Chapter", 1]])
            doc.save(path)
            doc.close()

            report = _report(path)

            self.assertIn("initial_view_bookmarks_and_page", _suggested_ids(report))

    def test_explicit_page_layout_triggers_suggestion(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "s.pdf")
            doc = fitz.open()
            doc.new_page()
            doc.xref_set_key(doc.pdf_catalog(), "PageLayout", "/TwoColumnLeft")
            doc.save(path)
            doc.close()

            report = _report(path)

            self.assertIn("page_layout_default", _suggested_ids(report))

    def test_missing_open_action_no_first_page_suggestion(self):
        # 无 OpenAction 时，现状行为是不触发 open_page_first / zoom_default
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "s.pdf")
            doc = fitz.open()
            doc.new_page()
            doc.save(path)
            doc.close()

            report = _report(path)

            ids = _suggested_ids(report)
            self.assertNotIn("open_page_first", ids)
            self.assertNotIn("zoom_default", ids)

    def test_fixed_zoom_open_action_triggers_zoom_default(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "s.pdf")
            doc = fitz.open()
            doc.new_page()
            doc.new_page()
            page0_xref = doc[0].xref
            doc.xref_set_key(
                doc.pdf_catalog(), "OpenAction", f"[{page0_xref} 0 R /XYZ 0 792 2.5]"
            )
            doc.save(path)
            doc.close()

            report = _report(path)

            ids = _suggested_ids(report)
            self.assertIn("zoom_default", ids)
            self.assertNotIn("open_page_first", ids)


class PrecheckBookmarkTests(unittest.TestCase):
    def _make_pdf(self, tmp, toc):
        path = os.path.join(tmp, "s.pdf")
        doc = fitz.open()
        doc.new_page()
        doc.new_page()
        doc.set_toc(toc)
        doc.save(path)
        doc.close()
        return path

    def test_fixed_zoom_bookmark_triggers_inherit_zoom(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = self._make_pdf(tmp, [
                [1, "Zoomed", 2, {"kind": fitz.LINK_GOTO, "page": 1, "to": fitz.Point(72, 72), "zoom": 2.0}],
            ])

            report = _report(path)

            self.assertIn("bookmark_inherit_zoom", _suggested_ids(report))

    def test_external_uri_bookmark_triggers_remove_external(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = self._make_pdf(tmp, [
                [1, "External", -1, {"kind": fitz.LINK_URI, "uri": "https://example.com"}],
            ])

            report = _report(path)

            ids = _suggested_ids(report)
            self.assertIn("bookmark_remove_external_links", ids)
            # URI 同时属于"非标准动作"
            self.assertIn("bookmark_remove_unknown_actions", ids)

    def test_clean_internal_bookmarks_no_bookmark_suggestions(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = self._make_pdf(tmp, [
                [1, "OK", 1, {"kind": fitz.LINK_GOTO, "page": 0, "to": fitz.Point(0, 0), "zoom": 0.0}],
            ])

            report = _report(path)

            ids = _suggested_ids(report)
            self.assertNotIn("bookmark_inherit_zoom", ids)
            self.assertNotIn("bookmark_remove_external_links", ids)
            self.assertNotIn("bookmark_remove_invalid", ids)

    def _make_named_dest_pdf(self, tmp, dest_names):
        """构造书签动作为命名目标的 PDF。

        set_toc 无法写出命名目标，只能保存后直接改写 outline 的 /A。
        catalog 的 /Dests 里只登记 RealDest，因此 NoSuchDest 是悬空目标。
        """
        path = os.path.join(tmp, "named.pdf")
        doc = fitz.open()
        doc.new_page()
        doc.new_page()
        doc.set_toc([[1, name, 1] for name in dest_names])
        doc.save(path)
        doc.close()

        doc = fitz.open(path)
        page1_xref = doc[1].xref
        dests = doc.get_new_xref()
        doc.update_object(dests, f"<< /RealDest [{page1_xref} 0 R /Fit] >>")
        doc.xref_set_key(doc.pdf_catalog(), "Dests", f"{dests} 0 R")
        for item, name in zip(doc.get_toc(simple=False), dest_names):
            doc.xref_set_key(item[3]["xref"], "Dest", "null")
            doc.xref_set_key(item[3]["xref"], "A", f"<< /S /GoTo /D ({name}) >>")
        doc.saveIncr()
        doc.close()
        return path

    def test_dangling_named_dest_bookmark_triggers_remove_invalid(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = self._make_named_dest_pdf(tmp, ["NoSuchDest"])

            report = _report(path)

            self.assertIn("bookmark_remove_invalid", _suggested_ids(report))

    def test_resolvable_named_dest_bookmark_no_remove_invalid(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = self._make_named_dest_pdf(tmp, ["RealDest"])

            report = _report(path)

            self.assertNotIn("bookmark_remove_invalid", _suggested_ids(report))


class PrecheckLinkTests(unittest.TestCase):
    def test_absolute_file_link_triggers_abs_to_rel(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "s.pdf")
            doc = fitz.open()
            page = doc.new_page()
            page.insert_link({
                "kind": fitz.LINK_GOTOR,
                "from": fitz.Rect(50, 50, 150, 70),
                "file": "C:/abs/path/other.pdf",
                "page": 0,
            })
            doc.save(path)
            doc.close()

            report = _report(path)

            self.assertIn("link_abs_to_rel_path", _suggested_ids(report))

    def test_external_uri_link_triggers_cleanup_suggestion(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "s.pdf")
            doc = fitz.open()
            page = doc.new_page()
            page.insert_link({
                "kind": fitz.LINK_URI,
                "from": fitz.Rect(50, 50, 150, 70),
                "uri": "https://example.com",
            })
            doc.save(path)
            doc.close()

            report = _report(path)

            self.assertIn("cleanup_remove_external_uri", _suggested_ids(report))

    def test_fixed_zoom_goto_link_not_detected_current_behavior(self):
        # 已知局限（特征测试固化现状）：PyMuPDF 1.27 的 get_links() 不回读
        # /XYZ 目标里的 zoom（恒为 0.0），因此链接固定缩放目前无法被预检发现。
        # 书签的 zoom 经 get_toc(simple=False) 可正常回读，不受影响。
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "s.pdf")
            doc = fitz.open()
            doc.new_page()
            doc.new_page()
            doc[0].insert_link({
                "kind": fitz.LINK_GOTO,
                "from": fitz.Rect(50, 50, 150, 70),
                "page": 1,
                "to": fitz.Point(72, 144),
                "zoom": 2.0,
            })
            doc.save(path)
            doc.close()

            probe = fitz.open(path)
            stored = probe.xref_object(probe[0].get_links()[0]["xref"])
            probe.close()
            self.assertIn("/XYZ 72 698 2", stored.replace("\n", " ").replace("  ", " "))

            report = _report(path)

            self.assertNotIn("link_inherit_zoom", _suggested_ids(report))


class PrecheckCleanupTests(unittest.TestCase):
    def test_attachments_trigger_suggestion_with_count(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "s.pdf")
            doc = fitz.open()
            doc.new_page()
            doc.embfile_add("a.txt", b"data-a")
            doc.embfile_add("b.txt", b"data-b")
            doc.save(path)
            doc.close()

            report = _report(path)

            suggestion = report["suggestions"].get("cleanup_remove_attachments")
            self.assertIsNotNone(suggestion)
            self.assertIn("2", suggestion["reason"])

    def test_metadata_triggers_cleanup_suggestion(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "s.pdf")
            doc = fitz.open()
            doc.new_page()
            doc.set_metadata({"author": "Somebody"})
            doc.save(path)
            doc.close()

            report = _report(path)

            self.assertIn("cleanup_remove_metadata", _suggested_ids(report))

    def test_tags_trigger_cleanup_suggestion(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "s.pdf")
            doc = fitz.open()
            doc.new_page()
            doc.xref_set_key(doc.pdf_catalog(), "MarkInfo", "<< /Marked true >>")
            doc.save(path)
            doc.close()

            report = _report(path)

            self.assertIn("cleanup_remove_tags", _suggested_ids(report))

    def test_javascript_names_trigger_dynamic_content_suggestion(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "s.pdf")
            doc = fitz.open()
            doc.new_page()
            doc.xref_set_key(
                doc.pdf_catalog(), "Names", "<< /JavaScript << /Names [] >> >>"
            )
            doc.save(path)
            doc.close()

            report = _report(path)

            self.assertIn("cleanup_remove_dynamic_content", _suggested_ids(report))

    def test_text_annotation_triggers_remove_annotations(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "s.pdf")
            doc = fitz.open()
            page = doc.new_page()
            page.add_text_annot(fitz.Point(60, 60), "a sticky note")
            doc.save(path)
            doc.close()

            report = _report(path)

            self.assertIn("cleanup_remove_annotations", _suggested_ids(report))

    def test_highlight_only_pdf_misses_remove_annotations_current_behavior(self):
        # 已知缺陷（特征测试固化现状）：build_precheck_report 的注释扫描用
        # annot.type[0] == 8 判定"链接注释"，但 page.annots() 里 8 是高亮
        # （Link 注释根本不会出现在 annots() 中，见 ANNOTATION_TYPE_LABELS 注释）。
        # 结果：仅含高亮批注的 PDF 不会触发 cleanup_remove_annotations 建议。
        # report_only 的批注复核项（annotation_precheck_review）不受影响。
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "s.pdf")
            doc = fitz.open()
            page = doc.new_page()
            page.insert_text(fitz.Point(55, 65), "mark", fontsize=12)
            page.add_highlight_annot(fitz.Rect(50, 50, 120, 72))
            doc.save(path)
            doc.close()

            report = _report(path)

            ids = _suggested_ids(report)
            self.assertNotIn("cleanup_remove_annotations", ids)
            self.assertIn("annotation_precheck_review", ids)


class PrecheckFileLevelTests(unittest.TestCase):
    def test_old_pdf_version_triggers_convert_suggestion(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "s.pdf")
            doc = fitz.open()
            doc.new_page()
            doc.save(path)
            doc.close()
            # 检查只读文件头前 32 字节，二进制补丁把版本头改成 1.4 即可模拟旧版
            with open(path, "r+b") as f:
                header = f.read(8)
                assert header == b"%PDF-1.7", header
                f.seek(0)
                f.write(b"%PDF-1.4")

            report = _report(path)

            self.assertIn("convert_pdf_version", _suggested_ids(report))

    def test_pdf_17_version_no_convert_suggestion(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "s.pdf")
            doc = fitz.open()
            doc.new_page()
            doc.save(path)
            doc.close()

            report = _report(path)

            self.assertNotIn("convert_pdf_version", _suggested_ids(report))

    def test_non_linearized_pdf_triggers_fast_web_view(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "s.pdf")
            doc = fitz.open()
            doc.new_page()
            doc.save(path)
            doc.close()

            report = _report(path)

            self.assertIn("fast_web_view", _suggested_ids(report))

    def test_missing_file_reports_unavailable(self):
        report = _report(os.path.join(tempfile.gettempdir(), "does_not_exist_x.pdf"))

        self.assertFalse(report["available"])
        self.assertEqual(report["error"], "文件不存在")

    def test_password_protected_file_reports_unavailable(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "s.pdf")
            doc = fitz.open()
            doc.new_page()
            doc.save(path, encryption=fitz.PDF_ENCRYPT_AES_256, user_pw="secret")
            doc.close()

            report = _report(path)

            self.assertFalse(report["available"])
            self.assertIn("密码", report["error"])


class PrecheckSelectedOptionsTests(unittest.TestCase):
    def test_selected_options_limits_scanned_checks(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "别名.pdf")
            doc = fitz.open()
            page = doc.new_page()
            doc.set_metadata({"title": "不一致"})
            page.insert_link({
                "kind": fitz.LINK_URI,
                "from": fitz.Rect(50, 50, 150, 70),
                "uri": "https://example.com",
            })
            doc.save(path)
            doc.close()

            report = _report(path, selected={"title_from_filename"})

            ids = _suggested_ids(report)
            self.assertIn("title_from_filename", ids)
            # 未勾选的 URI 清理项不应出现
            self.assertNotIn("cleanup_remove_external_uri", ids)

    def test_alias_option_recorded_under_canonical_id(self):
        # cleanup_remove_external_uri_and_text_black 是 cleanup_remove_external_uri
        # 的别名变体。现状行为：别名会展开出规范 id 一起进入过滤集，建议最终
        # 记录在规范 id 下（调用方通过别名映射反查）。
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "s.pdf")
            doc = fitz.open()
            page = doc.new_page()
            page.insert_link({
                "kind": fitz.LINK_URI,
                "from": fitz.Rect(50, 50, 150, 70),
                "uri": "https://example.com",
            })
            doc.save(path)
            doc.close()

            report = _report(path, selected={"cleanup_remove_external_uri_and_text_black"})

            self.assertIn("cleanup_remove_external_uri", _suggested_ids(report))

    def test_full_precheck_includes_report_only_findings_key(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "s.pdf")
            doc = fitz.open()
            page = doc.new_page()
            page.insert_text(fitz.Point(55, 65), "note", fontsize=12)
            page.add_text_annot(fitz.Point(60, 60), "a sticky note")
            doc.save(path)
            doc.close()

            report = _report(path)

            suggestion = report["suggestions"].get("annotation_precheck_review")
            self.assertIsNotNone(suggestion)
            self.assertTrue(suggestion.get("report_only"))


if __name__ == "__main__":
    unittest.main()
