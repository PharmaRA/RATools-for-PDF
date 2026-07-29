"""process_document 主管线的特征测试（characterization tests）。

在重构 processor.py 之前固化现有行为：每组规则合成一个最小 PDF，
用 force 模式执行处理，断言输出 PDF 的结构状态与返回消息。
所有 PDF 均由 fitz 运行时合成，不依赖仓库内二进制样例。
"""

import os
import tempfile
import unittest

import fitz

from ratools_pdf.pdf.processor import PDFProcessor


def _save_new_pdf(path, page_count=1, page_size=None):
    doc = fitz.open()
    for _ in range(page_count):
        if page_size:
            doc.new_page(width=page_size[0], height=page_size[1])
        else:
            doc.new_page()
    doc.save(path)
    doc.close()


def _process(source, output, options, mode="force"):
    ok, msg = PDFProcessor.process_document(source, output, set(options), processing_mode=mode)
    return ok, msg


class InitialViewRulesTests(unittest.TestCase):
    def test_title_from_filename_sets_metadata_title(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = os.path.join(tmp, "报告A.pdf")
            output = os.path.join(tmp, "out.pdf")
            _save_new_pdf(source)

            ok, msg = _process(source, output, {"title_from_filename"})

            self.assertTrue(ok, msg)
            doc = fitz.open(output)
            self.assertEqual(doc.metadata.get("title"), "报告A")
            doc.close()
            self.assertIn("标题同步为文件名", msg)

    def test_title_from_filename_no_change_when_title_matches(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = os.path.join(tmp, "same.pdf")
            output = os.path.join(tmp, "out.pdf")
            doc = fitz.open()
            doc.new_page()
            doc.set_metadata({"title": "same"})
            doc.save(source)
            doc.close()

            ok, msg = _process(source, output, {"title_from_filename"})

            self.assertTrue(ok, msg)
            self.assertIn("无实际修改", msg)
            self.assertTrue(os.path.exists(output))

    def test_open_page_first_and_zoom_default_write_open_action(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = os.path.join(tmp, "s.pdf")
            output = os.path.join(tmp, "out.pdf")
            _save_new_pdf(source, page_count=2)

            ok, msg = _process(source, output, {"open_page_first", "zoom_default"})

            self.assertTrue(ok, msg)
            doc = fitz.open(output)
            catalog = doc.pdf_catalog()
            kind, value = doc.xref_get_key(catalog, "OpenAction")
            first_page_xref = doc[0].xref
            doc.close()
            self.assertNotEqual(kind, "null")
            compact = value.replace(" ", "")
            self.assertIn(f"{first_page_xref}0R", compact)
            self.assertIn("/XYZnullnullnull", compact)
            self.assertIn("打开页设为第一页", msg)
            self.assertIn("打开缩放设为默认", msg)

    def test_page_layout_default_removes_explicit_layout(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = os.path.join(tmp, "s.pdf")
            output = os.path.join(tmp, "out.pdf")
            doc = fitz.open()
            doc.new_page()
            doc.xref_set_key(doc.pdf_catalog(), "PageLayout", "/TwoColumnLeft")
            doc.save(source)
            doc.close()

            ok, msg = _process(source, output, {"page_layout_default"})

            self.assertTrue(ok, msg)
            doc = fitz.open(output)
            kind, _value = doc.xref_get_key(doc.pdf_catalog(), "PageLayout")
            doc.close()
            self.assertEqual(kind, "null")
            self.assertIn("页面布局恢复默认", msg)

    def test_initial_view_uses_outlines_when_bookmarks_exist(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = os.path.join(tmp, "s.pdf")
            output = os.path.join(tmp, "out.pdf")
            doc = fitz.open()
            doc.new_page()
            doc.new_page()
            doc.set_toc([[1, "第一章", 1]])
            doc.save(source)
            doc.close()

            ok, msg = _process(source, output, {"initial_view_bookmarks_and_page"})

            self.assertTrue(ok, msg)
            doc = fitz.open(output)
            kind, value = doc.xref_get_key(doc.pdf_catalog(), "PageMode")
            doc.close()
            self.assertEqual((kind, value), ("name", "/UseOutlines"))

    def test_initial_view_uses_none_without_bookmarks(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = os.path.join(tmp, "s.pdf")
            output = os.path.join(tmp, "out.pdf")
            _save_new_pdf(source)

            ok, msg = _process(source, output, {"initial_view_bookmarks_and_page"})

            self.assertTrue(ok, msg)
            doc = fitz.open(output)
            kind, value = doc.xref_get_key(doc.pdf_catalog(), "PageMode")
            doc.close()
            self.assertEqual((kind, value), ("name", "/UseNone"))

    def test_collapse_all_bookmarks_marks_items_collapsed(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = os.path.join(tmp, "s.pdf")
            output = os.path.join(tmp, "out.pdf")
            doc = fitz.open()
            for _ in range(3):
                doc.new_page()
            doc.set_toc([
                [1, "Chapter", 1],
                [2, "Section", 2],
            ])
            doc.save(source)
            doc.close()

            ok, msg = _process(source, output, {"collapse_all_bookmarks"})

            self.assertTrue(ok, msg)
            doc = fitz.open(output)
            toc = doc.get_toc(simple=False)
            doc.close()
            # collapse 标志只在有子节点的书签上可回读；叶子节点无 /Count 键
            self.assertTrue(toc[0][-1].get("collapse"), f"父书签未折叠: {toc[0]}")
            self.assertIn("折叠全部书签", msg)


class PageSizeRulesTests(unittest.TestCase):
    def test_page_size_a4_resizes_nonstandard_page(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = os.path.join(tmp, "s.pdf")
            output = os.path.join(tmp, "out.pdf")
            _save_new_pdf(source, page_size=(500, 700))

            ok, msg = _process(source, output, {"page_size_a4"})

            self.assertTrue(ok, msg)
            doc = fitz.open(output)
            rect = doc[0].rect
            doc.close()
            self.assertAlmostEqual(rect.width, 210 / 25.4 * 72, delta=0.5)
            self.assertAlmostEqual(rect.height, 297 / 25.4 * 72, delta=0.5)
            self.assertIn("页面尺寸标准化", msg)

    def test_page_size_letter_keeps_landscape_orientation(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = os.path.join(tmp, "s.pdf")
            output = os.path.join(tmp, "out.pdf")
            _save_new_pdf(source, page_size=(700, 500))

            ok, msg = _process(source, output, {"page_size_letter"})

            self.assertTrue(ok, msg)
            doc = fitz.open(output)
            rect = doc[0].rect
            doc.close()
            # 横向页面应映射到横向 Letter（11x8.5 英寸）
            self.assertAlmostEqual(rect.width, 11 * 72, delta=0.5)
            self.assertAlmostEqual(rect.height, 8.5 * 72, delta=0.5)

    def test_page_size_a4_no_change_for_a4_page(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = os.path.join(tmp, "s.pdf")
            output = os.path.join(tmp, "out.pdf")
            _save_new_pdf(source, page_size=(210 / 25.4 * 72, 297 / 25.4 * 72))

            ok, msg = _process(source, output, {"page_size_a4"})

            self.assertTrue(ok, msg)
            self.assertIn("无实际修改", msg)


class BookmarkRulesTests(unittest.TestCase):
    def test_bookmark_remove_external_links_drops_uri_bookmarks(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = os.path.join(tmp, "s.pdf")
            output = os.path.join(tmp, "out.pdf")
            doc = fitz.open()
            doc.new_page()
            doc.set_toc([
                [1, "External", -1, {"kind": fitz.LINK_URI, "uri": "https://example.com"}],
                [1, "Internal", 1, {"kind": fitz.LINK_GOTO, "page": 0, "to": fitz.Point(0, 0), "zoom": 0.0}],
            ])
            doc.save(source)
            doc.close()

            ok, msg = _process(source, output, {"bookmark_remove_external_links"})

            self.assertTrue(ok, msg)
            doc = fitz.open(output)
            toc = doc.get_toc(simple=False)
            doc.close()
            titles = [item[1] for item in toc]
            self.assertEqual(titles, ["Internal"])
            self.assertIn("书签规则已更新", msg)

    def test_bookmark_inherit_zoom_resets_fixed_zoom(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = os.path.join(tmp, "s.pdf")
            output = os.path.join(tmp, "out.pdf")
            doc = fitz.open()
            doc.new_page()
            doc.new_page()
            doc.set_toc([
                [1, "Zoomed", 2, {"kind": fitz.LINK_GOTO, "page": 1, "to": fitz.Point(72, 72), "zoom": 2.5}],
            ])
            doc.save(source)
            doc.close()

            ok, msg = _process(source, output, {"bookmark_inherit_zoom"})

            self.assertTrue(ok, msg)
            doc = fitz.open(output)
            toc = doc.get_toc(simple=False)
            doc.close()
            self.assertEqual(toc[0][3].get("zoom"), 0.0)

    def test_bookmark_remove_unknown_actions_drops_uri_and_none(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = os.path.join(tmp, "s.pdf")
            output = os.path.join(tmp, "out.pdf")
            doc = fitz.open()
            doc.new_page()
            doc.set_toc([
                [1, "External", -1, {"kind": fitz.LINK_URI, "uri": "https://example.com"}],
                [1, "Internal", 1, {"kind": fitz.LINK_GOTO, "page": 0, "to": fitz.Point(0, 0), "zoom": 0.0}],
            ])
            doc.save(source)
            doc.close()

            # URI 属于"未知动作"白名单（GOTO/GOTOR/LAUNCH）之外，应被删除
            ok, msg = _process(source, output, {"bookmark_remove_unknown_actions"})

            self.assertTrue(ok, msg)
            doc = fitz.open(output)
            toc = doc.get_toc(simple=False)
            doc.close()
            self.assertEqual([item[1] for item in toc], ["Internal"])

    def _make_named_dest_pdf(self, tmp, real_dest_view="/Fit"):
        """构造带命名目标书签的 PDF：一个悬空、一个可解析。

        真实 PDF 里书签常用 ``/GoTo`` + 命名目标，PyMuPDF 会报 LINK_NAMED；
        set_toc 写不出这种形状，只能直接改写 outline 对象。
        real_dest_view 用于切换 RealDest 的视图（如 /XYZ 72 720 2.5 带固定缩放）。
        """
        source = os.path.join(tmp, "named.pdf")
        doc = fitz.open()
        doc.new_page()
        doc.new_page()
        doc.set_toc([
            [1, "dangling-named", 1],
            [1, "valid-named", 2],
            [1, "internal", 1],
        ])
        doc.save(source)
        doc.close()

        doc = fitz.open(source)
        page1_xref = doc[1].xref
        # catalog /Dests 字典：RealDest 可解析到第 2 页，NoSuchDest 不存在
        dests_xref = doc.get_new_xref()
        doc.update_object(dests_xref, f"<< /RealDest [{page1_xref} 0 R {real_dest_view}] >>")
        doc.xref_set_key(doc.pdf_catalog(), "Dests", f"{dests_xref} 0 R")

        outline_xrefs = []
        for xref in range(1, doc.xref_length()):
            if doc.xref_get_key(xref, "Title")[0] != "null":
                outline_xrefs.append((doc.xref_get_key(xref, "Title")[1], xref))
        by_title = {title: xref for title, xref in outline_xrefs}

        def _outline_xref(label):
            for title, xref in by_title.items():
                if label in title:
                    return xref
            raise AssertionError(f"未找到书签对象: {label} in {list(by_title)}")

        doc.xref_set_key(_outline_xref("dangling-named"), "Dest", "null")
        doc.xref_set_key(
            _outline_xref("dangling-named"), "A", "<< /S /GoTo /D (NoSuchDest) >>"
        )
        doc.xref_set_key(_outline_xref("valid-named"), "Dest", "null")
        doc.xref_set_key(
            _outline_xref("valid-named"), "A", "<< /S /GoTo /D (RealDest) >>"
        )
        doc.saveIncr()
        doc.close()

        doc = fitz.open(source)
        kinds = {item[1]: item[3].get("kind") for item in doc.get_toc(simple=False)}
        doc.close()
        # 前置条件：两个命名目标书签都必须被报成 LINK_NAMED，否则用例失去意义
        assert kinds.get("dangling-named") == fitz.LINK_NAMED, kinds
        assert kinds.get("valid-named") == fitz.LINK_NAMED, kinds
        return source

    def test_bookmark_remove_invalid_drops_dangling_named_dest(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = self._make_named_dest_pdf(tmp)
            output = os.path.join(tmp, "out.pdf")

            ok, msg = _process(source, output, {"bookmark_remove_invalid"})

            self.assertTrue(ok, msg)
            doc = fitz.open(output)
            toc = doc.get_toc(simple=False)
            doc.close()
            titles = [item[1] for item in toc]
            # 悬空命名目标被删除，可解析的命名目标与普通内部书签都要保留
            self.assertEqual(titles, ["valid-named", "internal"])

    def test_bookmark_remove_invalid_keeps_resolvable_named_dest_target(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = self._make_named_dest_pdf(tmp)
            output = os.path.join(tmp, "out.pdf")

            ok, msg = _process(source, output, {"bookmark_remove_invalid"})

            self.assertTrue(ok, msg)
            doc = fitz.open(output)
            toc = doc.get_toc(simple=False)
            doc.close()
            kept = {item[1]: item for item in toc}
            # 命名目标降级为等价内部跳转后，仍须指向原来的第 2 页
            self.assertEqual(kept["valid-named"][2], 2)
            self.assertEqual(kept["valid-named"][3].get("kind"), fitz.LINK_GOTO)

    def test_bookmark_rules_untouched_when_invalid_option_absent(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = self._make_named_dest_pdf(tmp)
            output = os.path.join(tmp, "out.pdf")

            # 未勾选"清理失效书签"时，悬空书签不应被顺带删除
            ok, msg = _process(source, output, {"bookmark_remove_external_links"})

            self.assertTrue(ok, msg)
            doc = fitz.open(output)
            titles = [item[1] for item in doc.get_toc(simple=False)]
            doc.close()
            self.assertIn("dangling-named", titles)

    def test_collapse_all_bookmarks_keeps_named_dest_target(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = self._make_named_dest_pdf(tmp)
            output = os.path.join(tmp, "out.pdf")

            # 折叠步骤同样要经过目的地规整：原样回写会把命名目标
            # 退化成无目标空书签（静默丢失跳转）
            ok, msg = _process(source, output, {"collapse_all_bookmarks"})

            self.assertTrue(ok, msg)
            doc = fitz.open(output)
            kept = {item[1]: item for item in doc.get_toc(simple=False)}
            doc.close()
            self.assertEqual(kept["valid-named"][2], 2)
            self.assertEqual(kept["valid-named"][3].get("kind"), fitz.LINK_GOTO)

    def test_bookmark_inherit_zoom_resets_named_dest_fixed_zoom(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = self._make_named_dest_pdf(tmp, real_dest_view="/XYZ 72 720 2.5")
            output = os.path.join(tmp, "out.pdf")

            ok, msg = _process(source, output, {"bookmark_inherit_zoom"})

            self.assertTrue(ok, msg)
            doc = fitz.open(output)
            kept = {item[1]: item for item in doc.get_toc(simple=False)}
            doc.close()
            # 命名目标的真实缩放 2.5 必须被识别并重置为承前缩放
            self.assertEqual(kept["valid-named"][3].get("zoom"), 0.0)
            self.assertEqual(kept["valid-named"][2], 2)

    def test_bookmark_remove_unknown_actions_keeps_resolvable_named_dest(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = self._make_named_dest_pdf(tmp)
            output = os.path.join(tmp, "out.pdf")

            ok, msg = _process(source, output, {"bookmark_remove_unknown_actions"})

            self.assertTrue(ok, msg)
            doc = fitz.open(output)
            titles = [item[1] for item in doc.get_toc(simple=False)]
            doc.close()
            # 可解析的命名目标等价于内部跳转，不属于非标准动作；仅悬空的应被删除
            self.assertEqual(titles, ["valid-named", "internal"])

    def test_collapse_all_bookmarks_preserves_named_dest_zoom(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = self._make_named_dest_pdf(tmp, real_dest_view="/XYZ 72 720 2.5")
            output = os.path.join(tmp, "out.pdf")

            # 只勾选折叠时不得顺带改动缩放：命名目标的真实缩放要原样保留
            ok, msg = _process(source, output, {"collapse_all_bookmarks"})

            self.assertTrue(ok, msg)
            doc = fitz.open(output)
            kept = {item[1]: item for item in doc.get_toc(simple=False)}
            doc.close()
            self.assertEqual(kept["valid-named"][3].get("zoom"), 2.5)
            self.assertEqual(kept["internal"][3].get("zoom"), 0.0)

    def _make_external_bookmark_pdf(self, tmp, subtype, file_name):
        """构造外部文件书签（/GoToR 或 /Launch）+ 一条普通内部书签。

        set_toc 写不出带 to 的 GOTOR（PyMuPDF 内部把 to 转成 tuple，
        getDestStr 又要求 .x/.y），因此直接改写 outline 的 /A。
        """
        source = os.path.join(tmp, f"{subtype.strip('/')}.pdf")
        doc = fitz.open()
        doc.new_page()
        doc.new_page()
        doc.set_toc([[1, "ext", 1], [1, "internal", 2]])
        doc.save(source)
        doc.close()

        doc = fitz.open(source)
        ext_xref = None
        for xref in range(1, doc.xref_length()):
            title = doc.xref_get_key(xref, "Title")
            if title[0] != "null" and "ext" in title[1]:
                ext_xref = xref
                break
        assert ext_xref is not None, "未找到 ext 书签对象"
        doc.xref_set_key(ext_xref, "Dest", "null")
        fspec = f"<</Type/Filespec/F({file_name})/UF({file_name})>>"
        if subtype == "/GoToR":
            action = f"<</S/GoToR/D[0/XYZ 72 720 0]/F{fspec}>>"
        else:
            action = f"<</S/Launch/F{fspec}>>"
        doc.xref_set_key(ext_xref, "A", action)
        doc.saveIncr()
        doc.close()
        return source

    def _ext_bookmark_state(self, path):
        doc = fitz.open(path)
        state = {}
        for item in doc.get_toc(simple=False):
            dest = item[3]
            xref = dest.get("xref")
            state[item[1]] = {
                "kind": dest.get("kind"),
                "file": dest.get("file"),
                "subtype": doc.xref_get_key(int(xref), "A/S")[1] if xref else None,
                "new_window": doc.xref_get_key(int(xref), "A/NewWindow")[1] if xref else None,
            }
        doc.close()
        return state

    def test_bookmark_open_new_window_keeps_gotor_file_and_sets_flag(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = self._make_external_bookmark_pdf(tmp, "/GoToR", "other.pdf")
            output = os.path.join(tmp, "out.pdf")

            ok, msg = _process(source, output, {"bookmark_open_new_window"})

            self.assertTrue(ok, msg)
            state = self._ext_bookmark_state(output)
            # 写回不得丢失外部文件目标（此前会异常降级为普通内部书签）
            self.assertEqual(state["ext"]["file"], "other.pdf")
            self.assertEqual(state["ext"]["subtype"], "/GoToR")
            # set_toc 从不写出 /NewWindow，必须在写回后补上
            self.assertEqual(state["ext"]["new_window"], "true")

    def test_bookmark_open_new_window_preserves_launch_subtype(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = self._make_external_bookmark_pdf(tmp, "/Launch", "run.exe")
            output = os.path.join(tmp, "out.pdf")

            ok, msg = _process(source, output, {"bookmark_open_new_window"})

            self.assertTrue(ok, msg)
            state = self._ext_bookmark_state(output)
            self.assertEqual(state["ext"]["file"], "run.exe")
            # PyMuPDF 把 /Launch 回读成 GOTOR，写回时须还原子类型而非退化成 /GoToR
            self.assertEqual(state["ext"]["subtype"], "/Launch")
            self.assertEqual(state["ext"]["new_window"], "true")

    def test_other_bookmark_rule_keeps_external_file_target(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = self._make_external_bookmark_pdf(tmp, "/GoToR", "other.pdf")
            output = os.path.join(tmp, "out.pdf")

            # 触发重写的是另一条规则：外部文件书签仍不能被写坏
            ok, msg = _process(source, output, {"bookmark_inherit_zoom"})

            self.assertTrue(ok, msg)
            state = self._ext_bookmark_state(output)
            self.assertEqual(state["ext"]["file"], "other.pdf")
            self.assertEqual(state["ext"]["subtype"], "/GoToR")


class HyperlinkRulesTests(unittest.TestCase):
    def test_link_abs_to_rel_path_strips_directory(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = os.path.join(tmp, "s.pdf")
            output = os.path.join(tmp, "out.pdf")
            doc = fitz.open()
            page = doc.new_page()
            # PyMuPDF 将 file 值 URL 编码存储（"C:" → "C%3A"），处理规则先 unquote 再判定绝对路径
            page.insert_link({
                "kind": fitz.LINK_GOTOR,
                "from": fitz.Rect(50, 50, 150, 70),
                "file": "C:/docs/annex/other.pdf",
                "page": 0,
            })
            doc.save(source)
            doc.close()

            ok, msg = _process(source, output, {"link_abs_to_rel_path"})

            self.assertTrue(ok, msg)
            doc = fitz.open(output)
            link = doc[0].get_links()[0]
            doc.close()
            self.assertEqual(link.get("file", ""), "other.pdf")

    def test_link_inherit_zoom_resets_goto_zoom(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = os.path.join(tmp, "s.pdf")
            output = os.path.join(tmp, "out.pdf")
            doc = fitz.open()
            doc.new_page()
            doc.new_page()
            doc[0].insert_link({
                "kind": fitz.LINK_GOTO,
                "from": fitz.Rect(50, 50, 150, 70),
                "page": 1,
                "to": fitz.Point(72, 144),
                "zoom": 3.0,
            })
            doc.save(source)
            doc.close()

            ok, msg = _process(source, output, {"link_inherit_zoom"})

            self.assertTrue(ok, msg)
            doc = fitz.open(output)
            raw = doc.xref_object(doc[0].get_links()[0]["xref"])
            doc.close()
            # 必须断言原始对象：get_links 的 zoom 恒为 0.0，
            # 用它断言的话规则完全不生效也照样通过
            compact = " ".join(raw.split())
            self.assertIn("/XYZ 72 698 0", compact)

    def _make_named_dest_link_pdf(self, tmp, view="/XYZ 72 720 2.5"):
        """构造动作为命名目标的页面链接。

        insert_link 写不出命名目标，只能存盘后改写链接注释的 /A。
        """
        source = os.path.join(tmp, "s.pdf")
        doc = fitz.open()
        doc.new_page()
        doc.new_page()
        doc[0].insert_link({
            "kind": fitz.LINK_GOTO,
            "from": fitz.Rect(50, 50, 150, 70),
            "page": 1,
            "to": fitz.Point(0, 0),
        })
        doc.save(source)
        doc.close()

        doc = fitz.open(source)
        page1_xref = doc[1].xref
        dests = doc.get_new_xref()
        doc.update_object(dests, f"<< /RealDest [{page1_xref} 0 R {view}] >>")
        doc.xref_set_key(doc.pdf_catalog(), "Dests", f"{dests} 0 R")
        doc.xref_set_key(doc[0].get_links()[0]["xref"], "A", "<< /S /GoTo /D (RealDest) >>")
        doc.saveIncr()
        doc.close()

        doc = fitz.open(source)
        kind = doc[0].get_links()[0].get("kind")
        doc.close()
        # 前置条件：必须被报成 LINK_NAMED，否则用例失去意义
        assert kind == fitz.LINK_NAMED, kind
        return source

    def test_link_inherit_zoom_resets_named_dest_zoom(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = self._make_named_dest_link_pdf(tmp)
            output = os.path.join(tmp, "out.pdf")

            ok, msg = _process(source, output, {"link_inherit_zoom"})

            self.assertTrue(ok, msg)
            doc = fitz.open(output)
            xref = doc[0].get_links()[0]["xref"]
            dest_type, dest_value = doc.xref_get_key(xref, "A/D")
            doc.close()
            self.assertEqual(dest_type, "array")
            # 缩放归零，且跳转点仍是原来的 y=720（不得因坐标翻转而偏移）
            self.assertIn("/XYZ 72 720 0", " ".join(dest_value.split()))

    def test_link_inherit_zoom_keeps_fit_view_untouched(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = self._make_named_dest_link_pdf(tmp, view="/Fit")
            output = os.path.join(tmp, "out.pdf")

            # /Fit 按定义没有固定缩放，不该被改写
            ok, msg = _process(source, output, {"link_inherit_zoom"})

            self.assertTrue(ok, msg)
            doc = fitz.open(output)
            dest_type, dest_value = doc.xref_get_key(doc[0].get_links()[0]["xref"], "A/D")
            doc.close()
            self.assertEqual((dest_type, dest_value), ("string", "RealDest"))

    def test_link_open_new_window_patches_gotor_action(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = os.path.join(tmp, "s.pdf")
            output = os.path.join(tmp, "out.pdf")
            doc = fitz.open()
            page = doc.new_page()
            page.insert_link({
                "kind": fitz.LINK_GOTOR,
                "from": fitz.Rect(50, 50, 150, 70),
                "file": "other.pdf",
                "page": 0,
            })
            doc.save(source)
            doc.close()

            ok, msg = _process(source, output, {"link_open_new_window"})

            self.assertTrue(ok, msg)
            doc = fitz.open(output)
            xref = doc[0].get_links()[0]["xref"]
            raw = doc.xref_object(xref)
            doc.close()
            self.assertIn("/NewWindow true", raw)

    def test_link_remove_border_zeroes_border_width(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = os.path.join(tmp, "s.pdf")
            output = os.path.join(tmp, "out.pdf")
            doc = fitz.open()
            page = doc.new_page()
            page.insert_link({
                "kind": fitz.LINK_URI,
                "from": fitz.Rect(50, 50, 150, 70),
                "uri": "https://example.com",
            })
            doc.save(source)
            doc.close()
            # 手动加一个可见边框。注意持有 page 引用，避免 first_link
            # 的弱引用父对象被回收（doc[0] 是临时 Page）。
            doc = fitz.open(source)
            page = doc[0]
            link_obj = page.first_link
            link_obj.set_border(width=2.0)
            link_obj.set_colors(stroke=(1, 0, 0))
            doc.saveIncr()
            doc.close()

            ok, msg = _process(source, output, {"link_remove_border"})

            self.assertTrue(ok, msg)
            doc = fitz.open(output)
            page = doc[0]
            border = page.first_link.border or {}
            doc.close()
            self.assertEqual(border.get("width", 0), 0)

    def test_link_black_border_adds_border(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = os.path.join(tmp, "s.pdf")
            output = os.path.join(tmp, "out.pdf")
            doc = fitz.open()
            page = doc.new_page()
            page.insert_link({
                "kind": fitz.LINK_URI,
                "from": fitz.Rect(50, 50, 150, 70),
                "uri": "https://example.com",
            })
            doc.save(source)
            doc.close()

            ok, msg = _process(source, output, {"link_black_border"})

            self.assertTrue(ok, msg)
            doc = fitz.open(output)
            page = doc[0]
            border = page.first_link.border or {}
            doc.close()
            self.assertGreater(border.get("width", 0), 0)

    def test_link_text_blue_recolors_text_inside_link(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = os.path.join(tmp, "s.pdf")
            output = os.path.join(tmp, "out.pdf")
            doc = fitz.open()
            page = doc.new_page()
            page.insert_text(fitz.Point(55, 65), "click here", fontsize=12, color=(0, 0, 0))
            page.insert_link({
                "kind": fitz.LINK_URI,
                "from": fitz.Rect(50, 50, 200, 72),
                "uri": "https://example.com",
            })
            doc.save(source)
            doc.close()

            ok, msg = _process(source, output, {"link_text_blue"})

            self.assertTrue(ok, msg)
            doc = fitz.open(output)
            traces = [t for t in doc[0].get_texttrace() if t.get("type") == 0]
            doc.close()
            self.assertTrue(traces)
            color = traces[0].get("color", (0, 0, 0))
            self.assertGreater(color[2], 0.9, f"文字未变蓝: {color}")


class CleanupRulesTests(unittest.TestCase):
    def test_cleanup_remove_external_uri_deletes_uri_links_only(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = os.path.join(tmp, "s.pdf")
            output = os.path.join(tmp, "out.pdf")
            doc = fitz.open()
            doc.new_page()
            doc.new_page()
            doc[0].insert_link({
                "kind": fitz.LINK_URI,
                "from": fitz.Rect(50, 50, 150, 70),
                "uri": "https://example.com",
            })
            doc[0].insert_link({
                "kind": fitz.LINK_GOTO,
                "from": fitz.Rect(50, 100, 150, 120),
                "page": 1,
                "to": fitz.Point(0, 0),
            })
            doc.save(source)
            doc.close()

            ok, msg = _process(source, output, {"cleanup_remove_external_uri"})

            self.assertTrue(ok, msg)
            doc = fitz.open(output)
            kinds = [l.get("kind") for l in doc[0].get_links()]
            doc.close()
            self.assertEqual(kinds, [fitz.LINK_GOTO])
            self.assertIn("已删除外部URI链接", msg)

    def test_cleanup_remove_unknown_action_links_deletes_named_action(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = os.path.join(tmp, "s.pdf")
            output = os.path.join(tmp, "out.pdf")
            doc = fitz.open()
            page = doc.new_page()
            page.insert_link({
                "kind": fitz.LINK_URI,
                "from": fitz.Rect(50, 50, 150, 70),
                "uri": "https://example.com",
            })
            doc.save(source)
            doc.close()
            # 把动作改成指向不存在命名目标的 GoTo：PyMuPDF 将其报告为
            # LINK_NAMED，落在 GOTO/GOTOR/LAUNCH 白名单之外，属于"未知动作"
            doc = fitz.open(source)
            xref = doc[0].get_links()[0]["xref"]
            doc.xref_set_key(xref, "A", "<< /S /GoTo /D (missing_dest) >>")
            doc.saveIncr()
            doc.close()

            doc = fitz.open(source)
            kinds = [l.get("kind") for l in doc[0].get_links()]
            doc.close()
            self.assertEqual(kinds, [fitz.LINK_NAMED])

            ok, msg = _process(source, output, {"cleanup_remove_unknown_action_links"})

            self.assertTrue(ok, msg)
            doc = fitz.open(output)
            links = doc[0].get_links()
            doc.close()
            self.assertEqual(links, [])
            self.assertIn("已删除未知动作链接", msg)

    def test_cleanup_remove_annotations_deletes_highlight(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = os.path.join(tmp, "s.pdf")
            output = os.path.join(tmp, "out.pdf")
            doc = fitz.open()
            page = doc.new_page()
            page.insert_text(fitz.Point(55, 65), "highlight me", fontsize=12)
            page.add_highlight_annot(fitz.Rect(50, 50, 200, 72))
            doc.save(source)
            doc.close()

            ok, msg = _process(source, output, {"cleanup_remove_annotations"})

            self.assertTrue(ok, msg)
            doc = fitz.open(output)
            annots = list(doc[0].annots() or [])
            doc.close()
            self.assertEqual(annots, [])
            self.assertIn("已删除PDF注释", msg)

    def test_cleanup_remove_metadata_clears_document_info(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = os.path.join(tmp, "s.pdf")
            output = os.path.join(tmp, "out.pdf")
            doc = fitz.open()
            doc.new_page()
            doc.set_metadata({"title": "Secret", "author": "Someone"})
            doc.save(source)
            doc.close()

            ok, msg = _process(source, output, {"cleanup_remove_metadata"})

            self.assertTrue(ok, msg)
            doc = fitz.open(output)
            meta = doc.metadata or {}
            doc.close()
            self.assertFalse((meta.get("title") or "").strip())
            self.assertFalse((meta.get("author") or "").strip())
            self.assertIn("已删除文档元数据", msg)

    def test_cleanup_remove_attachments_deletes_embedded_files(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = os.path.join(tmp, "s.pdf")
            output = os.path.join(tmp, "out.pdf")
            doc = fitz.open()
            doc.new_page()
            doc.embfile_add("note.txt", b"hello attachment")
            doc.save(source)
            doc.close()

            ok, msg = _process(source, output, {"cleanup_remove_attachments"})

            self.assertTrue(ok, msg)
            doc = fitz.open(output)
            count = doc.embfile_count()
            doc.close()
            self.assertEqual(count, 0)
            self.assertIn("已删除文档附件", msg)

    def test_cleanup_remove_tags_clears_struct_tree_keys(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = os.path.join(tmp, "s.pdf")
            output = os.path.join(tmp, "out.pdf")
            doc = fitz.open()
            doc.new_page()
            doc.xref_set_key(doc.pdf_catalog(), "MarkInfo", "<< /Marked true >>")
            doc.save(source)
            doc.close()

            ok, msg = _process(source, output, {"cleanup_remove_tags"})

            self.assertTrue(ok, msg)
            doc = fitz.open(output)
            kind, _ = doc.xref_get_key(doc.pdf_catalog(), "MarkInfo")
            doc.close()
            self.assertEqual(kind, "null")
            self.assertIn("已删除文档标签", msg)

    def test_cleanup_remove_dynamic_content_nulls_names(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = os.path.join(tmp, "s.pdf")
            output = os.path.join(tmp, "out.pdf")
            doc = fitz.open()
            doc.new_page()
            doc.xref_set_key(doc.pdf_catalog(), "Names", "<< /JavaScript << /Names [] >> >>")
            doc.save(source)
            doc.close()

            ok, msg = _process(source, output, {"cleanup_remove_dynamic_content"})

            self.assertTrue(ok, msg)
            doc = fitz.open(output)
            kind, _ = doc.xref_get_key(doc.pdf_catalog(), "Names")
            doc.close()
            self.assertEqual(kind, "null")
            self.assertIn("已删除动态内容/JavaScript", msg)

    def test_cleanup_remove_all_links_bookmarks_clears_both(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = os.path.join(tmp, "s.pdf")
            output = os.path.join(tmp, "out.pdf")
            doc = fitz.open()
            doc.new_page()
            doc.new_page()
            doc.set_toc([[1, "Chapter", 1]])
            doc[0].insert_link({
                "kind": fitz.LINK_URI,
                "from": fitz.Rect(50, 50, 150, 70),
                "uri": "https://example.com",
            })
            doc[0].insert_link({
                "kind": fitz.LINK_GOTO,
                "from": fitz.Rect(50, 100, 150, 120),
                "page": 1,
                "to": fitz.Point(0, 0),
            })
            doc.save(source)
            doc.close()

            ok, msg = _process(source, output, {"cleanup_remove_all_links_bookmarks"})

            self.assertTrue(ok, msg)
            doc = fitz.open(output)
            toc = doc.get_toc()
            links = doc[0].get_links()
            doc.close()
            self.assertEqual(toc, [])
            self.assertEqual(links, [])
            self.assertIn("已删除全部链接和书签", msg)

    def test_cleanup_remove_external_uri_and_text_black_recolors_blue_text(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = os.path.join(tmp, "s.pdf")
            output = os.path.join(tmp, "out.pdf")
            doc = fitz.open()
            page = doc.new_page()
            page.insert_text(fitz.Point(55, 65), "blue link", fontsize=12, color=(0, 0, 1))
            page.insert_link({
                "kind": fitz.LINK_URI,
                "from": fitz.Rect(50, 50, 200, 72),
                "uri": "https://example.com",
            })
            doc.save(source)
            doc.close()

            ok, msg = _process(source, output, {"cleanup_remove_external_uri_and_text_black"})

            self.assertTrue(ok, msg)
            doc = fitz.open(output)
            links = doc[0].get_links()
            traces = [t for t in doc[0].get_texttrace() if t.get("type") == 0]
            doc.close()
            self.assertEqual(links, [])
            self.assertTrue(traces)
            color = traces[0].get("color", (1, 1, 1))
            self.assertLess(color[2], 0.1, f"文字未变黑: {color}")


class QpdfRewriteRulesTests(unittest.TestCase):
    def test_fast_web_view_linearizes_output(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = os.path.join(tmp, "s.pdf")
            output = os.path.join(tmp, "out.pdf")
            _save_new_pdf(source)

            ok, msg = _process(source, output, {"fast_web_view"})

            self.assertTrue(ok, msg)
            with open(output, "rb") as f:
                head = f.read(4096)
            self.assertIn(b"/Linearized", head)
            self.assertIn("已启用快速网页浏览", msg)

    def test_convert_pdf_version_outputs_pdf_17(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = os.path.join(tmp, "s.pdf")
            output = os.path.join(tmp, "out.pdf")
            _save_new_pdf(source)

            ok, msg = _process(source, output, {"convert_pdf_version"})

            self.assertTrue(ok, msg)
            with open(output, "rb") as f:
                header = f.read(16)
            self.assertIn(b"%PDF-1.7", header)
            self.assertIn("已转换PDF版本", msg)

    def test_remove_pdf_restrictions_decrypts_owner_protected_pdf(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = os.path.join(tmp, "s.pdf")
            output = os.path.join(tmp, "out.pdf")
            doc = fitz.open()
            doc.new_page()
            doc.save(
                source,
                encryption=fitz.PDF_ENCRYPT_AES_256,
                owner_pw="owner-secret",
                permissions=fitz.PDF_PERM_ACCESSIBILITY,
            )
            doc.close()

            probe = fitz.open(source)
            self.assertFalse(probe.needs_pass)
            self.assertTrue(probe.is_encrypted or probe.metadata.get("encryption"))
            probe.close()

            ok, msg = _process(source, output, {"remove_pdf_restrictions"})

            self.assertTrue(ok, msg)
            doc = fitz.open(output)
            encrypted = doc.is_encrypted
            doc.close()
            self.assertFalse(encrypted)
            self.assertIn("已解除PDF权限限制", msg)


class PipelineContractTests(unittest.TestCase):
    def test_encrypted_pdf_requiring_password_is_rejected(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = os.path.join(tmp, "s.pdf")
            output = os.path.join(tmp, "out.pdf")
            doc = fitz.open()
            doc.new_page()
            doc.save(source, encryption=fitz.PDF_ENCRYPT_AES_256, user_pw="open-secret")
            doc.close()

            ok, msg = _process(source, output, {"title_from_filename"})

            self.assertFalse(ok)
            self.assertIn("已加密", msg)

    def test_no_matching_options_copies_file_unchanged(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = os.path.join(tmp, "same.pdf")
            output = os.path.join(tmp, "out.pdf")
            doc = fitz.open()
            doc.new_page()
            doc.set_metadata({"title": "same"})
            doc.save(source)
            doc.close()

            ok, msg = _process(source, output, {"title_from_filename"})

            self.assertTrue(ok, msg)
            with open(source, "rb") as f1, open(output, "rb") as f2:
                self.assertEqual(f1.read(), f2.read())

    def test_smart_mode_skips_rule_without_findings(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = os.path.join(tmp, "same.pdf")
            output = os.path.join(tmp, "out.pdf")
            doc = fitz.open()
            doc.new_page()
            doc.set_metadata({"title": "same"})
            doc.save(source)
            doc.close()

            ok, msg = _process(source, output, {"title_from_filename"}, mode="smart")

            self.assertTrue(ok, msg)
            self.assertIn("智能处理", msg)
            self.assertIn("已跳过未命中规则", msg)
            self.assertTrue(os.path.exists(output))

    def test_smart_mode_applies_rule_with_findings(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = os.path.join(tmp, "目标名.pdf")
            output = os.path.join(tmp, "out.pdf")
            doc = fitz.open()
            doc.new_page()
            doc.set_metadata({"title": "旧标题"})
            doc.save(source)
            doc.close()

            ok, msg = _process(source, output, {"title_from_filename"}, mode="smart")

            self.assertTrue(ok, msg)
            doc = fitz.open(output)
            title = doc.metadata.get("title")
            doc.close()
            self.assertEqual(title, "目标名")
            self.assertIn("预检命中", msg)


if __name__ == "__main__":
    unittest.main()
