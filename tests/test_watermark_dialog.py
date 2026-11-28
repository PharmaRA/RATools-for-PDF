import tempfile
import time
from pathlib import Path
import unittest
from unittest.mock import patch

import fitz
from PySide6.QtCore import QCoreApplication, QEvent, Qt
from PySide6.QtWidgets import QApplication
from shiboken6 import isValid

from ratools_pdf.ui.dialogs.watermark_dialog import WatermarkDialog


class WatermarkDialogTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.source = Path(self.temp.name) / 'source.pdf'
        self.output = Path(self.temp.name) / 'output.pdf'
        with fitz.open() as doc:
            doc.new_page().insert_text((30, 80), 'DRAFT')
            doc.save(self.source)
        self.dialog = WatermarkDialog(str(self.source))
        self.dialog.show()

    def tearDown(self):
        self.wait_idle()
        self.dialog.close()
        self.dialog.deleteLater()
        # Do not drain deferred destruction belonging to earlier test modules.
        QCoreApplication.sendPostedEvents(self.dialog, QEvent.DeferredDelete)
        self.assertFalse(isValid(self.dialog))
        self.temp.cleanup()

    def wait_idle(self):
        deadline = time.monotonic() + 10
        while self.dialog.worker is not None or self.dialog.preview_timer.isActive():
            self.app.processEvents()
            if time.monotonic() > deadline:
                self.fail('watermark task did not finish')
            time.sleep(.005)
        self.app.processEvents()

    def test_analyze_select_preview_confirm_and_save(self):
        self.dialog.btn_analyze.click()
        self.wait_idle()
        self.assertEqual(self.dialog.objects.topLevelItemCount(), 1)
        item = self.dialog.objects.topLevelItem(0)
        self.assertEqual(item.checkState(0), Qt.Unchecked)
        self.assertFalse(self.dialog.btn_save.isEnabled())
        item.setCheckState(0, Qt.Checked)
        self.wait_idle()
        self.assertNotEqual(self.dialog.before.pixmap().toImage(), self.dialog.after.pixmap().toImage())
        self.dialog.confirm.setChecked(True)
        self.assertTrue(self.dialog.btn_save.isEnabled())
        with patch('ratools_pdf.ui.dialogs.watermark_dialog.QFileDialog.getSaveFileName',
                   return_value=(str(self.output), 'PDF')):
            self.dialog.btn_save.click()
        self.wait_idle()
        self.assertTrue(self.output.exists(), self.dialog.lbl_status.text())
        with fitz.open(self.output) as result:
            self.assertNotIn('DRAFT', result[0].get_text())

    def test_range_change_invalidates_selection_preview_and_confirmation(self):
        self.dialog.btn_analyze.click()
        self.wait_idle()
        self.dialog.objects.topLevelItem(0).setCheckState(0, Qt.Checked)
        self.wait_idle()
        self.dialog.confirm.setChecked(True)
        self.dialog.txt_pages.setText('1')
        self.assertIsNone(self.dialog.analysis)
        self.assertFalse(self.dialog.confirm.isChecked())
        self.assertFalse(self.dialog.btn_save.isEnabled())
        self.assertEqual(self.dialog.objects.topLevelItemCount(), 0)

    def test_close_waits_for_running_task(self):
        import threading
        started, release = threading.Event(), threading.Event()

        def action(cancel=None):
            started.set()
            release.wait(5)

        self.dialog._start(action, (), lambda _: None, 'busy')
        try:
            self.assertTrue(started.wait(2))
            self.dialog.reject()
            self.assertTrue(self.dialog.isVisible())
            self.assertFalse(self.dialog.btn_save.isEnabled())
        finally:
            release.set()
        self.wait_idle()
        self.dialog.reject()
        self.assertFalse(self.dialog.isVisible())

    def test_each_selected_page_requires_preview_of_current_selection(self):
        with fitz.open() as doc:
            doc.new_page().insert_text((30, 80), 'DRAFT ONE')
            doc.new_page().insert_text((30, 80), 'DRAFT TWO')
            doc.save(self.source)
        self.dialog.btn_analyze.click()
        self.wait_idle()
        # Two selections before the debounce fires: only the last page renders.
        self.dialog.objects.topLevelItem(0).setCheckState(0, Qt.Checked)
        self.dialog.objects.topLevelItem(1).setCheckState(0, Qt.Checked)
        self.wait_idle()
        self.dialog.confirm.setChecked(True)
        self.assertFalse(self.dialog.btn_save.isEnabled())
        self.dialog.page_number.setValue(1)
        self.wait_idle()
        self.assertTrue(self.dialog.btn_save.isEnabled())
        self.dialog.objects.topLevelItem(0).setCheckState(0, Qt.Unchecked)
        self.assertFalse(self.dialog.confirm.isChecked())

    def test_main_window_entry_passes_selected_pdf(self):
        from types import SimpleNamespace
        from unittest.mock import Mock
        from ratools_pdf.ui.main_window import MainWindow

        selected = Mock()
        selected.text.return_value = str(self.source)
        view = SimpleNamespace(tree=Mock())
        view.tree.selectedItems.return_value = [selected]
        with patch('ratools_pdf.ui.dialogs.watermark_dialog.WatermarkDialog') as dialog:
            MainWindow.show_watermark_dialog(view)
        dialog.assert_called_once_with(initial_file=str(self.source), parent=view)
        dialog.return_value.exec.assert_called_once()
        dialog.return_value.deleteLater.assert_called_once()

    def test_real_main_window_exposes_watermark_button(self):
        from ratools_pdf.ui.main_window import MainWindow

        window = MainWindow(settings_path=str(Path(self.temp.name) / 'settings.ini'))
        try:
            window.show()
            index = next(i for i, module in enumerate(window.MODULES_DATA)
                         if module['title'] == '内容合规与安全性')
            window.switch_settings_page(index)
            self.app.processEvents()
            self.assertTrue(window.btn_remove_watermark.isVisible())
            self.assertTrue(window.btn_remove_watermark.isEnabled())
            with patch('ratools_pdf.ui.dialogs.watermark_dialog.WatermarkDialog') as dialog:
                window.btn_remove_watermark.click()
            dialog.return_value.exec.assert_called_once()
        finally:
            window.close()
            window.deleteLater()
            QCoreApplication.sendPostedEvents(window, QEvent.DeferredDelete)
            self.assertFalse(isValid(window))

    def test_dialog_uses_fluent_theme_roles_without_local_stylesheets(self):
        from PySide6.QtWidgets import QPushButton, QWidget

        for btn in (self.dialog.btn_source, self.dialog.btn_preview,
                    self.dialog.btn_cancel, self.dialog.btn_close_dialog):
            self.assertEqual('dialogSecondaryBtn', btn.objectName())
            self.assertEqual(Qt.PointingHandCursor, btn.cursor().shape())
        for btn in (self.dialog.btn_analyze, self.dialog.btn_save):
            self.assertEqual('dialogPrimaryBtn', btn.objectName())
            self.assertEqual(Qt.PointingHandCursor, btn.cursor().shape())

        self.assertEqual('', self.dialog.styleSheet())
        for widget in self.dialog.findChildren(QWidget):
            self.assertEqual('', widget.styleSheet(), widget.objectName())
        self.assertGreaterEqual(len(self.dialog.findChildren(QPushButton)), 6)

    def test_notes_visibility_summary_and_confirm_card_sync(self):
        self.assertFalse(self.dialog.lbl_notes.isVisible())
        self.assertEqual('未分析', self.dialog.lbl_objects_summary.text())
        self.assertEqual('false', self.dialog.confirm_card.property('checked'))

        self.dialog.btn_analyze.click()
        self.wait_idle()
        self.assertEqual('共 1 项 · 已选 0 项', self.dialog.lbl_objects_summary.text())

        item = self.dialog.objects.topLevelItem(0)
        item.setCheckState(0, Qt.Checked)
        self.wait_idle()
        self.assertEqual('共 1 项 · 已选 1 项', self.dialog.lbl_objects_summary.text())

        self.dialog.confirm.setChecked(True)
        self.assertEqual('true', self.dialog.confirm_card.property('checked'))

        self.dialog.txt_pages.setText('1')
        self.assertEqual('未分析', self.dialog.lbl_objects_summary.text())
        self.assertFalse(self.dialog.lbl_notes.isVisible())
        self.assertEqual('false', self.dialog.confirm_card.property('checked'))

    def test_notes_panel_becomes_visible_when_analysis_has_notes(self):
        with fitz.open() as doc:
            page = doc.new_page()
            page.insert_text((30, 80), 'DRAFT')
            page.add_text_annot((50, 120), 'Note')
            doc.save(self.source)
        with patch('ratools_pdf.ui.dialogs.watermark_dialog.watermarks.analyze') as mock_analyze:
            from ratools_pdf.pdf.watermarks import Analysis
            mock_analyze.return_value = Analysis(
                path=str(self.source),
                source=self.source.read_bytes(),
                digest='abc',
                page_count=1,
                pages=(0,),
                candidates=(),
                notes=('部分复杂绘图对象已跳过',),
            )
            self.dialog.btn_analyze.click()
            self.wait_idle()
        self.assertTrue(self.dialog.lbl_notes.isVisible())
        self.assertIn('部分复杂绘图对象已跳过', self.dialog.lbl_notes.text())

    def test_preview_pixmap_adapts_to_scroll_viewport_width(self):
        self.dialog.resize(1040, 680)
        self.app.processEvents()
        self.dialog.btn_analyze.click()
        self.wait_idle()
        self.dialog.objects.topLevelItem(0).setCheckState(0, Qt.Checked)
        self.wait_idle()
        before_pixmap = self.dialog.before.pixmap()
        after_pixmap = self.dialog.after.pixmap()
        self.assertFalse(before_pixmap.isNull())
        self.assertFalse(after_pixmap.isNull())
        self.assertLessEqual(before_pixmap.width(), max(320, self.dialog.before_scroll.viewport().width()))
        self.assertLessEqual(after_pixmap.width(), max(320, self.dialog.after_scroll.viewport().width()))

    def test_objects_tree_shows_readable_labels_and_likely_watermark_badge(self):
        with fitz.open() as doc:
            p = doc.new_page()
            p.insert_text((30, 80), 'BODY TEXT', fontsize=12)
            p.insert_text((30, 140), 'DRAFT', fontsize=24, fill_opacity=0.4)
            doc.save(self.source)
        self.dialog.btn_analyze.click()
        self.wait_idle()
        self.assertEqual(self.dialog.objects.topLevelItemCount(), 2)
        body_item = self.dialog.objects.topLevelItem(0)
        draft_item = self.dialog.objects.topLevelItem(1)
        self.assertEqual('文字', body_item.text(2))
        self.assertIn('BODY TEXT', body_item.text(3))
        self.assertEqual('文字 · 疑似', draft_item.text(2))
        self.assertIn('DRAFT', draft_item.text(3))

    def test_filter_controls_hide_non_matching_items(self):
        with fitz.open() as doc:
            p1 = doc.new_page()
            p1.insert_text((30, 80), 'BODY ONE', fontsize=12)
            p1.insert_text((30, 140), 'DRAFT', fontsize=24, fill_opacity=0.4)
            p2 = doc.new_page()
            p2.insert_text((30, 80), 'BODY TWO', fontsize=12)
            doc.save(self.source)
        self.dialog.btn_analyze.click()
        self.wait_idle()
        self.assertEqual(self.dialog.objects.topLevelItemCount(), 3)

        self.dialog.txt_filter.setText('DRAFT')
        hidden = [self.dialog.objects.topLevelItem(i).isHidden() for i in range(3)]
        self.assertEqual(hidden, [True, False, True])

        self.dialog.txt_filter.clear()
        self.dialog.combo_kind.setCurrentIndex(self.dialog.combo_kind.findData('likely'))
        hidden = [self.dialog.objects.topLevelItem(i).isHidden() for i in range(3)]
        self.assertEqual(hidden, [True, False, True])

        self.dialog.combo_kind.setCurrentIndex(0)
        self.dialog.page_number.setValue(2)
        self.dialog.cb_current_page_only.setChecked(True)
        hidden = [self.dialog.objects.topLevelItem(i).isHidden() for i in range(3)]
        self.assertEqual(hidden, [True, True, False])

    def test_batch_select_similar_across_pages_requires_representative_previews(self):
        with fitz.open() as doc:
            for idx in range(4):
                page = doc.new_page()
                page.insert_text((30, 80), f'BODY {idx + 1}', fontsize=12)
                page.insert_text((30, 140), 'CONFIDENTIAL', fontsize=24, fill_opacity=0.4)
            doc.save(self.source)
        self.dialog.btn_analyze.click()
        self.wait_idle()

        # 聚焦第 1 页的水印对象并一键勾选同款跨页
        draft_item = self.dialog.objects.topLevelItem(1)
        self.dialog.objects.setCurrentItem(draft_item)
        self.assertTrue(self.dialog.btn_select_similar.isEnabled())
        self.dialog.btn_select_similar.click()
        self.wait_idle()
        self.assertEqual(len(self.dialog.selected_ids()), 4)

        # 此时第 1 页已预览，只需再预览末页（第 4 页）即可启用保存
        self.dialog.confirm.setChecked(True)
        self.assertFalse(self.dialog.btn_save.isEnabled())
        self.dialog.page_number.setValue(4)
        self.wait_idle()
        self.assertTrue(self.dialog.btn_save.isEnabled())

        # 清空勾选后恢复未选状态
        self.dialog.btn_unselect_all.click()
        self.wait_idle()
        self.assertEqual(len(self.dialog.selected_ids()), 0)
        self.assertFalse(self.dialog.btn_save.isEnabled())

    def test_preview_fit_mode_scroll_sync_cache_and_focus_highlight(self):
        with fitz.open() as doc:
            page = doc.new_page(width=595, height=842)
            page.insert_text((60, 120), 'BODY TEXT', fontsize=12)
            page.insert_text((160, 420), 'DRAFT WATERMARK', fontsize=28, fill_opacity=0.4)
            doc.save(self.source)

        self.dialog.resize(1040, 680)
        self.app.processEvents()
        self.dialog.btn_analyze.click()
        self.wait_idle()
        self.assertIn(0, self.dialog._before_cache)
        self.assertTrue(self.dialog.btn_fit_mode.isChecked())

        # 切换到适应宽度模式，宽度应增大
        fit_page_w = self.dialog.before.pixmap().width()
        self.dialog.btn_fit_mode.setChecked(False)
        self.app.processEvents()
        fit_width_w = self.dialog.before.pixmap().width()
        self.assertGreaterEqual(fit_width_w, fit_page_w)

        # 验证同步滚动（垂直滚动条若有可滚动范围应同步）
        v_before = self.dialog.before_scroll.verticalScrollBar()
        v_after = self.dialog.after_scroll.verticalScrollBar()
        if v_before.maximum() > 0:
            v_before.setValue(min(40, v_before.maximum()))
            self.app.processEvents()
            self.assertEqual(v_after.value(), v_before.value())

        # 聚焦不同对象时，原始页高亮框应变化
        self.dialog.objects.setCurrentItem(self.dialog.objects.topLevelItem(0))
        self.app.processEvents()
        img_focus_0 = self.dialog.before.pixmap().toImage()
        self.dialog.objects.setCurrentItem(self.dialog.objects.topLevelItem(1))
        self.app.processEvents()
        img_focus_1 = self.dialog.before.pixmap().toImage()
        self.assertNotEqual(img_focus_0, img_focus_1)

    def test_open_output_dir_button_after_save_and_empty_candidates_guidance(self):
        with fitz.open() as doc:
            doc.new_page()
            doc.save(self.source)
        self.dialog.btn_analyze.click()
        self.wait_idle()
        self.assertIn('未找到可独立移除的对象', self.dialog.lbl_status.text())
        self.assertFalse(self.dialog.btn_open_output.isVisible())

        with fitz.open() as doc:
            doc.new_page().insert_text((30, 80), 'DRAFT')
            doc.save(self.source)
        self.dialog.btn_analyze.click()
        self.wait_idle()
        self.dialog.objects.topLevelItem(0).setCheckState(0, Qt.Checked)
        self.wait_idle()
        self.dialog.confirm.setChecked(True)
        with patch('ratools_pdf.ui.dialogs.watermark_dialog.QFileDialog.getSaveFileName',
                   return_value=(str(self.output), 'PDF')):
            self.dialog.btn_save.click()
        self.wait_idle()
        self.assertTrue(self.dialog.btn_open_output.isVisible())
        with patch('ratools_pdf.ui.dialogs.watermark_dialog.os.startfile') as startfile:
            self.dialog.btn_open_output.click()
        startfile.assert_called_once_with(str(self.output.parent))

    def test_drag_and_drop_pdf_triggers_analysis(self):
        from PySide6.QtCore import QMimeData, QPointF, QUrl
        from PySide6.QtGui import QDropEvent

        dropped = Path(self.temp.name) / 'dropped.pdf'
        with fitz.open() as doc:
            doc.new_page().insert_text((30, 80), 'DROPPED WATERMARK')
            doc.save(dropped)

        mime = QMimeData()
        mime.setUrls([QUrl.fromLocalFile(str(dropped))])
        event = QDropEvent(
            QPointF(20, 20), Qt.CopyAction, mime, Qt.LeftButton, Qt.NoModifier
        )
        self.dialog.dropEvent(event)
        self.wait_idle()
        self.assertEqual(self.dialog.txt_source.text(), str(dropped))
        self.assertEqual(self.dialog.objects.topLevelItemCount(), 1)
        self.assertIn('DROPPED WATERMARK', self.dialog.objects.topLevelItem(0).text(3))
