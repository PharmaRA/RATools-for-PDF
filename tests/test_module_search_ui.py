"""功能模块与规则搜索界面的 Qt 集成测试。"""

import gc
import os
import tempfile
import unittest
from unittest.mock import patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import QCoreApplication, QEvent, Qt
from PySide6.QtTest import QTest
from PySide6.QtWidgets import QApplication, QPushButton

from ratools_pdf.controllers.main_controller import MainController
from ratools_pdf.ui.main_window import MainWindow
from ratools_pdf.ui.widgets import RuleOptionCard


class ModuleSearchUITests(unittest.TestCase):
    def setUp(self):
        self.app = QApplication.instance() or QApplication([])
        self._previous_app_stylesheet = self.app.styleSheet()
        self._previous_app_font = self.app.font()
        self.app.setStyleSheet("")
        self._windows = []
        self._controllers = []

    def tearDown(self):
        for controller in self._controllers:
            controller.shutdown()
        self._controllers.clear()
        for window in self._windows:
            window.close()
            window.deleteLater()
        self._windows.clear()
        QCoreApplication.sendPostedEvents(None, QEvent.DeferredDelete)
        self.app.processEvents()
        gc.collect()
        self.app.setStyleSheet(self._previous_app_stylesheet)
        self.app.setFont(self._previous_app_font)

    def _create_window(self, *, with_controller=False):
        with tempfile.TemporaryDirectory() as tmp_dir:
            with patch("ratools_pdf.ui.main_window.get_app_dir", return_value=tmp_dir):
                window = MainWindow()
                self._windows.append(window)
                if with_controller:
                    controller = MainController(window)
                    self._controllers.append(controller)
                return window

    def test_search_filters_left_nav_and_displays_matching_result_cards(self):
        window = self._create_window()
        window.show()
        self.app.processEvents()

        window.module_search_input.setText("书签")
        self.app.processEvents()

        visible_titles = [
            btn.text() for btn in window.nav_buttons if not btn.isHidden()
        ]
        self.assertIn("书签管理", visible_titles)
        self.assertIn("初始视图与文档属性", visible_titles)
        self.assertNotIn("文件级优化与输出", visible_titles)
        self.assertFalse(window.search_results_page.isHidden())
        self.assertTrue(all(page.isHidden() for page in window.settings_pages))
        self.assertEqual("搜索结果", window.rh_title.text())
        self.assertFalse(window.btn_clear_search.isHidden())

    def test_toggling_rule_in_search_results_syncs_with_real_checkbox_and_exclusivity(self):
        window = self._create_window(with_controller=True)
        window.all_checkboxes["page_size_a4"].setChecked(False)
        window.all_checkboxes["page_size_letter"].setChecked(True)

        window.module_search_input.setText("尺寸")
        self.app.processEvents()

        cards = window.search_results_page.findChildren(RuleOptionCard)
        a4_card = next(c for c in cards if "A4" in c.title_label.text())
        letter_card = next(c for c in cards if "Letter" in c.title_label.text())

        self.assertFalse(a4_card.checkbox.isChecked())
        self.assertTrue(letter_card.checkbox.isChecked())

        a4_card.checkbox.setChecked(True)
        self.app.processEvents()

        self.assertTrue(window.all_checkboxes["page_size_a4"].isChecked())
        self.assertFalse(window.all_checkboxes["page_size_letter"].isChecked())
        self.assertFalse(letter_card.checkbox.isChecked())
        self.assertIn("已选", window.selection_summary_label.text())

    def test_lossless_rule_search_jumps_to_optimization_module_without_breaking_group(self):
        window = self._create_window()
        window.show()
        self.app.processEvents()

        window.module_search_input.setText("对象流")
        self.app.processEvents()

        action_btns = [
            btn for btn in window.search_results_page.findChildren(QPushButton)
            if btn.text() == "前往配置"
        ]
        self.assertGreaterEqual(len(action_btns), 1)
        action_btns[0].click()
        self.app.processEvents()

        self.assertEqual("", window.module_search_input.text())
        self.assertTrue(window.search_results_page.isHidden())
        opt_idx = next(
            i for i, mod in enumerate(window.MODULES_DATA)
            if mod["title"] == "文件级优化与输出"
        )
        self.assertEqual(opt_idx, window.current_module_index)
        self.assertTrue(window.nav_buttons[opt_idx].isChecked())
        self.assertFalse(window.settings_pages[opt_idx].isHidden())

    def test_clicking_nav_button_escape_empty_search_and_theme_switch(self):
        window = self._create_window()
        window.show()
        self.app.processEvents()

        # 1. 无匹配查询展示空态并可通过按钮恢复
        window.module_search_input.setText("不存在的规则xyz999")
        self.app.processEvents()
        self.assertFalse(window.module_search_empty_label.isHidden())
        reset_btn = next(
            btn for btn in window.search_results_page.findChildren(QPushButton)
            if btn.text() == "清空搜索条件"
        )
        reset_btn.click()
        self.app.processEvents()
        self.assertEqual("", window.module_search_input.text())
        self.assertTrue(all(not btn.isHidden() for btn in window.nav_buttons))

        # 2. 按 Escape 清空搜索
        window.module_search_input.setText("水印")
        self.app.processEvents()
        QTest.keyClick(window.module_search_input, Qt.Key.Key_Escape)
        self.app.processEvents()
        self.assertEqual("", window.module_search_input.text())
        self.assertTrue(window.search_results_page.isHidden())

        # 3. 搜索期间切换主题不崩溃且保持结果页可见；点击左侧可见模块可直接跳转并清空搜索
        window.module_search_input.setText("书签")
        self.app.processEvents()
        window.on_theme_mode_changed("dark")
        self.app.processEvents()
        self.assertFalse(window.search_results_page.isHidden())

        bm_idx = next(
            i for i, mod in enumerate(window.MODULES_DATA)
            if mod["title"] == "书签管理"
        )
        window.nav_buttons[bm_idx].click()
        self.app.processEvents()
        self.assertEqual("", window.module_search_input.text())
        self.assertEqual(bm_idx, window.current_module_index)
        self.assertTrue(window.search_results_page.isHidden())


if __name__ == "__main__":
    unittest.main()
