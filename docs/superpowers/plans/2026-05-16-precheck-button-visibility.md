# Precheck Button Visibility Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Hide the `预检` button while formal processing is active, then restore it when processing ends or fails.

**Architecture:** Keep the visibility rule in `controller.py`, where processing state already changes. Use the existing `MainWindow` widget instance for the actual `hide()` / `show()` calls so the footer layout does not need to change. Add one regression test that exercises the controller lifecycle directly and checks button visibility across start and finish.

**Tech Stack:** Python, PySide6, unittest, unittest.mock, pytest

---

### Task 1: Add a regression test for precheck button visibility

**Files:**
- Modify: `tests/test_regressions.py`

- [ ] **Step 1: Write the failing test**

```python
    def test_precheck_button_hides_during_processing_and_returns_after_finish(self):
        window = MainWindow()
        controller = MainController(window)
        try:
            controller.loaded_files = [os.path.join(tempfile.gettempdir(), "sample.pdf")]
            controller.view.get_selected_options = lambda: {"page_size_a4"}
            fake_worker = MagicMock()
            with patch("controller.QFileDialog.getExistingDirectory", return_value=tempfile.gettempdir()), patch(
                "controller.ProcessWorker",
                return_value=fake_worker,
            ), patch.object(controller.view, "show_success_message"):
                controller.start_processing()

            self.assertFalse(window.btn_precheck.isHidden())
            self.assertTrue(window.btn_precheck.isHidden())

            controller.processing_finished("已完成")

            self.assertFalse(window.btn_precheck.isHidden())
        finally:
            window.close()

    def test_precheck_button_returns_after_processing_error(self):
        window = MainWindow()
        controller = MainController(window)
        try:
            controller.loaded_files = [os.path.join(tempfile.gettempdir(), "sample.pdf")]
            controller.view.get_selected_options = lambda: {"page_size_a4"}
            fake_worker = MagicMock()
            with patch("controller.QFileDialog.getExistingDirectory", return_value=tempfile.gettempdir()), patch(
                "controller.ProcessWorker",
                return_value=fake_worker,
            ), patch.object(controller.view, "show_error_message"):
                controller.start_processing()

            self.assertTrue(window.btn_precheck.isHidden())

            controller.processing_error("boom")

            self.assertFalse(window.btn_precheck.isHidden())
        finally:
            window.close()
```

- [ ] **Step 2: Run test to verify it fails**

Run: `pytest tests/test_regressions.py::RegressionTests::test_precheck_button_hides_during_processing_and_returns_after_finish -v tests/test_regressions.py::RegressionTests::test_precheck_button_returns_after_processing_error -v`
Expected: FAIL because `btn_precheck` is still hidden state is not changed by the controller yet.

- [ ] **Step 3: Write minimal implementation**

Do not implement yet.

- [ ] **Step 4: Run test to verify it passes**

Run: `pytest tests/test_regressions.py::RegressionTests::test_precheck_button_hides_during_processing_and_returns_after_finish -v tests/test_regressions.py::RegressionTests::test_precheck_button_returns_after_processing_error -v`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add tests/test_regressions.py
git commit -m "test: cover precheck button visibility during processing"
```

### Task 2: Hide and restore the precheck button in the controller

**Files:**
- Modify: `controller.py:1045-1328`

- [ ] **Step 1: Write the failing test**

Use the test from Task 1 and confirm it fails before implementation.

- [ ] **Step 2: Run test to verify it fails**

Run: `pytest tests/test_regressions.py::RegressionTests::test_precheck_button_hides_during_processing_and_returns_after_finish -v`
Expected: FAIL because the visibility change is missing.

- [ ] **Step 3: Write minimal implementation**

```python
    def start_processing(self):
        if self.worker and self.worker.isRunning():
            self.worker.request_stop()
            self.view.btn_start.setEnabled(False)
            self.view.btn_start.setText("正在停止...")
            self.view.btn_skip_current.setEnabled(False)
            return

        if self._is_precheck_running():
            self.view.show_warning_message("⚠️ 正在预检", "请等待当前预检完成后再开始处理。")
            return

        if not self.loaded_files:
            self.view.show_warning_message("⚠️ 警告", "请至少添加一个 PDF 文件！")
            return

        selected_options = self.view.get_selected_options()
        if not selected_options:
            self.view.show_warning_message("⚠️ 警告", "请至少在右侧勾选一个处理规则！")
            return

        if "embed_nonstandard_fonts" in selected_options:
            self.view.show_warning_message(
                "⚠️ 功能暂不可用",
                "【嵌入全部非标准字体】当前版本暂不可用。"
            )
            return

        if "filename_ectd_format" in selected_options:
            rename_pairs = []
            for i, file_path in enumerate(self.loaded_files, start=1):
                base_name = os.path.basename(file_path)
                name, ext = os.path.splitext(base_name)
                normalized = name.lower().replace(" ", "-")
                normalized = re.sub(r'[^a-z0-9_-]', '', normalized)
                if not normalized:
                    normalized = f"doc_{i:03d}"
                new_name = f"{normalized}{ext.lower()}"
                if new_name != base_name:
                    rename_pairs.append((base_name, new_name))

            if rename_pairs:
                details = "\n".join([
                    f"{idx:>2}. {old}\n    -> {new}"
                    for idx, (old, new) in enumerate(rename_pairs, start=1)
                ])
                msg = (
                    "已启用【eCTD 文件名合规格式化】。\n"
                    "以下文件在输出时将被重命名：\n\n"
                    f"{details}\n\n"
                    "确认后继续处理。"
                )
                if not self.view.show_confirm_message("📝 确认文件名格式化", msg):
                    return

        overwrite_cb = self.view.all_checkboxes.get("覆盖原始文件 (不推荐)")
        overwrite_original = overwrite_cb.isChecked() if overwrite_cb else False
        processing_files = list(self.loaded_files)

        out_dir = ""
        common_base = ""

        if overwrite_original:
            if not self.view.show_confirm_message("⚠️ 危险操作确认",
                                                  "您勾选了【覆盖原始文件】。\n此操作不可逆，强烈建议您在操作前备份文件！\n\n是否继续？"):
                return
        else:
            default_output_dir = self.view.settings_dialog.default_output_edit.text().strip()
            start_dir = default_output_dir if default_output_dir and os.path.isdir(default_output_dir) else os.path.expanduser("~")

            user_selected_dir = QFileDialog.getExistingDirectory(
                self.view,
                "选择输出文件保存的根目录",
                start_dir
            )
            if not user_selected_dir:
                return

            out_dir = os.path.join(user_selected_dir, "RATools_Output")
            self.last_output_dir = out_dir

            try:
                dirs = [os.path.dirname(os.path.abspath(f)) for f in processing_files]
                common_base = os.path.commonpath(dirs)
            except ValueError:
                common_base = ""

        self.view.btn_start.setEnabled(True)
        self.view.btn_start.setText("■ 停止处理")
        self.view.btn_start.setProperty("stopMode", True)
        self.view.btn_precheck.setEnabled(False)
        self.view.btn_precheck.hide()
        self.view.btn_skip_current.show()
        self.view.btn_skip_current.setEnabled(True)
        self.view.style().unpolish(self.view.btn_start)
        self.view.style().polish(self.view.btn_start)

        self.processing_started_at = datetime.now()
        self.processing_files = processing_files
        self.processing_total = len(processing_files)
        self.processing_done = 0
        self.processing_done_paths = set()
        self.processing_current_file = ""
        self._last_processing_hint = ""
        self._refresh_processing_hint()
        self.processing_timer.start()

        self.worker = ProcessWorker(processing_files, selected_options, out_dir, common_base, overwrite_original)
        self.worker.progress.connect(self.update_progress)
        self.worker.finished_all.connect(self.processing_finished)
        self.worker.error.connect(self.processing_error)
        self.worker.start()

    def processing_finished(self, summary):
        self.process_logs += f"\n{'=' * 56}\n批量处理结束\n{summary}\n{'=' * 56}\n"
        self.processing_timer.stop()
        self.view.processing_hint_label.setText("")
        self.view.btn_start.setEnabled(True)
        self.view.btn_start.setText("▶ 开始批量处理")
        self.view.btn_start.setProperty("stopMode", False)
        self.view.btn_precheck.show()
        self.view.btn_skip_current.setEnabled(False)
        self.view.btn_skip_current.hide()
        self.view.style().unpolish(self.view.btn_start)
        self.view.style().polish(self.view.btn_start)
        self.view.refresh_selection_summary()
        self.processing_started_at = None
        self.processing_total = 0
        self.processing_done = 0
        self.processing_done_paths.clear()
        self.processing_files = []
        self.processing_current_file = ""
        self._last_processing_hint = ""

    def processing_error(self, error_msg):
        self.process_logs += f"\n{'!' * 56}\n[致命错误] {error_msg}\n{'!' * 56}\n"
        self.processing_timer.stop()
        self.view.processing_hint_label.setText("")
        self.view.btn_start.setEnabled(True)
        self.view.btn_start.setText("▶ 开始批量处理")
        self.view.btn_start.setProperty("stopMode", False)
        self.view.btn_precheck.show()
        self.view.btn_skip_current.setEnabled(False)
        self.view.btn_skip_current.hide()
        self.view.style().unpolish(self.view.btn_start)
        self.view.style().polish(self.view.btn_start)
        self.view.refresh_selection_summary()
        self.processing_started_at = None
        self.processing_total = 0
        self.processing_done = 0
        self.processing_done_paths.clear()
        self.processing_files = []
        self.processing_current_file = ""
        self._last_processing_hint = ""
        self.view.show_error_message("❌ 处理异常", f"处理过程中发生错误：\n{error_msg}")
```

- [ ] **Step 4: Run test to verify it passes**

Run: `pytest tests/test_regressions.py::RegressionTests::test_precheck_button_hides_during_processing_and_returns_after_finish -v`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add controller.py tests/test_regressions.py
git commit -m "fix: hide precheck button during processing"
```
