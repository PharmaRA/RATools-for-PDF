# Precheck Export Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add CSV export for the latest batch precheck results.

**Architecture:** `PreCheckWorker` emits structured row data while keeping the existing log/progress behavior. `MainController` stores the latest precheck rows and writes them as UTF-8 BOM CSV through a new footer button in `MainWindow`.

**Tech Stack:** Python, PySide6, csv, unittest/pytest.

---

## File Structure

- Modify `controller.py`: add structured precheck result signal handling, CSV export method, and button wiring.
- Modify `view.py`: add `btn_export_precheck` in the footer and keep its enabled state in `refresh_selection_summary`.
- Modify `tests/test_regressions.py`: add regression tests for worker result emission, controller state/button enablement, and CSV content.

### Task 1: Worker Emits Structured Precheck Results

**Files:**
- Modify: `controller.py`
- Test: `tests/test_regressions.py`

- [ ] **Step 1: Write the failing test**

Add this test near the existing precheck worker tests in `tests/test_regressions.py`:

```python
def test_precheck_worker_emits_structured_result_rows(self):
    from controller import PreCheckWorker

    pdf_path = os.path.join(tempfile.gettempdir(), "worker-result.pdf")
    worker = PreCheckWorker([pdf_path])
    rows = []
    worker.result_ready.connect(rows.append)

    with patch.object(
        PDFProcessor,
        "build_precheck_report",
        return_value={
            "available": True,
            "suggestions": {
                "title_from_filename": {
                    "title": "同步文件名为标题",
                    "reason": "PDF标题属性与文件名不一致或为空",
                },
            },
        },
    ):
        worker.run()

    self.assertEqual(len(rows), 1)
    self.assertEqual(rows[0]["file_path"], pdf_path)
    self.assertEqual(rows[0]["file_name"], "worker-result.pdf")
    self.assertEqual(rows[0]["status"], "建议处理")
    self.assertEqual(rows[0]["suggestions"], "同步文件名为标题")
    self.assertEqual(rows[0]["error"], "")
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_regressions.py::RegressionTests::test_precheck_worker_emits_structured_result_rows -v`

Expected: FAIL because `PreCheckWorker` does not have `result_ready`.

- [ ] **Step 3: Write minimal implementation**

In `controller.py`, add a signal to `PreCheckWorker`:

```python
result_ready = Signal(dict)
```

Emit rows in `PreCheckWorker.run()` after each report is built:

```python
self.result_ready.emit({
    "file_name": base_name,
    "file_path": file_path,
    "status": "预检失败",
    "suggestions": "",
    "error": reason,
})
```

For suggested files:

```python
self.result_ready.emit({
    "file_name": base_name,
    "file_path": file_path,
    "status": "建议处理",
    "suggestions": advice,
    "error": "",
})
```

For clean files:

```python
self.result_ready.emit({
    "file_name": base_name,
    "file_path": file_path,
    "status": "无需处理",
    "suggestions": "",
    "error": "",
})
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest tests/test_regressions.py::RegressionTests::test_precheck_worker_emits_structured_result_rows -v`

Expected: PASS.

### Task 2: Store Latest Precheck Results And Enable Export Button

**Files:**
- Modify: `controller.py`
- Modify: `view.py`
- Test: `tests/test_regressions.py`

- [ ] **Step 1: Write the failing test**

Add this test to `tests/test_regressions.py`:

```python
def test_precheck_results_are_stored_and_export_button_enables(self):
    window = MainWindow()
    controller = MainController(window)
    try:
        self.assertTrue(hasattr(window, "btn_export_precheck"))
        self.assertFalse(window.btn_export_precheck.isEnabled())

        controller._record_precheck_result({
            "file_name": "sample.pdf",
            "file_path": os.path.join(tempfile.gettempdir(), "sample.pdf"),
            "status": "无需处理",
            "suggestions": "",
            "error": "",
        })

        self.assertEqual(len(controller.last_precheck_results), 1)
        self.assertTrue(window.btn_export_precheck.isEnabled())
    finally:
        window.close()
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_regressions.py::RegressionTests::test_precheck_results_are_stored_and_export_button_enables -v`

Expected: FAIL because the button and controller storage do not exist.

- [ ] **Step 3: Write minimal implementation**

In `view.py`, create `btn_export_precheck` next to the precheck button:

```python
self.btn_export_precheck = QPushButton("⬇ 导出预检结果")
self.btn_export_precheck.setObjectName("actionBtn")
self.btn_export_precheck.setEnabled(False)
```

Add it to the footer before `btn_precheck`:

```python
footer_layout.addWidget(self.btn_export_precheck)
footer_layout.addSpacing(10)
```

In `view.py` `refresh_selection_summary`, disable it during processing/prechecking and otherwise enable it only when it has results:

```python
has_precheck_results = self.btn_export_precheck.property("hasPrecheckResults") is True
self.btn_export_precheck.setEnabled(has_precheck_results and not is_prechecking)
self.btn_export_precheck.setToolTip("导出最近一次批量预检结果" if has_precheck_results else "请先执行一次批量预检")
```

In `controller.py` `MainController.__init__`, add:

```python
self.last_precheck_results = []
```

Wire the button:

```python
self.view.btn_export_precheck.clicked.connect(self.export_precheck_results)
```

Add a recorder method:

```python
def _record_precheck_result(self, row):
    self.last_precheck_results.append(dict(row))
    self.view.btn_export_precheck.setProperty("hasPrecheckResults", True)
    self.view.refresh_selection_summary()
```

At the start of `start_precheck`, clear previous results:

```python
self.last_precheck_results = []
self.view.btn_export_precheck.setProperty("hasPrecheckResults", False)
```

Connect the worker signal:

```python
self.precheck_worker.result_ready.connect(self._record_precheck_result)
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest tests/test_regressions.py::RegressionTests::test_precheck_results_are_stored_and_export_button_enables -v`

Expected: PASS.

### Task 3: Export Latest Precheck Results To CSV

**Files:**
- Modify: `controller.py`
- Test: `tests/test_regressions.py`

- [ ] **Step 1: Write the failing test**

Add this test to `tests/test_regressions.py`:

```python
def test_export_precheck_results_writes_csv(self):
    window = MainWindow()
    controller = MainController(window)
    try:
        with tempfile.TemporaryDirectory() as temp_dir:
            csv_path = os.path.join(temp_dir, "precheck.csv")
            controller.last_precheck_results = [
                {
                    "file_name": "needs-work.pdf",
                    "file_path": os.path.join(temp_dir, "needs-work.pdf"),
                    "status": "建议处理",
                    "suggestions": "设置导览标签、同步文件名为标题",
                    "error": "",
                },
                {
                    "file_name": "broken.pdf",
                    "file_path": os.path.join(temp_dir, "broken.pdf"),
                    "status": "预检失败",
                    "suggestions": "",
                    "error": "无法读取PDF结构",
                },
            ]

            with patch("controller.QFileDialog.getSaveFileName", return_value=(csv_path, "CSV Files (*.csv)")), patch.object(
                window,
                "show_success_message",
            ):
                controller.export_precheck_results()

            with open(csv_path, "r", encoding="utf-8-sig", newline="") as f:
                rows = list(csv.DictReader(f))

        self.assertEqual(rows[0]["file_name"], "needs-work.pdf")
        self.assertEqual(rows[0]["status"], "建议处理")
        self.assertEqual(rows[0]["suggestions"], "设置导览标签、同步文件名为标题")
        self.assertEqual(rows[1]["file_name"], "broken.pdf")
        self.assertEqual(rows[1]["error"], "无法读取PDF结构")
    finally:
        window.close()
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_regressions.py::RegressionTests::test_export_precheck_results_writes_csv -v`

Expected: FAIL because `export_precheck_results` does not exist.

- [ ] **Step 3: Write minimal implementation**

Add `export_precheck_results` to `MainController` near `export_logs`:

```python
def export_precheck_results(self):
    if not self.last_precheck_results:
        self.view.show_warning_message("⚠️ 提示", "请先执行一次批量预检，再导出预检结果。")
        return

    default_dir = ""
    if self.view.settings_dialog.default_output_edit.text().strip() and os.path.isdir(self.view.settings_dialog.default_output_edit.text().strip()):
        default_dir = self.view.settings_dialog.default_output_edit.text().strip()
    elif self.loaded_files:
        try:
            file_dirs = [os.path.dirname(os.path.abspath(f)) for f in self.loaded_files]
            default_dir = os.path.commonpath(file_dirs)
        except ValueError:
            default_dir = os.path.dirname(os.path.abspath(self.loaded_files[0]))

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    default_filename = f"RATools_precheck_results_{timestamp}.csv"
    default_path = os.path.join(default_dir, default_filename) if default_dir else default_filename

    file_path, _selected_filter = QFileDialog.getSaveFileName(
        self.view,
        "导出预检结果",
        default_path,
        "CSV Files (*.csv);;All Files (*)",
    )
    if not file_path:
        return
    if not file_path.lower().endswith(".csv"):
        file_path += ".csv"

    try:
        fieldnames = ["file_name", "file_path", "status", "suggestions", "error"]
        with open(file_path, "w", encoding="utf-8-sig", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(self.last_precheck_results)
        self.view.show_success_message("✅ 导出成功", "预检结果已成功保存！")
    except Exception as e:
        self.view.show_error_message("❌ 导出失败", f"文件保存失败：\n{str(e)}")
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest tests/test_regressions.py::RegressionTests::test_export_precheck_results_writes_csv -v`

Expected: PASS.

### Task 4: Regression Verification

**Files:**
- Test: `tests/test_regressions.py`

- [ ] **Step 1: Run focused precheck/export regressions**

Run: `python -m pytest tests/test_regressions.py -v`

Expected: PASS.

- [ ] **Step 2: Inspect changed files**

Run: `git diff -- controller.py view.py tests/test_regressions.py docs/superpowers/specs/2026-05-29-precheck-export-design.md docs/superpowers/plans/2026-05-29-precheck-export.md`

Expected: only precheck export changes and docs are present.

## Self-Review

- Spec coverage: covered button, latest-result storage, CSV fields, no rescan, error handling, and tests.
- Placeholder scan: no TODO/TBD placeholders.
- Type consistency: result rows use `dict` with `file_name`, `file_path`, `status`, `suggestions`, `error` across worker, controller, and CSV export.
