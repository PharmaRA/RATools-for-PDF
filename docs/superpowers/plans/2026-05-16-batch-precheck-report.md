# Batch Pre-Check Report Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a batch pre-check action that scans queued PDFs and reports which processing options are likely relevant for each file before the user starts processing.

**Architecture:** Add a read-only analysis layer in `pdf_processor.py` that inspects document structure and returns suggested option IDs plus human-readable reasons. Add a dedicated background worker in `controller.py` to run the analysis without blocking the UI, then surface the results through a new `预检` button in the footer and file-row status markers in the tree. Keep the existing processing pipeline unchanged.

**Tech Stack:** Python 3.14, PySide6, PyMuPDF (`fitz`), existing qpdf binary for file-level checks where useful.

---

### Task 1: Add failing analysis tests

**Files:**
- Modify: `tests/test_regressions.py`
- Modify: `test_pdf_processor_metadata.py`

- [ ] **Step 1: Write the failing test**

```python
def test_batch_precheck_suggests_initial_view_and_metadata_actions():
    with tempfile.TemporaryDirectory() as temp_dir:
        input_pdf = os.path.join(temp_dir, "input.pdf")

        doc = fitz.open()
        doc.new_page()
        doc.set_toc([
            [1, "Chapter 1", 1, {"kind": fitz.LINK_GOTO, "page": 0, "to": fitz.Point(72, 72), "zoom": 1.0}],
        ])
        doc.set_metadata({"title": "", "author": ""})
        doc.save(input_pdf)
        doc.close()

        report = PDFProcessor.build_precheck_report(input_pdf)

        assert report["available"] is True
        assert "initial_view_bookmarks_and_page" in report["suggestions"]
        assert "title_from_filename" in report["suggestions"]
        assert report["suggestions"]["initial_view_bookmarks_and_page"]["matched"] is True
        assert report["suggestions"]["title_from_filename"]["matched"] is True
```

```python
def test_batch_precheck_reports_linearization_and_restrictions_signals():
    with tempfile.TemporaryDirectory() as temp_dir:
        input_pdf = os.path.join(temp_dir, "input.pdf")

        doc = fitz.open()
        doc.new_page()
        doc.save(input_pdf)
        doc.close()

        report = PDFProcessor.build_precheck_report(input_pdf)

        assert "fast_web_view" in report["suggestions"]
        assert "convert_pdf_version" in report["suggestions"]
        assert "remove_pdf_restrictions" in report["suggestions"]
```
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_regressions.py::RegressionTests::test_batch_precheck_suggests_initial_view_and_metadata_actions -v`

Expected: FAIL with `AttributeError` or missing assertion because `build_precheck_report` does not exist yet.

- [ ] **Step 3: Write minimal implementation**

Implement a new static method in `pdf_processor.py`:

```python
@staticmethod
def build_precheck_report(input_path):
    ...
```

It should return a dictionary with at least these keys:

```python
{
    "available": True,
    "file_path": input_path,
    "file_name": os.path.basename(input_path),
    "suggestions": {
        "initial_view_bookmarks_and_page": {"matched": True, "reason": "..."},
        "title_from_filename": {"matched": True, "reason": "..."},
        "fast_web_view": {"matched": True, "reason": "..."},
        "convert_pdf_version": {"matched": True, "reason": "..."},
        "remove_pdf_restrictions": {"matched": True, "reason": "..."},
    },
}
```

Use `fitz.open(input_path)` for file inspection. Reuse existing logic where possible:
- `doc.get_toc(simple=False)` for bookmarks
- `doc.metadata` for title checking
- `doc.needs_pass` for encryption detection
- `doc.page_count` and `doc.pdf_catalog()` if needed for initial-view checks
- a small helper to parse the PDF header version if you need a direct version signal

For this first pass, the matcher can be conservative and only report obvious matches.

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest tests/test_regressions.py::RegressionTests::test_batch_precheck_suggests_initial_view_and_metadata_actions -v`

Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add tests/test_regressions.py test_pdf_processor_metadata.py pdf_processor.py
git commit -m "feat: add batch pre-check diagnostics"
```

### Task 2: Add the pre-check worker and controller plumbing

**Files:**
- Modify: `controller.py`
- Modify: `view.py`
- Modify: `tests/test_regressions.py`

- [ ] **Step 1: Write the failing test**

```python
def test_precheck_button_is_visible_and_disabled_without_selection():
    window = MainWindow()
    try:
        self.assertTrue(hasattr(window, "btn_precheck"))
        self.assertFalse(window.btn_precheck.isEnabled())
    finally:
        window.close()
```

```python
def test_update_progress_can_mark_precheck_state():
    window = MainWindow()
    controller = MainController(window)
    try:
        controller.loaded_files = [os.path.join(tempfile.gettempdir(), "sample.pdf")]
        controller.processing_files = []
        controller.update_progress(0, "预检完成", "")
        self.assertEqual(window.tree.topLevelItem(0).text(2), "预检完成")
    finally:
        window.close()
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_regressions.py::RegressionTests::test_precheck_button_is_visible_and_disabled_without_selection -v`

Expected: FAIL because `btn_precheck` does not exist yet.

- [ ] **Step 3: Write minimal implementation**

Add a `PreCheckWorker(QThread)` in `controller.py` that mirrors `ProcessWorker` but calls `PDFProcessor.build_precheck_report(file_path)` instead of `process_document(...)`.

Wire a new `self.view.btn_precheck` into `MainController.setup_connections()` and implement `start_precheck()`.

In `view.py`, add the button next to `btn_start` in the footer:

```python
self.btn_precheck = QPushButton("🔎 预检")
self.btn_precheck.setObjectName("actionBtn")
footer_layout.addWidget(self.btn_precheck)
```

The controller should:
- require at least one loaded file and one selected rule
- disable the button while processing or while a precheck is already running
- update `process_logs` with a precheck summary block
- update tree row status text to `预检完成` / `预检失败`

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest tests/test_regressions.py::RegressionTests::test_precheck_button_is_visible_and_disabled_without_selection -v`

Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add controller.py view.py tests/test_regressions.py
git commit -m "feat: wire batch pre-check action into the UI"
```

### Task 3: Render pre-check suggestions in the tree and log

**Files:**
- Modify: `controller.py`
- Modify: `tests/test_regressions.py`

- [ ] **Step 1: Write the failing test**

```python
def test_precheck_results_color_tree_rows_and_append_logs():
    window = MainWindow()
    controller = MainController(window)
    try:
        controller.loaded_files = [os.path.join(tempfile.gettempdir(), "sample.pdf")]
        controller.process_logs = ""
        controller.update_progress(0, "预检完成", "[12:00:00] sample.pdf\n    建议: 设置导览标签")
        self.assertEqual(window.tree.topLevelItem(0).text(2), "预检完成")
        self.assertIn("建议", controller.process_logs)
    finally:
        window.close()
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_regressions.py::RegressionTests::test_precheck_results_color_tree_rows_and_append_logs -v`

Expected: FAIL because `update_progress` does not yet color precheck rows or store the new log format.

- [ ] **Step 3: Write minimal implementation**

Extend `update_progress` so that it treats `预检完成` and `预检失败` like other statuses, with green and red colors respectively.

When precheck logs arrive, append them to `process_logs` exactly like processing logs so the existing log dialog can display them.

Format each row’s log entry as:

```text
[HH:MM:SS] filename.pdf
    状态: 预检完成
    建议: 设置导览标签；同步文件名为标题
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest tests/test_regressions.py::RegressionTests::test_precheck_results_color_tree_rows_and_append_logs -v`

Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add controller.py tests/test_regressions.py
git commit -m "feat: show batch pre-check results in the queue"
```

### Task 4: Final verification and cleanup

**Files:**
- Modify: none unless a small follow-up fix is needed

- [ ] **Step 1: Run the focused test set**

Run: `python -m pytest tests/test_regressions.py test_pdf_processor_metadata.py test_pdf_processor_presets.py test_pdf_processor_navigation.py -v`

Expected: All tests pass.

- [ ] **Step 2: Run the app smoke check**

Run the application and verify manually that:
- the new `🔎 预检` button appears beside `▶ 开始处理`
- the button is disabled until files and rules are selected
- precheck updates the tree row state and logs

- [ ] **Step 3: Commit any follow-up fixes**

```bash
git add .
git commit -m "feat: polish batch pre-check report"
```
