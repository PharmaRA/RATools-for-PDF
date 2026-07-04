# PDF Decrypt Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a batch-processing option that removes PDF permission restrictions through bundled `qpdf`, without adding any password prompt workflow.

**Architecture:** Extend the existing `qpdf` rewrite helper in `pdf_processor.py` so one invocation can combine decryption, version conversion, and linearization. Surface the feature as one additional checkbox in `view.py`, and reuse the existing `MainController -> ProcessWorker -> PDFProcessor.process_document(...)` path unchanged.

**Tech Stack:** Python, PySide6, PyMuPDF (`fitz`), `qpdf`, `unittest`, `unittest.mock`

---

### Task 1: Add failing tests for the new qpdf rewrite mode

**Files:**
- Modify: `test_pdf_processor_tools.py`
- Test: `test_pdf_processor_tools.py`

- [ ] **Step 1: Write the failing tests**

Add these test methods to `PDFProcessorToolPathTests` in `test_pdf_processor_tools.py`:

```python
    @patch("pdf_processor.subprocess.run")
    @patch.object(PDFProcessor, "_get_qpdf_path", return_value="qpdf.exe")
    def test_rewrite_with_qpdf_adds_decrypt_flag(self, _get_qpdf_path, mock_run):
        mock_run.return_value.returncode = 0
        mock_run.return_value.stderr = ""

        PDFProcessor._rewrite_with_qpdf(
            "input.pdf",
            "output.pdf",
            decrypt_restrictions=True,
        )

        cmd = mock_run.call_args.args[0]
        self.assertIn("--decrypt", cmd)
        self.assertEqual(cmd[-2:], ["input.pdf", "output.pdf"])

    @patch("pdf_processor.subprocess.run")
    @patch.object(PDFProcessor, "_get_qpdf_path", return_value="qpdf.exe")
    def test_rewrite_with_qpdf_combines_decrypt_linearize_and_version(self, _get_qpdf_path, mock_run):
        mock_run.return_value.returncode = 0
        mock_run.return_value.stderr = ""

        PDFProcessor._rewrite_with_qpdf(
            "input.pdf",
            "output.pdf",
            force_version="1.7",
            linearize=True,
            decrypt_restrictions=True,
        )

        cmd = mock_run.call_args.args[0]
        self.assertIn("--decrypt", cmd)
        self.assertIn("--linearize", cmd)
        self.assertIn("--force-version=1.7", cmd)
        self.assertEqual(cmd[-2:], ["input.pdf", "output.pdf"])
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest test_pdf_processor_tools.py -v`

Expected: FAIL because `_rewrite_with_qpdf(...)` does not yet accept `decrypt_restrictions`.

- [ ] **Step 3: Write minimal implementation**

Do not implement the full feature yet. Only extend the `_rewrite_with_qpdf(...)` signature in `pdf_processor.py` and insert `--decrypt` when `decrypt_restrictions` is `True`.

Target shape:

```python
    @staticmethod
    def _rewrite_with_qpdf(input_pdf, output_pdf, force_version=None, linearize=False, decrypt_restrictions=False):
        qpdf_exe = PDFProcessor._get_qpdf_path()
        ...
        cmd = [qpdf_exe]
        if decrypt_restrictions:
            cmd.append("--decrypt")
        if linearize:
            cmd.append("--linearize")
        if force_version:
            cmd.append(f"--force-version={force_version}")
        cmd.extend([input_pdf, output_pdf])
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest test_pdf_processor_tools.py -v`

Expected: PASS for the two new tests.

- [ ] **Step 5: Commit**

```bash
git add test_pdf_processor_tools.py pdf_processor.py
git commit -m "test: cover qpdf decrypt rewrite options"
```

### Task 2: Add failing processing-level tests for option handling and user-facing errors

**Files:**
- Modify: `tests/test_regressions.py`
- Test: `tests/test_regressions.py`

- [ ] **Step 1: Write the failing tests**

Add these tests to `RegressionTests` in `tests/test_regressions.py`:

```python
    def test_remove_pdf_restrictions_triggers_qpdf_rewrite_and_reports_change(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            input_pdf = os.path.join(temp_dir, "input.pdf")
            output_pdf = os.path.join(temp_dir, "output.pdf")

            doc = fitz.open()
            doc.new_page()
            doc.save(input_pdf)
            doc.close()

            with patch.object(PDFProcessor, "_rewrite_with_qpdf") as rewrite:
                success, message = PDFProcessor.process_document(
                    input_pdf,
                    output_pdf,
                    {"remove_pdf_restrictions"},
                )

            self.assertTrue(success)
            rewrite.assert_called_once()
            self.assertTrue(rewrite.call_args.kwargs["decrypt_restrictions"])
            self.assertIn("已解除PDF权限限制", message)

    def test_remove_pdf_restrictions_password_error_is_reworded(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            input_pdf = os.path.join(temp_dir, "input.pdf")
            output_pdf = os.path.join(temp_dir, "output.pdf")

            doc = fitz.open()
            doc.new_page()
            doc.save(input_pdf)
            doc.close()

            with patch.object(
                PDFProcessor,
                "_rewrite_with_qpdf",
                side_effect=RuntimeError("qpdf 执行失败: invalid password"),
            ):
                success, message = PDFProcessor.process_document(
                    input_pdf,
                    output_pdf,
                    {"remove_pdf_restrictions"},
                )

            self.assertFalse(success)
            self.assertIn("当前模式不支持输入密码解锁", message)
```

Also add this import at the top of `tests/test_regressions.py`:

```python
from unittest.mock import patch
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest tests/test_regressions.py -v`

Expected: FAIL because `process_document(...)` does not yet route `remove_pdf_restrictions` through `qpdf` or rewrite password-related errors.

- [ ] **Step 3: Write minimal implementation**

Update `pdf_processor.py` inside `process_document(...)` so the output-transform logic includes the new option.

Target changes:

```python
            is_linear = "fast_web_view" in options
            force_pdf_version = "1.7" if "convert_pdf_version" in options else None
            remove_pdf_restrictions = "remove_pdf_restrictions" in options
            needs_qpdf_rewrite = bool(is_linear or force_pdf_version or remove_pdf_restrictions)
```

Pass the new flag in both `_rewrite_with_qpdf(...)` call sites:

```python
                            decrypt_restrictions=remove_pdf_restrictions,
```

Mark the change after successful rewrite:

```python
                        if remove_pdf_restrictions:
                            PDFProcessor._mark_change(applied_changes, "已解除PDF权限限制")
```

Add a small helper in `PDFProcessor` to translate common password-related `qpdf` failures into the approved user-facing message.

Target shape:

```python
    @staticmethod
    def _format_qpdf_error(error):
        text = str(error)
        lowered = text.lower()
        password_markers = ["password", "invalid password", "incorrect password", "encrypted"]
        if any(marker in lowered for marker in password_markers):
            return "该PDF需要密码，当前模式不支持输入密码解锁"
        if text.startswith("qpdf 执行失败:"):
            return f"未能移除PDF权限限制：{text.split(':', 1)[1].strip()}"
        return text
```

Use that helper in the exception path when the new option is selected.

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest tests/test_regressions.py -v`

Expected: PASS for the two new tests.

- [ ] **Step 5: Commit**

```bash
git add tests/test_regressions.py pdf_processor.py
git commit -m "feat: route pdf restriction removal through qpdf"
```

### Task 3: Add the UI option and a regression test for discoverability

**Files:**
- Modify: `view.py`
- Modify: `tests/test_regressions.py`
- Test: `tests/test_regressions.py`

- [ ] **Step 1: Write the failing test**

Add this test to `RegressionTests` in `tests/test_regressions.py`:

```python
    def test_remove_pdf_restrictions_option_is_visible_in_main_window(self):
        window = MainWindow()
        try:
            option_titles = []
            for mod in window.MODULES_DATA:
                for opt in mod["options"]:
                    option_titles.append(opt["title"])
            self.assertIn("PDF解除权限限制", option_titles)
        finally:
            window.close()
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_regressions.py::RegressionTests::test_remove_pdf_restrictions_option_is_visible_in_main_window -v`

Expected: FAIL because the new option does not exist in `view.py` yet.

- [ ] **Step 3: Write minimal implementation**

In `view.py`, add one option entry to the `文件级优化与输出` section in `MODULES_DATA`.

Insert this item alongside `convert_pdf_version` and `fast_web_view`:

```python
                    {"id": "remove_pdf_restrictions", "title": "PDF解除权限限制", "desc": "尝试移除禁止复制、打印、编辑等权限限制，不处理需要打开密码的加密文档"},
```

No controller changes are needed for this task.

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest tests/test_regressions.py::RegressionTests::test_remove_pdf_restrictions_option_is_visible_in_main_window -v`

Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add view.py tests/test_regressions.py
git commit -m "feat: expose pdf restriction removal option"
```

### Task 4: Verify the whole feature set together

**Files:**
- Modify: none
- Test: `test_pdf_processor_tools.py`
- Test: `tests/test_regressions.py`

- [ ] **Step 1: Run the focused automated test set**

Run: `python -m pytest test_pdf_processor_tools.py tests/test_regressions.py -v`

Expected: PASS.

- [ ] **Step 2: Run one combined rewrite test manually if a suitable sample exists**

Use one PDF that can be opened normally and run the app with these combined options selected:

```text
PDF解除权限限制
PDF版本转换
启用线性化 (快速网页浏览)
```

Expected:
- output file is produced
- processing log includes `已解除PDF权限限制`
- processing log includes `已转换PDF版本`
- processing log includes `已启用快速网页浏览`

If no suitable restricted sample PDF exists locally, explicitly record that only the automated path-level verification was run.

- [ ] **Step 3: Check worktree status before handoff**

Run: `git status --short`

Expected: only the intended feature files are modified, plus any pre-existing unrelated edits already present in the workspace.

- [ ] **Step 4: Commit**

```bash
git add view.py pdf_processor.py test_pdf_processor_tools.py tests/test_regressions.py
git commit -m "feat: add pdf restriction removal with qpdf"
```
