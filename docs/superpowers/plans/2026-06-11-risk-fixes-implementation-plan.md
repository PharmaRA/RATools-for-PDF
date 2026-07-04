# RATools Risk Fixes Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Repair the highest-risk behavior in `RATools-for-PDF`, add regression coverage for the repaired paths, and align build/docs behavior with the actual repository state.

**Architecture:** Keep the existing `controller.py` / `pdf_processor.py` / `view.py` split, but move risky behavior behind explicit data-structure and validation boundaries. Fix data-loss paths first (`bookmarks`, `links`, path traversal, rename collisions), then stabilize build and export/reporting behavior, then update docs and CI so future regressions are caught automatically.

**Tech Stack:** Python 3, PyMuPDF (`fitz`), PySide6, qpdf, unittest/pytest-style test runner compatible with `python -m unittest`.

---

## File Map

- Modify: `pdf_processor.py`
  Responsibility: bookmark/link import-export fidelity, smart-mode rule resolution behavior, qpdf-related processing behavior.
- Modify: `controller.py`
  Responsibility: IO path validation, rename collision detection, structured log export source, batch orchestration guardrails.
- Modify: `view.py`
  Responsibility: user-facing wording for smart mode and cleanup options where semantics must be clarified.
- Modify: `README.md`
  Responsibility: document actual test/build/support status and clarified feature semantics.
- Modify: `build_nuitka.bat`
  Responsibility: make the optional Nuitka path self-consistent or explicitly deprecate it.
- Modify: `.github/workflows/build.yml`
  Responsibility: run regression tests before packaging.
- Create: `tests/test_pdf_processor_roundtrip.py`
  Responsibility: bookmark/link round-trip and qpdf behavior regression tests.
- Create: `tests/test_controller_guards.py`
  Responsibility: rename collision, IO path traversal, structured log export parsing/serialization tests.

---

### Task 1: Add regression test scaffolding

**Files:**
- Create: `tests/test_pdf_processor_roundtrip.py`
- Create: `tests/test_controller_guards.py`
- Modify: `README.md:389`

- [ ] **Step 1: Create the initial round-trip regression tests**

```python
import os
import tempfile
import unittest

import fitz

from pdf_processor import PDFProcessor


class PDFProcessorRoundTripTests(unittest.TestCase):
    def test_bookmark_export_import_preserves_external_and_internal_targets(self):
        with tempfile.TemporaryDirectory() as tmp:
            source_pdf = os.path.join(tmp, "source.pdf")
            csv_path = os.path.join(tmp, "bookmarks.csv")
            output_pdf = os.path.join(tmp, "output.pdf")

            doc = fitz.open()
            doc.new_page()
            doc.new_page()
            doc.set_toc([
                [1, "External", -1, {"kind": fitz.LINK_URI, "uri": "https://example.com"}],
                [1, "Internal", 2, {"kind": fitz.LINK_GOTO, "page": 1, "to": fitz.Point(144, 288), "zoom": 2.0}],
            ])
            doc.save(source_pdf)
            doc.close()

            PDFProcessor.export_bookmarks(source_pdf, csv_path)
            PDFProcessor.import_bookmarks(source_pdf, csv_path, output_pdf)

            restored = fitz.open(output_pdf)
            toc = restored.get_toc(simple=False)
            restored.close()

            self.assertEqual(toc[0][3].get("kind"), fitz.LINK_URI)
            self.assertEqual(toc[0][3].get("uri"), "https://example.com")
            self.assertEqual(toc[1][3].get("kind"), fitz.LINK_GOTO)
            self.assertEqual(tuple(toc[1][3].get("to")), (144.0, 288.0))
            self.assertEqual(toc[1][3].get("zoom"), 2.0)
```

- [ ] **Step 2: Create the initial controller guard regression tests**

```python
import unittest

from controller import _build_io_paths_for_file


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
```

- [ ] **Step 3: Run test discovery to verify the new tests fail for the known defects**

Run: `python -m unittest tests.test_pdf_processor_roundtrip tests.test_controller_guards -v`
Expected: FAIL in bookmark/link fidelity and IO path escape assertions.

- [ ] **Step 4: Update `README.md` test section to reflect the new canonical command**

```markdown
## 测试

当前仓库维护最小回归测试集，优先覆盖：

- 书签导入/导出 round-trip
- 链接导入/导出 round-trip
- qpdf 相关重写行为
- 批量导出路径与文件名保护逻辑

运行方式：

```bash
python -m unittest
```
```

- [ ] **Step 5: Commit**

```bash
git add tests/test_pdf_processor_roundtrip.py tests/test_controller_guards.py README.md
git commit -m "test: add regression scaffolding for risk fixes"
```

### Task 2: Repair bookmark import-export fidelity

**Files:**
- Modify: `pdf_processor.py:1259`
- Test: `tests/test_pdf_processor_roundtrip.py`

- [ ] **Step 1: Extend the bookmark export format with explicit action columns**

```python
writer.writerow([
    "Level",
    "Title",
    "Page",
    "Kind",
    "Uri",
    "File",
    "TargetPage",
    "ToX",
    "ToY",
    "Zoom",
    "NewWindow",
])
for level, title, page, dest in toc:
    dest = dest if isinstance(dest, dict) else {}
    point = dest.get("to")
    writer.writerow([
        level,
        title,
        page,
        dest.get("kind", fitz.LINK_NONE),
        dest.get("uri", ""),
        dest.get("file", ""),
        dest.get("page", ""),
        getattr(point, "x", ""),
        getattr(point, "y", ""),
        dest.get("zoom", ""),
        dest.get("newWindow", False),
    ])
```

- [ ] **Step 2: Keep backward compatibility for the old 3-column CSV format**

```python
kind_text = row.get("Kind", "")
if not kind_text:
    new_toc.append([level, title, bounded_page])
    continue
```

- [ ] **Step 3: Rebuild bookmark destinations from the exported action fields**

```python
kind = int(row.get("Kind", fitz.LINK_NONE))
if kind == fitz.LINK_URI:
    dest = {"kind": fitz.LINK_URI, "uri": row.get("Uri", "")}
elif kind == fitz.LINK_GOTO:
    dest = {
        "kind": fitz.LINK_GOTO,
        "page": int(row.get("TargetPage", bounded_page - 1)),
        "to": fitz.Point(float(row.get("ToX", 72.0)), float(row.get("ToY", 36.0))),
        "zoom": float(row.get("Zoom", 0.0) or 0.0),
    }
elif kind == fitz.LINK_GOTOR:
    dest = {
        "kind": fitz.LINK_GOTOR,
        "file": row.get("File", ""),
        "page": int(row.get("TargetPage", 0) or 0),
        "to": fitz.Point(float(row.get("ToX", 72.0)), float(row.get("ToY", 36.0))),
        "zoom": float(row.get("Zoom", 0.0) or 0.0),
        "newWindow": str(row.get("NewWindow", "false")).lower() == "true",
    }
else:
    dest = {"kind": fitz.LINK_NONE}
new_toc.append([level, title, bounded_page, dest])
```

- [ ] **Step 4: Run the bookmark round-trip test until it passes**

Run: `python -m unittest tests.test_pdf_processor_roundtrip.PDFProcessorRoundTripTests.test_bookmark_export_import_preserves_external_and_internal_targets -v`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add pdf_processor.py tests/test_pdf_processor_roundtrip.py
git commit -m "fix: preserve bookmark actions during import export"
```

### Task 3: Repair link import-export fidelity

**Files:**
- Modify: `pdf_processor.py:1309`
- Test: `tests/test_pdf_processor_roundtrip.py`

- [ ] **Step 1: Add a failing link round-trip test**

```python
def test_link_export_import_preserves_internal_target_coordinates(self):
    with tempfile.TemporaryDirectory() as tmp:
        source_pdf = os.path.join(tmp, "source.pdf")
        json_path = os.path.join(tmp, "links.json")
        output_pdf = os.path.join(tmp, "output.pdf")

        doc = fitz.open()
        doc.new_page()
        doc.new_page()
        doc[0].insert_link({
            "kind": fitz.LINK_GOTO,
            "from": fitz.Rect(70, 60, 180, 80),
            "page": 1,
            "to": fitz.Point(144, 288),
        })
        doc.save(source_pdf)
        doc.close()

        PDFProcessor.export_links(source_pdf, json_path)
        PDFProcessor.import_links(source_pdf, json_path, output_pdf)

        restored = fitz.open(output_pdf)
        link = restored[0].get_links()[0]
        restored.close()

        self.assertEqual(tuple(link.get("to")), (144.0, 288.0))
```

- [ ] **Step 2: Export `to` and `newWindow` explicitly**

```python
target_point = link.get("to")
link_dict = {
    "page_index": page.number,
    "rect": [rect.x0, rect.y0, rect.x1, rect.y1],
    "kind": link.get("kind", fitz.LINK_NONE),
    "uri": link.get("uri", ""),
    "file": link.get("file", ""),
    "target_page": link.get("page", 0),
    "zoom": link.get("zoom", 0.0),
    "to": [getattr(target_point, "x", 0.0), getattr(target_point, "y", 0.0)] if target_point else None,
    "new_window": bool(link.get("newWindow", False)),
}
```

- [ ] **Step 3: Rebuild the link destination from the exported coordinates**

```python
to_value = ld.get("to")
if isinstance(to_value, (list, tuple)) and len(to_value) >= 2:
    new_link["to"] = fitz.Point(float(to_value[0]), float(to_value[1]))
if kind in [fitz.LINK_GOTOR, fitz.LINK_LAUNCH]:
    new_link["newWindow"] = bool(ld.get("new_window", False))
```

- [ ] **Step 4: Run the link round-trip tests**

Run: `python -m unittest tests.test_pdf_processor_roundtrip -v`
Expected: PASS for both bookmark and link round-trip tests.

- [ ] **Step 5: Commit**

```bash
git add pdf_processor.py tests/test_pdf_processor_roundtrip.py
git commit -m "fix: preserve link targets during import export"
```

### Task 4: Block IO path traversal outside target roots

**Files:**
- Modify: `controller.py:36`
- Test: `tests/test_controller_guards.py`

- [ ] **Step 1: Add a helper that sanitizes relative subpaths**

```python
def _safe_relative_subdir(file_path, common_base):
    if not common_base:
        return ""
    try:
        rel_dir = os.path.relpath(os.path.dirname(os.path.abspath(file_path)), common_base)
    except ValueError:
        return ""
    rel_norm = os.path.normpath(rel_dir)
    if rel_norm in (".", ""):
        return ""
    if rel_norm.startswith("..") or os.path.isabs(rel_norm):
        safe_leaf = os.path.basename(os.path.dirname(os.path.abspath(file_path))) or "external"
        return os.path.join("_external", safe_leaf)
    return rel_norm
```

- [ ] **Step 2: Use the sanitized helper inside `_build_io_paths_for_file()`**

```python
rel_dir = _safe_relative_subdir(file_path, common_base)
data_parent = os.path.join(target_dir, rel_dir) if rel_dir else target_dir
output_parent = os.path.join(output_dir, rel_dir) if rel_dir else output_dir
```

- [ ] **Step 3: Expand controller tests to cover nested and external inputs**

```python
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
```

- [ ] **Step 4: Run the controller guard tests**

Run: `python -m unittest tests.test_controller_guards -v`
Expected: PASS with no `..` in output paths.

- [ ] **Step 5: Commit**

```bash
git add controller.py tests/test_controller_guards.py
git commit -m "fix: prevent io exports from escaping target roots"
```

### Task 5: Detect eCTD rename collisions before processing

**Files:**
- Modify: `controller.py:1545`
- Test: `tests/test_controller_guards.py`

- [ ] **Step 1: Extract filename normalization into a reusable helper**

```python
def _normalized_ectd_name(base_name, fallback_index):
    name, ext = os.path.splitext(base_name)
    normalized = name.lower().replace(" ", "-")
    normalized = re.sub(r"[^a-z0-9_-]", "", normalized)
    if not normalized:
        normalized = f"doc_{fallback_index:03d}"
    return f"{normalized}{ext.lower()}"
```

- [ ] **Step 2: Build a destination map and stop processing on collisions**

```python
target_names = {}
for i, file_path in enumerate(processing_files, start=1):
    target_name = _normalized_ectd_name(os.path.basename(file_path), i)
    target_names.setdefault(target_name, []).append(file_path)

collisions = {name: paths for name, paths in target_names.items() if len(paths) > 1}
if collisions:
    details = "\n".join(f"{name}: {len(paths)} files" for name, paths in sorted(collisions.items()))
    self.view.show_error_message("文件名冲突", f"eCTD 格式化会产生重名输出，已阻止处理：\n{details}")
    return
```

- [ ] **Step 3: Add a direct unit test for collision detection input data**

```python
def test_normalized_ectd_names_detect_collisions(self):
    names = ["A B.pdf", "a-b.pdf"]
    normalized = [controller._normalized_ectd_name(name, idx + 1) for idx, name in enumerate(names)]
    self.assertEqual(normalized[0], normalized[1])
```

- [ ] **Step 4: Run the collision tests**

Run: `python -m unittest tests.test_controller_guards -v`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add controller.py tests/test_controller_guards.py
git commit -m "fix: block ectd rename collisions before processing"
```

### Task 6: Stabilize structured log export and stop parsing human-readable logs

**Files:**
- Modify: `controller.py:61`
- Test: `tests/test_controller_guards.py`

- [ ] **Step 1: Add a failing test for structured export rows**

```python
def test_render_logs_as_csv_rows_parses_known_log_format(self):
    rows = _render_logs_as_csv_rows(
        "[10:00:00] 开始处理: C:/a.pdf\n"
        "    输出文件: C:/out/a.pdf\n"
        "[10:00:05] C:/a.pdf\n"
        "    状态: 处理完成\n"
        "    结果: ✅ 处理成功；修改项：Foo、Bar\n"
    )
    self.assertEqual(len(rows), 1)
    self.assertEqual(rows[0]["changes"], "Foo、Bar")
```

- [ ] **Step 2: Introduce a structured in-memory event list on the controller**

```python
self.process_log_rows = []
```

- [ ] **Step 3: Record structured rows directly inside `update_progress()` for terminal statuses**

```python
if status_text in ["处理完成", "处理失败", "已跳过"]:
    self.process_log_rows.append({
        "time": datetime.now().strftime("%H:%M:%S"),
        "file_original": file_path,
        "file_output": "",
        "status": status_text,
        "success": "true" if status_text == "处理完成" else "false",
        "duration_sec": "",
        "changes": "",
    })
```

- [ ] **Step 4: Prefer structured rows during CSV export and keep `_render_logs_as_csv_rows()` only as fallback**

```python
rows = self.process_log_rows or _render_logs_as_csv_rows(self.process_logs)
writer.writerows(rows)
```

- [ ] **Step 5: Run guard tests and smoke-check CSV export behavior**

Run: `python -m unittest tests.test_controller_guards -v`
Expected: PASS

- [ ] **Step 6: Commit**

```bash
git add controller.py tests/test_controller_guards.py
git commit -m "fix: export structured process logs reliably"
```

### Task 7: Repair or retire the optional Nuitka build path

**Files:**
- Modify: `build_nuitka.bat:1`
- Modify: `README.md:314`

- [ ] **Step 1: Replace the missing icon reference with the actual repository asset**

```bat
  --include-data-files="%ROOT_DIR%icon.ico=icon.ico" ^
```

- [ ] **Step 2: Add upfront file existence checks for optional build assets**

```bat
if not exist "%ROOT_DIR%icon.ico" (
    echo [ERROR] Cannot find icon.ico in %ROOT_DIR%
    exit /b 1
)
if not exist "%ROOT_DIR%plugins" (
    echo [ERROR] Cannot find plugins directory in %ROOT_DIR%
    exit /b 1
)
```

- [ ] **Step 3: Update README to state whether Nuitka is supported or deprecated**

```markdown
### 5. Nuitka 说明

当前仓库仅将 `build_nuitka.bat` 视为可选构建路径。若维护该路径，需与 `icon.ico`、`plugins/` 目录和当前依赖保持同步；若未验证，请不要将其作为正式发布流程。
```

- [ ] **Step 4: Smoke-run script validation without packaging**

Run: `cmd /c build_nuitka.bat`
Expected: Either progresses past input-file validation or fails only on missing Nuitka/dependency state, not on missing repository assets.

- [ ] **Step 5: Commit**

```bash
git add build_nuitka.bat README.md
git commit -m "build: repair optional nuitka script inputs"
```

### Task 8: Clarify smart-mode and cleanup semantics in UI and docs

**Files:**
- Modify: `view.py:1070`
- Modify: `README.md:135`
- Test: `tests/test_pdf_processor_roundtrip.py`

- [ ] **Step 1: Add a regression test for smart mode unsupported-option logging**

```python
def test_smart_mode_reports_unsupported_options_as_forced(self):
    with tempfile.TemporaryDirectory() as tmp:
        source_pdf = os.path.join(tmp, "source.pdf")
        doc = fitz.open()
        doc.new_page(width=595, height=842)
        doc.save(source_pdf)
        doc.close()

        result = PDFProcessor.resolve_processing_options(source_pdf, {"page_size_a4"}, "smart")

        self.assertIn("page_size_a4", result["forced_unsupported"])
```

- [ ] **Step 2: Update smart-mode descriptions in the UI**

```python
{"id": "page_size_a4", "title": "适配到A4尺寸", "desc": "此项无法通过预检精确判断是否需要处理；智能模式下若勾选仍会执行"}
```

- [ ] **Step 3: Clarify the cleanup option wording**

```python
{"id": "cleanup_remove_all_links_bookmarks", "title": "移除全部链接和书签", "desc": "仅删除页面链接与书签，不删除普通批注"}
```

- [ ] **Step 4: Update README feature text to match the actual semantics**

```markdown
- 智能处理模式只会自动跳过“可预检且未命中”的规则；无法预检的规则在勾选后仍会执行。
- “移除全部链接和书签”不会删除普通批注；如需清理批注，请单独启用“清理所有高亮/批注”。
```

- [ ] **Step 5: Run the smart-mode unit test**

Run: `python -m unittest tests.test_pdf_processor_roundtrip -v`
Expected: PASS

- [ ] **Step 6: Commit**

```bash
git add view.py README.md tests/test_pdf_processor_roundtrip.py
git commit -m "docs: clarify smart mode and cleanup semantics"
```

### Task 9: Wire tests into CI and finish documentation alignment

**Files:**
- Modify: `.github/workflows/build.yml:1`
- Modify: `README.md:389`
- Test: `tests/test_pdf_processor_roundtrip.py`
- Test: `tests/test_controller_guards.py`

- [ ] **Step 1: Add a test step before packaging in CI**

```yaml
    - name: Run regression tests
      run: |
        python -m unittest -v
```

- [ ] **Step 2: Ensure README test/build sections match the actual CI behavior**

```markdown
GitHub Actions 在打包前会先执行：

```bash
python -m unittest -v
```
```

- [ ] **Step 3: Run the full local regression suite**

Run: `python -m unittest -v`
Expected: PASS for all new regression tests.

- [ ] **Step 4: Commit**

```bash
git add .github/workflows/build.yml README.md tests/test_pdf_processor_roundtrip.py tests/test_controller_guards.py
git commit -m "ci: run regression tests before packaging"
```

---

## Execution Notes

- Implement Tasks 2 through 5 before touching docs-only tasks. These are the data-loss and boundary-protection fixes.
- Keep the old import/export formats backward compatible where practical; preserve forward fidelity first, compatibility second.
- Do not add GUI automation in this pass. Keep tests focused on deterministic, no-UI logic.
- Remove `tmp_review/` from the worktree before finalizing implementation if it is still untracked.

## Self-Review

- Spec coverage: covers bookmark fidelity, link fidelity, path escape prevention, rename collision prevention, build-script repair, smart-mode wording, log export stability, test/CI/docs alignment.
- Placeholder scan: no `TODO`/`TBD` placeholders remain in tasks.
- Type consistency: all planned tests and code changes reference current module names and existing functions in `controller.py`, `pdf_processor.py`, `view.py`, `README.md`, and `.github/workflows/build.yml`.

## Execution Handoff

Plan complete and saved to `docs/superpowers/plans/2026-06-11-risk-fixes-implementation-plan.md`. Two execution options:

**1. Subagent-Driven (recommended)** - I dispatch a fresh subagent per task, review between tasks, fast iteration

**2. Inline Execution** - Execute tasks in this session using executing-plans, batch execution with checkpoints

Which approach?
