# Font Precheck Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add read-only font precheck that reports unembedded fonts, non-Base-14 fonts, and substitute-font risk without enabling font embedding.

**Architecture:** Implement font scanning as focused helpers on `PDFProcessor`, integrate the result into the existing `build_precheck_report()` `report_only` path, then propagate two text fields through `PreCheckWorker`, CSV export, and file details. Keep `embed_nonstandard_fonts` disabled and blocked.

**Tech Stack:** Python 3, PyMuPDF (`fitz`), PySide6 controller/view model, `unittest`/`pytest`.

---

## File Structure

- Create: `test_pdf_processor_font_precheck.py`
- Modify: `pdf_processor.py`
- Modify: `controller.py`
- Modify: `README.md`
- Preserve: `test_gs_removal.py`

`test_pdf_processor_font_precheck.py` owns all new font precheck unit tests. It uses fake document/page objects for deterministic font-resource cases, and temporary PDFs only when exercising `build_precheck_report()` end-to-end.

`pdf_processor.py` owns font classification, embedding-state inspection, finding aggregation, text formatting, and `report_only` precheck integration.

`controller.py` owns propagation of `font_summary` and `font_details` into worker result rows, CSV export, and file detail text.

`README.md` documents that precheck now reports font risks while embedding remains unavailable.

No new external dependency is introduced.

---

### Task 1: Font Classification Helpers

**Files:**
- Create: `test_pdf_processor_font_precheck.py`
- Modify: `pdf_processor.py`

- [ ] **Step 1: Write failing helper tests**

Create `test_pdf_processor_font_precheck.py` with these tests:

```python
import csv
import os
import tempfile
import unittest
from unittest.mock import patch

import fitz

from controller import MainController, PreCheckWorker
from pdf_processor import PDFProcessor


class FontPrecheckHelperTests(unittest.TestCase):
    def test_normalize_font_name_removes_subset_prefix(self):
        self.assertEqual("Calibri", PDFProcessor._normalize_font_name("ABCDEF+Calibri"))
        self.assertEqual("SimSun", PDFProcessor._normalize_font_name("SimSun"))
        self.assertEqual("", PDFProcessor._normalize_font_name(""))

    def test_base14_font_recognition_includes_variants(self):
        self.assertTrue(PDFProcessor._is_base14_font("Helvetica"))
        self.assertTrue(PDFProcessor._is_base14_font("Helvetica-BoldOblique"))
        self.assertTrue(PDFProcessor._is_base14_font("Times-BoldItalic"))
        self.assertTrue(PDFProcessor._is_base14_font("ABCDEF+Courier-Oblique"))
        self.assertTrue(PDFProcessor._is_base14_font("Symbol"))
        self.assertTrue(PDFProcessor._is_base14_font("ZapfDingbats"))
        self.assertFalse(PDFProcessor._is_base14_font("Calibri"))
        self.assertFalse(PDFProcessor._is_base14_font("SimSun"))

    def test_format_font_page_numbers_compacts_ranges(self):
        self.assertEqual("第1-3,5页", PDFProcessor._format_font_page_numbers([3, 1, 2, 5]))
        self.assertEqual("第7页", PDFProcessor._format_font_page_numbers([7]))
        self.assertEqual("页码未知", PDFProcessor._format_font_page_numbers([]))


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Run helper tests and verify failure**

Run: `python -m pytest test_pdf_processor_font_precheck.py::FontPrecheckHelperTests -v`

Expected: FAIL because `PDFProcessor._normalize_font_name`, `PDFProcessor._is_base14_font`, or `PDFProcessor._format_font_page_numbers` does not exist.

- [ ] **Step 3: Add minimal helper implementation**

In `pdf_processor.py`, inside `class PDFProcessor`, add these definitions after `PRECHECK_OPTION_TITLES`:

```python
    BASE14_FONT_NAMES = {
        "Courier",
        "Courier-Bold",
        "Courier-Oblique",
        "Courier-BoldOblique",
        "Helvetica",
        "Helvetica-Bold",
        "Helvetica-Oblique",
        "Helvetica-BoldOblique",
        "Times-Roman",
        "Times-Bold",
        "Times-Italic",
        "Times-BoldItalic",
        "Symbol",
        "ZapfDingbats",
    }

    @staticmethod
    def _normalize_font_name(font_name):
        name = str(font_name or "").strip()
        if re.match(r"^[A-Z]{6}\+", name):
            name = name.split("+", 1)[1]
        return name

    @staticmethod
    def _is_base14_font(font_name):
        return PDFProcessor._normalize_font_name(font_name) in PDFProcessor.BASE14_FONT_NAMES

    @staticmethod
    def _format_font_page_numbers(page_numbers):
        pages = sorted({int(page) for page in page_numbers if int(page) > 0})
        if not pages:
            return "页码未知"

        ranges = []
        start = pages[0]
        prev = pages[0]
        for page in pages[1:]:
            if page == prev + 1:
                prev = page
                continue
            ranges.append(f"{start}-{prev}" if start != prev else str(start))
            start = page
            prev = page
        ranges.append(f"{start}-{prev}" if start != prev else str(start))
        return f"第{','.join(ranges)}页"
```

- [ ] **Step 4: Run helper tests and verify pass**

Run: `python -m pytest test_pdf_processor_font_precheck.py::FontPrecheckHelperTests -v`

Expected: PASS for all three tests.

- [ ] **Step 5: Review diff without committing**

Run: `git diff -- pdf_processor.py test_pdf_processor_font_precheck.py`

Expected: Diff only contains the new helper tests and helper methods. Do not commit unless the user explicitly asks for a commit.

---

### Task 2: Font Finding Collection

**Files:**
- Modify: `test_pdf_processor_font_precheck.py`
- Modify: `pdf_processor.py`

- [ ] **Step 1: Add fake document tests for collection**

Append these fake classes and tests to `test_pdf_processor_font_precheck.py` before the `if __name__ == "__main__"` block:

```python
class _FakePage:
    def __init__(self, fonts):
        self._fonts = fonts

    def get_fonts(self, full=False):
        return list(self._fonts)


class _FakeFontDoc:
    def __init__(self, pages, objects=None, failing_xrefs=None):
        self._pages = list(pages)
        self._objects = dict(objects or {})
        self._failing_xrefs = set(failing_xrefs or [])

    def __iter__(self):
        return iter(self._pages)

    def xref_object(self, xref):
        if xref in self._failing_xrefs:
            raise RuntimeError("xref read failed")
        return self._objects.get(xref, "")


class FontPrecheckCollectionTests(unittest.TestCase):
    def test_collects_unembedded_non_base14_substitution_risk(self):
        doc = _FakeFontDoc([
            _FakePage([(10, "n/a", "TrueType", "ABCDEF+Calibri", "F1", "WinAnsiEncoding")]),
            _FakePage([(10, "n/a", "TrueType", "ABCDEF+Calibri", "F1", "WinAnsiEncoding")]),
        ], objects={10: "<< /Type /Font /BaseFont /ABCDEF+Calibri >>"})

        result = PDFProcessor._collect_font_precheck_findings(doc)

        self.assertEqual("未嵌入字体 1 个，非标准字体 1 个，替代字体风险 1 个", result["font_summary"])
        self.assertIn("Calibri", result["font_details"])
        self.assertIn("第1-2页", result["font_details"])
        self.assertIn("替代风险", result["font_details"])
        finding = result["font_findings"][0]
        self.assertEqual("Calibri", finding["normalized_name"])
        self.assertFalse(finding["embedded"])
        self.assertFalse(finding["base14"])
        self.assertTrue(finding["substitution_risk"])

    def test_collects_embedded_non_base14_without_substitution_risk(self):
        doc = _FakeFontDoc([
            _FakePage([(11, "ttf", "TrueType", "SimSun", "F2", "Identity-H")]),
        ], objects={
            11: "<< /Type /Font /BaseFont /SimSun /FontDescriptor 20 0 R >>",
            20: "<< /Type /FontDescriptor /FontFile2 30 0 R >>",
        })

        result = PDFProcessor._collect_font_precheck_findings(doc)

        self.assertEqual("非标准字体 1 个", result["font_summary"])
        self.assertIn("SimSun", result["font_details"])
        self.assertIn("已嵌入", result["font_details"])
        self.assertNotIn("替代风险", result["font_details"])

    def test_collects_unknown_embedding_status_as_review_only_signal(self):
        doc = _FakeFontDoc([
            _FakePage([(12, "n/a", "TrueType", "UnknownFont", "F3", "WinAnsiEncoding")]),
        ], failing_xrefs={12})

        result = PDFProcessor._collect_font_precheck_findings(doc)

        self.assertEqual("非标准字体 1 个，嵌入状态未知字体 1 个", result["font_summary"])
        self.assertIn("UnknownFont", result["font_details"])
        self.assertIn("嵌入状态未知", result["font_details"])
        self.assertNotIn("替代风险", result["font_details"])
```

- [ ] **Step 2: Run collection tests and verify failure**

Run: `python -m pytest test_pdf_processor_font_precheck.py::FontPrecheckCollectionTests -v`

Expected: FAIL because `_collect_font_precheck_findings` does not exist.

- [ ] **Step 3: Add embedding and collection helpers**

In `pdf_processor.py`, inside `class PDFProcessor`, add these methods after `_format_font_page_numbers`:

```python
    @staticmethod
    def _font_object_has_embedded_file(doc, xref, seen=None):
        if not xref:
            return False, True
        if seen is None:
            seen = set()
        if xref in seen:
            return False, True
        seen.add(xref)

        try:
            obj = doc.xref_object(int(xref))
        except Exception:
            return False, False

        if re.search(r"/FontFile(?:2|3)?\s+\d+\s+0\s+R", obj):
            return True, True

        refs = []
        refs.extend(int(match) for match in re.findall(r"/FontDescriptor\s+(\d+)\s+0\s+R", obj))
        refs.extend(int(match) for match in re.findall(r"/DescendantFonts\s*\[\s*(\d+)\s+0\s+R", obj))

        known = True
        for ref in refs:
            embedded, child_known = PDFProcessor._font_object_has_embedded_file(doc, ref, seen)
            known = known and child_known
            if embedded:
                return True, True
        return False, known

    @staticmethod
    def _font_tuple_value(font_tuple, index):
        if len(font_tuple) <= index:
            return ""
        return font_tuple[index]

    @staticmethod
    def _font_tuple_embedded_fallback(font_tuple):
        ext = str(PDFProcessor._font_tuple_value(font_tuple, 1) or "").strip().lower()
        if ext and ext not in {"n/a", "none", "null"}:
            return True, True
        return False, True

    @staticmethod
    def _collect_font_precheck_findings(doc):
        fonts = {}
        for page_index, page in enumerate(doc, start=1):
            try:
                page_fonts = page.get_fonts(full=True)
            except Exception:
                continue

            for font_tuple in page_fonts:
                original_name = str(PDFProcessor._font_tuple_value(font_tuple, 3) or "").strip()
                if not original_name:
                    original_name = str(PDFProcessor._font_tuple_value(font_tuple, 4) or "").strip()
                normalized_name = PDFProcessor._normalize_font_name(original_name)
                if not normalized_name:
                    continue

                try:
                    xref = int(PDFProcessor._font_tuple_value(font_tuple, 0) or 0)
                except Exception:
                    xref = 0
                embedded, known = PDFProcessor._font_object_has_embedded_file(doc, xref)
                if not embedded and known:
                    embedded, known = PDFProcessor._font_tuple_embedded_fallback(font_tuple)

                entry = fonts.setdefault(normalized_name, {
                    "font_name": original_name,
                    "original_names": set(),
                    "normalized_name": normalized_name,
                    "pages": set(),
                    "has_unembedded": False,
                    "has_embedded": False,
                    "embedding_unknown": False,
                    "base14": PDFProcessor._is_base14_font(normalized_name),
                })
                entry["original_names"].add(original_name)
                entry["pages"].add(page_index)
                if not known:
                    entry["embedding_unknown"] = True
                elif embedded:
                    entry["has_embedded"] = True
                else:
                    entry["has_unembedded"] = True

        findings = []
        for normalized_name, entry in sorted(fonts.items(), key=lambda item: item[0].lower()):
            embedding_status_known = not entry["embedding_unknown"]
            embedded = entry["has_embedded"] and not entry["has_unembedded"] and embedding_status_known
            if entry["has_unembedded"]:
                embedded = False
            substitution_risk = bool((not entry["base14"]) and entry["has_unembedded"] and embedding_status_known)

            if embedded and entry["base14"] and not entry["embedding_unknown"]:
                continue

            findings.append({
                "font_name": sorted(entry["original_names"])[0],
                "original_names": sorted(entry["original_names"]),
                "normalized_name": normalized_name,
                "pages": sorted(entry["pages"]),
                "embedded": embedded,
                "base14": entry["base14"],
                "substitution_risk": substitution_risk,
                "embedding_status_known": embedding_status_known,
            })

        unembedded_count = sum(1 for item in findings if item["embedding_status_known"] and not item["embedded"])
        non_base14_count = sum(1 for item in findings if not item["base14"])
        risk_count = sum(1 for item in findings if item["substitution_risk"])
        unknown_count = sum(1 for item in findings if not item["embedding_status_known"])

        summary_parts = []
        if unembedded_count:
            summary_parts.append(f"未嵌入字体 {unembedded_count} 个")
        if non_base14_count:
            summary_parts.append(f"非标准字体 {non_base14_count} 个")
        if risk_count:
            summary_parts.append(f"替代字体风险 {risk_count} 个")
        if unknown_count:
            summary_parts.append(f"嵌入状态未知字体 {unknown_count} 个")

        detail_parts = []
        for item in findings:
            labels = []
            if not item["embedding_status_known"]:
                labels.append("嵌入状态未知")
            elif item["embedded"]:
                labels.append("已嵌入")
            else:
                labels.append("未嵌入")
            labels.append("Base14" if item["base14"] else "非Base14")
            if item["substitution_risk"]:
                labels.append("替代风险")
            detail_parts.append(
                f"{item['normalized_name']}({PDFProcessor._format_font_page_numbers(item['pages'])}，{'，'.join(labels)})"
            )

        return {
            "font_summary": "，".join(summary_parts),
            "font_details": "; ".join(detail_parts),
            "font_findings": findings,
        }
```

- [ ] **Step 4: Run collection tests and verify pass**

Run: `python -m pytest test_pdf_processor_font_precheck.py::FontPrecheckCollectionTests -v`

Expected: PASS for all three collection tests.

- [ ] **Step 5: Run helper and collection tests together**

Run: `python -m pytest test_pdf_processor_font_precheck.py::FontPrecheckHelperTests test_pdf_processor_font_precheck.py::FontPrecheckCollectionTests -v`

Expected: PASS for all six tests.

---

### Task 3: Build Precheck Report Integration

**Files:**
- Modify: `test_pdf_processor_font_precheck.py`
- Modify: `pdf_processor.py`

- [ ] **Step 1: Add report integration test**

Append this test class to `test_pdf_processor_font_precheck.py` before the `if __name__ == "__main__"` block:

```python
class FontPrecheckReportIntegrationTests(unittest.TestCase):
    def test_build_precheck_report_adds_report_only_font_review(self):
        with tempfile.TemporaryDirectory() as td:
            src = os.path.join(td, "src.pdf")
            doc = fitz.open()
            try:
                doc.new_page(width=595, height=842)
                doc.save(src)
            finally:
                doc.close()

            font_result = {
                "font_summary": "未嵌入字体 1 个，非标准字体 1 个，替代字体风险 1 个",
                "font_details": "Calibri(第1页，未嵌入，非Base14，替代风险)",
                "font_findings": [{
                    "font_name": "ABCDEF+Calibri",
                    "normalized_name": "Calibri",
                    "pages": [1],
                    "embedded": False,
                    "base14": False,
                    "substitution_risk": True,
                    "embedding_status_known": True,
                }],
            }

            with patch.object(PDFProcessor, "_collect_font_precheck_findings", return_value=font_result):
                report = PDFProcessor.build_precheck_report(src)

        self.assertTrue(report.get("available"))
        self.assertEqual(font_result["font_summary"], report.get("font_summary"))
        self.assertEqual(font_result["font_details"], report.get("font_details"))
        self.assertEqual(font_result["font_findings"], report.get("font_findings"))
        suggestions = report.get("suggestions", {})
        self.assertIn("font_precheck_review", suggestions)
        self.assertTrue(suggestions["font_precheck_review"].get("report_only"))
        self.assertIn("字体预检", suggestions["font_precheck_review"].get("title", ""))
        self.assertIn("Calibri", suggestions["font_precheck_review"].get("reason", ""))
```

- [ ] **Step 2: Run report integration test and verify failure**

Run: `python -m pytest test_pdf_processor_font_precheck.py::FontPrecheckReportIntegrationTests -v`

Expected: FAIL because `build_precheck_report()` does not populate `font_summary`, `font_details`, `font_findings`, or `font_precheck_review`.

- [ ] **Step 3: Initialize font report fields**

In `pdf_processor.py`, update the `report` dictionary in `build_precheck_report()` to include font fields:

```python
        report = {
            "available": False,
            "file_path": input_path,
            "file_name": os.path.basename(input_path),
            "suggestions": {},
            "error": "",
            "font_summary": "",
            "font_details": "",
            "font_findings": [],
        }
```

- [ ] **Step 4: Add font collection to precheck report**

In `pdf_processor.py`, inside `build_precheck_report()` after the qpdf restriction suggestion block and before `return report`, insert:

```python
            font_precheck = PDFProcessor._collect_font_precheck_findings(doc)
            report["font_summary"] = font_precheck.get("font_summary", "")
            report["font_details"] = font_precheck.get("font_details", "")
            report["font_findings"] = font_precheck.get("font_findings", [])
            if report["font_summary"]:
                reason = report["font_summary"]
                if report["font_details"]:
                    reason = f"{reason}；明细：{report['font_details']}"
                PDFProcessor._add_precheck_report_finding(
                    suggestions,
                    "font_precheck_review",
                    "字体预检：需要复核",
                    reason,
                )
```

Keep this block after other structure checks so font precheck does not prevent existing actionable suggestions from being found.

- [ ] **Step 5: Run report integration test and verify pass**

Run: `python -m pytest test_pdf_processor_font_precheck.py::FontPrecheckReportIntegrationTests -v`

Expected: PASS.

- [ ] **Step 6: Run existing report-only tests**

Run: `python -m pytest test_pdf_processor_navigation.py::PrecheckReportOnlyTests -v`

Expected: PASS, confirming `report_only` behavior is unchanged.

---

### Task 4: Worker Rows And CSV Export

**Files:**
- Modify: `test_pdf_processor_font_precheck.py`
- Modify: `controller.py`

- [ ] **Step 1: Add worker propagation and CSV export tests**

Append these fake view classes and tests to `test_pdf_processor_font_precheck.py` before the `if __name__ == "__main__"` block:

```python
class _FakeExportEdit:
    def __init__(self, value):
        self._value = value

    def text(self):
        return self._value


class _FakeExportSettingsDialog:
    def __init__(self, default_dir):
        self.default_output_edit = _FakeExportEdit(default_dir)


class _FakeExportView:
    def __init__(self, default_dir):
        self.settings_dialog = _FakeExportSettingsDialog(default_dir)
        self.success_messages = []
        self.warning_messages = []
        self.error_messages = []

    def show_success_message(self, title, message):
        self.success_messages.append((title, message))

    def show_warning_message(self, title, message):
        self.warning_messages.append((title, message))

    def show_error_message(self, title, message):
        self.error_messages.append((title, message))


class FontPrecheckWorkerAndExportTests(unittest.TestCase):
    def test_precheck_worker_carries_font_fields_without_suggestion_ids(self):
        worker = PreCheckWorker(["src.pdf"])
        rows = []
        worker.result_ready.connect(rows.append)

        with patch.object(PDFProcessor, "build_precheck_report", return_value={
            "available": True,
            "font_summary": "未嵌入字体 1 个，非标准字体 1 个，替代字体风险 1 个",
            "font_details": "Calibri(第1页，未嵌入，非Base14，替代风险)",
            "suggestions": {
                "font_precheck_review": {
                    "matched": True,
                    "title": "字体预检：需要复核",
                    "reason": "未嵌入字体 1 个；明细：Calibri(第1页，未嵌入，非Base14，替代风险)",
                    "report_only": True,
                }
            },
        }):
            worker.run()

        self.assertEqual(1, len(rows))
        self.assertEqual("需要复核", rows[0].get("status"))
        self.assertEqual("", rows[0].get("suggestion_ids"))
        self.assertEqual("未嵌入字体 1 个，非标准字体 1 个，替代字体风险 1 个", rows[0].get("font_summary"))
        self.assertEqual("Calibri(第1页，未嵌入，非Base14，替代风险)", rows[0].get("font_details"))

    def test_export_precheck_results_includes_font_columns(self):
        with tempfile.TemporaryDirectory() as td:
            export_path = os.path.join(td, "precheck.csv")
            controller = MainController.__new__(MainController)
            controller.loaded_files = []
            controller.view = _FakeExportView(td)
            controller.last_precheck_results = [{
                "file_name": "src.pdf",
                "file_path": "src.pdf",
                "status": "需要复核",
                "suggestions": "字体预检：需要复核：未嵌入字体 1 个",
                "suggestion_ids": "",
                "error": "",
                "font_summary": "未嵌入字体 1 个，非标准字体 1 个，替代字体风险 1 个",
                "font_details": "Calibri(第1页，未嵌入，非Base14，替代风险)",
            }]

            with patch("controller.QFileDialog.getSaveFileName", return_value=(export_path, "CSV Files (*.csv)")):
                controller.export_precheck_results()

            with open(export_path, "r", encoding="utf-8-sig", newline="") as f:
                rows = list(csv.DictReader(f))

        self.assertEqual(1, len(rows))
        self.assertIn("font_summary", rows[0])
        self.assertIn("font_details", rows[0])
        self.assertEqual("未嵌入字体 1 个，非标准字体 1 个，替代字体风险 1 个", rows[0]["font_summary"])
        self.assertEqual("Calibri(第1页，未嵌入，非Base14，替代风险)", rows[0]["font_details"])
```

- [ ] **Step 2: Run worker/export tests and verify failure**

Run: `python -m pytest test_pdf_processor_font_precheck.py::FontPrecheckWorkerAndExportTests -v`

Expected: FAIL because worker rows and CSV fieldnames do not yet include `font_summary` and `font_details`.

- [ ] **Step 3: Propagate font fields in `PreCheckWorker`**

In `controller.py`, inside `PreCheckWorker.run()`, update the `result_ready.emit()` payload for the suggestions branch to include:

```python
                        "font_summary": report.get("font_summary", ""),
                        "font_details": report.get("font_details", ""),
```

The full payload in the suggestions branch becomes:

```python
                    self.result_ready.emit({
                        "file_name": base_name,
                        "file_path": file_path,
                        "status": status,
                        "suggestions": advice,
                        "suggestion_ids": suggestion_ids,
                        "error": "",
                        "font_summary": report.get("font_summary", ""),
                        "font_details": report.get("font_details", ""),
                    })
```

Update the no-suggestions branch payload to include empty or report-provided font fields:

```python
                    self.result_ready.emit({
                        "file_name": base_name,
                        "file_path": file_path,
                        "status": "无需处理",
                        "suggestions": "",
                        "suggestion_ids": "",
                        "error": "",
                        "font_summary": report.get("font_summary", ""),
                        "font_details": report.get("font_details", ""),
                    })
```

Update the failure payload to keep CSV rows consistent:

```python
                    self.result_ready.emit({
                        "file_name": base_name,
                        "file_path": file_path,
                        "status": "预检失败",
                        "suggestions": "",
                        "suggestion_ids": "",
                        "error": reason,
                        "font_summary": "",
                        "font_details": "",
                    })
```

- [ ] **Step 4: Add font fields to precheck CSV export**

In `controller.py`, in `export_precheck_results()`, replace the `fieldnames` list with:

```python
            fieldnames = [
                "file_name",
                "file_path",
                "status",
                "suggestions",
                "suggestion_ids",
                "error",
                "font_summary",
                "font_details",
            ]
```

Before `writer.writerows(self.last_precheck_results)`, normalize rows so older rows without font fields export cleanly:

```python
                rows = []
                for row in self.last_precheck_results:
                    export_row = {key: row.get(key, "") for key in fieldnames}
                    rows.append(export_row)
                writer.writerows(rows)
```

- [ ] **Step 5: Run worker/export tests and verify pass**

Run: `python -m pytest test_pdf_processor_font_precheck.py::FontPrecheckWorkerAndExportTests -v`

Expected: PASS.

- [ ] **Step 6: Run existing report-only tests again**

Run: `python -m pytest test_pdf_processor_navigation.py::PrecheckReportOnlyTests -v`

Expected: PASS, confirming added row fields did not change actionability behavior.

---

### Task 5: File Details And Documentation

**Files:**
- Modify: `test_pdf_processor_font_precheck.py`
- Modify: `controller.py`
- Modify: `README.md`

- [ ] **Step 1: Add file detail test for report-only reasons**

Append this test class to `test_pdf_processor_font_precheck.py` before the `if __name__ == "__main__"` block:

```python
class FontPrecheckFileDetailsTests(unittest.TestCase):
    def test_file_details_include_report_only_reason(self):
        with tempfile.TemporaryDirectory() as td:
            src = os.path.join(td, "src.pdf")
            doc = fitz.open()
            try:
                doc.new_page(width=595, height=842)
                doc.save(src)
            finally:
                doc.close()

            controller = MainController.__new__(MainController)
            report = {
                "available": True,
                "suggestions": {
                    "font_precheck_review": {
                        "matched": True,
                        "title": "字体预检：需要复核",
                        "reason": "未嵌入字体 1 个；明细：Calibri(第1页，未嵌入，非Base14，替代风险)",
                        "report_only": True,
                    }
                },
            }

            with patch.object(PDFProcessor, "build_precheck_report", return_value=report):
                details = controller._build_pdf_detail_text(src)

        self.assertIn("字体预检：需要复核", details)
        self.assertIn("Calibri", details)
        self.assertIn("替代风险", details)
```

- [ ] **Step 2: Run file detail test and verify failure**

Run: `python -m pytest test_pdf_processor_font_precheck.py::FontPrecheckFileDetailsTests -v`

Expected: FAIL because `_build_pdf_detail_text()` currently lists suggestion titles only and omits report-only reasons.

- [ ] **Step 3: Include report-only reasons in file details**

In `controller.py`, in `_build_pdf_detail_text()`, replace:

```python
            suggestions = [item.get("title", "") for item in report.get("suggestions", {}).values() if item.get("title")]
```

with:

```python
            suggestions = []
            for item in report.get("suggestions", {}).values():
                title = item.get("title", "")
                if not title:
                    continue
                reason = item.get("reason", "")
                if item.get("report_only") and reason:
                    suggestions.append(f"{title}：{reason}")
                else:
                    suggestions.append(title)
```

- [ ] **Step 4: Update README feature text**

In `README.md`, update the precheck feature bullets.

At the project features section near the existing precheck bullet, keep the existing bullet and do not add a duplicate.

In `### 8. 批量预检`, replace:

```markdown
- 可检查 PDF 版本、线性化状态与权限限制信号
```

with:

```markdown
- 可检查 PDF 版本、线性化状态、权限限制信号与字体嵌入风险
```

After that bullet, add:

```markdown
- 字体预检会报告未嵌入字体、非 PDF Base 14 字体与替代字体风险，仅作为人工复核信息，不会自动勾选处理规则
```

In `### 2. 从源码运行`, after the qpdf capability list, add:

```markdown
字体预检基于 PyMuPDF 读取 PDF 内部字体资源，不依赖 qpdf，也不会嵌入或替换字体。
```

- [ ] **Step 5: Run file detail test and verify pass**

Run: `python -m pytest test_pdf_processor_font_precheck.py::FontPrecheckFileDetailsTests -v`

Expected: PASS.

- [ ] **Step 6: Run disabled font embedding tests**

Run: `python -m pytest test_gs_removal.py -v`

Expected: PASS, confirming `embed_nonstandard_fonts` remains disabled and blocked.

---

### Task 6: Focused Regression Suite

**Files:**
- Modify only if failures point to code changed in Tasks 1-5.

- [ ] **Step 1: Run the new font precheck suite**

Run: `python -m pytest test_pdf_processor_font_precheck.py -v`

Expected: PASS.

- [ ] **Step 2: Run adjacent precheck and disabled-feature suites**

Run: `python -m pytest test_pdf_processor_navigation.py::PrecheckReportOnlyTests test_gs_removal.py -v`

Expected: PASS.

- [ ] **Step 3: Run broader processor-related tests**

Run: `python -m pytest test_pdf_processor_navigation.py test_pdf_processor_tools.py test_pdf_processor_metadata.py -v`

Expected: PASS. If a qpdf-dependent test fails because qpdf is unavailable, record the exact failing test and stderr, then run the font and report-only focused suites again to verify the font work.

- [ ] **Step 4: Inspect final diff**

Run: `git diff -- pdf_processor.py controller.py README.md test_pdf_processor_font_precheck.py docs/superpowers/specs/2026-06-06-font-precheck-design.md docs/superpowers/plans/2026-06-06-font-precheck.md`

Expected: Diff is limited to font precheck implementation, tests, README, and planning docs.

- [ ] **Step 5: Check worktree status**

Run: `git status --short`

Expected: Shows only intended files. Do not commit unless the user explicitly asks for a commit.

---

## Self-Review Checklist

- Spec coverage: Tasks 1-3 implement Base 14 detection, unembedded detection, non-standard detection, substitution-risk detection, and `report_only` integration.
- Spec coverage: Task 4 implements `font_summary` and `font_details` propagation and CSV export.
- Spec coverage: Task 5 keeps existing UI surfaces, improves file details, updates README, and preserves disabled embedding behavior.
- Spec coverage: Task 6 verifies focused and adjacent regressions.
- Type consistency: The plan uses `font_summary`, `font_details`, `font_findings`, `embedded`, `base14`, `substitution_risk`, and `embedding_status_known` consistently across tests and implementation.
- Version control: This plan intentionally avoids commit steps because repository instructions require explicit user approval before committing.
