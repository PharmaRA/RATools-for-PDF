# Font Precheck Design

## Context

`embed_nonstandard_fonts` is currently shown in the UI as temporarily unavailable and is blocked when processing starts. The first font-related capability will be a read-only font precheck that reports risk signals before any embedding implementation is considered.

The existing precheck flow already supports automatic suggestions and `report_only` findings. Font precheck will use that flow so it can report risks without enabling or auto-selecting any processing rule.

## Goals

- Detect unembedded fonts.
- Detect non-standard fonts using the PDF Base 14 definition.
- Detect substitute-font risk using the rule: non-Base-14 and unembedded.
- Show a file-level summary plus font-level details.
- Keep font findings as review-only output, never as auto-applicable suggestions.
- Preserve the current disabled state and processing block for `embed_nonstandard_fonts`.

## Non-Goals

- Do not embed fonts in this phase.
- Do not introduce Ghostscript, qpdf, Acrobat Preflight, Poppler, or any new external font engine in this phase.
- Do not infer whether the current user's system can actually substitute or render a missing font.
- Do not add a new standalone font inspection page in the first version.

## Standard Font Definition

Standard fonts are the PDF Base 14 fonts and their bold, italic, and bold-italic variants:

- `Courier`
- `Courier-Bold`
- `Courier-Oblique`
- `Courier-BoldOblique`
- `Helvetica`
- `Helvetica-Bold`
- `Helvetica-Oblique`
- `Helvetica-BoldOblique`
- `Times-Roman`
- `Times-Bold`
- `Times-Italic`
- `Times-BoldItalic`
- `Symbol`
- `ZapfDingbats`

Subset font prefixes are ignored for classification. For example, `ABCDEF+Calibri` is classified as `Calibri`, while the original font name is retained in the detail output.

## Architecture

Font precheck is implemented inside the existing `PDFProcessor.build_precheck_report(input_path)` flow. A focused helper such as `_collect_font_precheck_findings(doc)` will scan the already-open PyMuPDF document and return structured font results.

The helper is read-only. It does not save the document, rewrite objects, call qpdf, or modify the `options` model.

The top-level precheck report will gain optional fields:

```python
{
    "font_summary": "发现 2 个未嵌入非标准字体，存在替代字体风险",
    "font_details": "Calibri(第1-3页，未嵌入，非Base14，替代风险); SimSun(第5页，已嵌入，非Base14)",
    "font_findings": [
        {
            "font_name": "ABCDEF+Calibri",
            "normalized_name": "Calibri",
            "pages": [1, 2, 3],
            "embedded": False,
            "base14": False,
            "substitution_risk": True,
            "embedding_status_known": True,
        }
    ]
}
```

`font_findings` is for internal structured handling. `font_summary` and `font_details` are persisted into precheck rows and CSV export.

## Detection Flow

For each PDF page, the helper collects font resources through PyMuPDF's page font APIs and aggregates them by normalized font name.

For each aggregated font:

- Store all original names encountered.
- Store the sorted one-based page numbers where the font appears.
- Compute `base14` from the normalized font name.
- Compute `embedded` by inspecting font objects and font descriptors for `/FontFile`, `/FontFile2`, or `/FontFile3` where available.
- Use PyMuPDF extension information as a fallback signal when direct object inspection is unavailable.
- Set `embedding_status_known` to `False` if the embedded state cannot be determined safely.
- Set `substitution_risk` to `True` only when `base14` is `False`, `embedded` is `False`, and `embedding_status_known` is `True`.

The helper produces review findings when any of these signals exist:

- At least one unembedded font.
- At least one non-Base-14 font.
- At least one substitute-font risk.
- At least one font with unknown embedding status.

## Precheck Integration

Font findings are added through `_add_precheck_report_finding()` and marked as `report_only`.

This means:

- A file with only font findings receives `需要复核`.
- A file with font findings plus actionable findings can still receive `建议处理`.
- Font findings never appear in `suggestion_ids`.
- `apply_precheck_suggestions()` continues to auto-select only actionable processing rules.
- `start_precheck_suggested_processing()` continues to process only files with actionable `suggestion_ids`.

The existing disabled `embed_nonstandard_fonts` behavior remains unchanged:

- The checkbox remains disabled and labeled as temporarily unavailable.
- Processing still blocks if the option is somehow present in selected options.

## Display And Export

The first version reuses existing UI surfaces.

The queue status column remains limited to the current status values: `建议处理`, `需要复核`, `无需处理`, and `预检失败`.

The precheck log includes the font summary and compressed detail text. Example:

```text
字体预检：未嵌入字体 2 个，非标准字体 3 个，替代字体风险 2 个；明细：Calibri(第1-3页，未嵌入，非Base14，替代风险); SimSun(第5页，已嵌入，非Base14)
```

`last_precheck_results` gains two exported fields:

- `font_summary`
- `font_details`

The precheck CSV export adds the same two columns while preserving the existing columns:

- `file_name`
- `file_path`
- `status`
- `suggestions`
- `suggestion_ids`
- `error`

The existing "View file details" action uses `build_precheck_report()` and will include font review findings in its existing recommendation section. No new dialog is added in the first version.

## Error Handling

Font precheck follows the current precheck principle: read-only inspection should not interrupt the full batch unless the whole file cannot be inspected.

- Password-protected PDFs keep the existing behavior: precheck fails for that file and does not scan fonts.
- If one font object cannot be parsed, the file precheck remains available and the font result includes an unknown embedding-status review finding.
- If the helper cannot determine embedding status for a font, it does not mark substitute-font risk for that font.
- Duplicate font resources are deduplicated by normalized name and page numbers.
- Font precheck exceptions do not generate actionable `suggestion_ids`.

## Testing

Tests will cover:

- Base 14 family and variant recognition.
- Subset prefix normalization, such as `ABCDEF+Calibri` to `Calibri`.
- Non-Base-14 unembedded fonts generating substitute-font risk.
- Non-Base-14 embedded fonts reporting non-standard font status without substitute-font risk.
- Font findings being `report_only` and excluded from `suggestion_ids`.
- Precheck CSV export including `font_summary` and `font_details`.
- Existing disabled-feature tests still passing for `embed_nonstandard_fonts`.

## Future Extension

After the read-only precheck is stable, a later font embedding phase can evaluate external engines separately. That later design should compare accuracy, licensing, packaging impact, path discovery, and failure fallback for options such as Ghostscript, Poppler `pdffonts`, Acrobat Preflight, or other font-capable PDF engines.
