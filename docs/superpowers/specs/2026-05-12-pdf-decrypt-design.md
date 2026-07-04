# PDF Decrypt Feature Design

Date: 2026-05-12
Project: `RATools-for-PDF`
Topic: Add PDF permission-restriction removal using bundled `qpdf`

## Summary

This feature adds a new batch-processing option that removes PDF permission restrictions by producing an unencrypted output copy through `qpdf`.

The first version only targets PDFs that can be opened normally but have restrictions on printing, copying, or editing. It does not support prompting for passwords, and it does not attempt to decrypt PDFs that require an open password.

## Goals

- Add a user-visible option for removing PDF permission restrictions.
- Reuse the existing batch-processing flow, logging, overwrite behavior, and output directory behavior.
- Allow this option to be combined with existing `qpdf`-backed output options such as PDF version conversion and linearization.
- Keep the implementation small and aligned with the current structure of `view.py`, `controller.py`, and `pdf_processor.py`.

## Non-Goals

- Prompting for owner passwords.
- Prompting for open passwords.
- Showing detailed encryption metadata in the UI.
- Adding a separate standalone decrypt tool window or workflow.
- Detecting and reporting PDF security settings outside the normal processing flow.

## User Experience

### Entry Point

Add a new option under the existing `文件级优化与输出` section in the right-side rules panel.

- `id`: `remove_pdf_restrictions`
- `title`: `PDF解除权限限制`
- `desc`: `尝试移除禁止复制、打印、编辑等权限限制，不处理需要打开密码的加密文档`

This option behaves like the existing `PDF版本转换` and `启用线性化 (快速网页浏览)` options. Users select files, choose the rule, and run the normal batch process.

### Behavior

- If the input PDF can be processed without a password, the output file is written without permission restrictions.
- If the input PDF requires a password or cannot be decrypted without credentials, that file fails and the batch continues.
- If overwrite mode is enabled, the original file is only replaced after a successful output has been produced.

### Failure Messaging

Failure messages should be specific instead of only surfacing raw `qpdf` errors.

Preferred user-facing failure messages:

- `该PDF需要密码，当前模式不支持输入密码解锁`
- `未能移除PDF权限限制：<qpdf错误摘要>`

Raw stderr may still be included in logs when useful, but the top-level message should clearly explain whether the issue is password-related or a general `qpdf` failure.

## Architecture

### `view.py`

Add one new checkbox option to the existing `文件级优化与输出` module data.

No new dialogs, buttons, panels, or separate workflows are required.

### `controller.py`

No structural changes are needed.

The existing processing entry point already collects selected options and passes them through the standard `ProcessWorker` into `PDFProcessor.process_document(...)`. The new option should flow through this existing path unchanged.

### `pdf_processor.py`

This is the only file that needs meaningful logic changes.

The current `qpdf` integration already supports:

- locating the executable
- rewriting a PDF
- optionally forcing a PDF version
- optionally linearizing output

Extend the `qpdf` rewrite path so it can also remove encryption and permission restrictions in the same invocation.

Recommended change:

- extend `_rewrite_with_qpdf(...)` with a new boolean flag such as `decrypt_restrictions=False`
- when enabled, include the appropriate `qpdf` argument to produce an unencrypted output
- keep version conversion and linearization in the same `qpdf` command when selected

This avoids multiple rewrite passes and keeps file handling simple.

## Processing Flow

Within `PDFProcessor.process_document(...)`:

1. Collect the existing output-related flags:
   - version conversion
   - linearization
   - permission-restriction removal
2. Determine whether a `qpdf` rewrite is needed.
3. If needed, perform one `qpdf` invocation that combines all selected output transformations.
4. Preserve the existing save and temp-file behavior as much as possible.

The new option should be treated as another file-level output transform, not as a separate pre-processing or post-processing stage.

## Option Interactions

The new option must be compatible with:

- `convert_pdf_version`
- `fast_web_view`

Expected combined behavior:

- `remove_pdf_restrictions` alone: output is rewritten without permission restrictions
- `remove_pdf_restrictions` + `convert_pdf_version`: output is unencrypted and forced to the target version
- `remove_pdf_restrictions` + `fast_web_view`: output is unencrypted and linearized
- all three together: one `qpdf` run produces an unencrypted, version-adjusted, linearized file

No new exclusivity rules are needed in the controller.

## Error Handling

### Password-Protected PDFs

If the file requires credentials to open or decrypt:

- mark the file as failed
- do not prompt for a password
- continue processing the rest of the batch

The message should explain that password input is not supported in the current mode.

### `qpdf` Execution Failures

If `qpdf` is missing or execution fails:

- preserve the current failure behavior used by other `qpdf`-backed features
- convert the error into a user-facing message that is actionable
- do not overwrite the source file

### Partial Batch Failure

One file failing to remove restrictions must not abort the whole batch. This should remain consistent with the current processing model.

## Testing Scope

Minimum validation cases:

1. Plain PDF with no encryption
   - selecting `remove_pdf_restrictions` still produces a valid output
2. PDF with permission restrictions that can be removed without prompting
   - processing succeeds and output is usable
3. PDF that requires a password
   - file fails with the expected message
   - remaining files continue processing
4. Combined with `convert_pdf_version`
   - output remains valid and version rewrite still works
5. Combined with `fast_web_view`
   - output remains valid and linearization still works
6. Overwrite-original mode
   - source file is only replaced on success

If automated tests are added later, they should focus on the `pdf_processor.py` `qpdf` invocation path and error handling. For this first implementation, manual verification is acceptable if test fixtures for encrypted PDFs are not already present.

## Risks And Constraints

- `qpdf` behavior depends on the exact encryption state of the input PDF. Some files that appear to only have restrictions may still require credentials to remove them.
- Error text from `qpdf` may vary by version, so password-related error classification should be tolerant and not depend on a single exact stderr string.
- The feature should not silently claim success if `qpdf` rewrites the file but restrictions remain. Verification should rely on `qpdf` succeeding in unencrypted output mode.

## Recommended Implementation Shape

Keep the implementation minimal:

- one new option entry in `view.py`
- no UI workflow changes in `controller.py`
- one targeted enhancement to the existing `qpdf` rewrite helper and its caller in `pdf_processor.py`

This preserves the current design of the application while making practical use of the already-bundled `qpdf` tool.
