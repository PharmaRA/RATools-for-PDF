# Batch Result Summary Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Show a concise batch result summary in the processing completion dialog and append it to logs.

**Architecture:** `MainController` owns batch counters for processing runs. Counters reset in `start_processing`, update from `update_progress` statuses, and render through a helper used by `processing_finished`.

**Tech Stack:** Python, PySide6, unittest/pytest.

---

## File Structure

- Modify `controller.py`: add batch counter state, counter updates, summary rendering, and completion dialog message changes.
- Modify `tests/test_regressions.py`: add focused tests for counter updates and completion summary content.

### Task 1: Track Batch Counts

- [ ] Add a failing test that starts with empty `batch_result_counts`, feeds `处理完成`, `处理失败`, and `已跳过` through `update_progress`, and asserts counts.
- [ ] Run the focused test and confirm it fails.
- [ ] Add `self.batch_result_counts = {"success": 0, "failure": 0, "skip": 0}` in `MainController.__init__`.
- [ ] Reset counters in `start_processing`.
- [ ] Increment counters in `update_progress` only once per file when a terminal processing status arrives.
- [ ] Re-run the focused test and confirm it passes.

### Task 2: Render Completion Summary

- [ ] Add a failing test that sets processing total, counters, failed files, and started time, then asserts `_build_batch_result_summary` includes total, success, failure, skip, elapsed seconds, and failed-item retry hint.
- [ ] Run the focused test and confirm it fails.
- [ ] Implement `_build_batch_result_summary(summary)` in `MainController`.
- [ ] Use Chinese labels: `总数`、`成功`、`失败`、`跳过`、`总耗时`、`失败项`.
- [ ] Re-run the focused test and confirm it passes.

### Task 3: Use Summary In Dialog And Logs

- [ ] Add a failing test patching `show_success_message` and calling `processing_finished`, asserting the dialog message includes batch summary details.
- [ ] Run the focused test and confirm it fails.
- [ ] In `processing_finished`, build the summary before resetting processing state.
- [ ] Append the rendered summary to `process_logs` and pass it to `show_success_message`/`show_info_message`.
- [ ] Re-run the focused test and confirm it passes.

### Task 4: Regression Verification

- [ ] Run `python -m pytest tests/test_regressions.py -v`.
- [ ] Inspect `git diff -- controller.py tests/test_regressions.py docs/superpowers/plans/2026-05-29-batch-result-summary.md`.

## Self-Review

- Spec coverage: completion dialog, log summary, counts from status events, processing runs only.
- Placeholder scan: no placeholders.
- Type consistency: counters are `success`, `failure`, and `skip` integers.
