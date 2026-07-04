# Favorite Rule Combo Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add one user-saved favorite processing rule combination.

**Architecture:** Reuse `MainWindow` checkbox state and existing `QSettings` storage. Add a checkable `我的常用` preset button and a regular `保存为常用` action in the preset bar; favorite options are stored as a string list under `Presets/FavoriteOptions`.

**Tech Stack:** Python, PySide6, QSettings, unittest/pytest.

---

## File Structure

- Modify `view.py`: add favorite preset UI, storage helpers, apply/save behavior, and preset summary integration.
- Modify `controller.py`: wire the save favorite button.
- Modify `tests/test_regressions.py`: add tests for saving/applying favorite combinations and missing favorite prompts.

### Task 1: Add Favorite Preset UI

- [ ] Write tests verifying `btn_preset_favorite` and `btn_save_favorite_preset` exist.
- [ ] Run focused test and confirm it fails.
- [ ] Add both buttons to the existing preset bar in `view.py`.
- [ ] Add favorite button to `preset_btn_group` so it is mutually exclusive with China/US presets.
- [ ] Re-run focused test and confirm it passes.

### Task 2: Save Favorite Combination

- [ ] Write a test selecting processing rules, calling `save_favorite_preset`, and verifying `Presets/FavoriteOptions` stores only processing rule IDs.
- [ ] Run focused test and confirm it fails.
- [ ] Implement `get_processing_options`, `get_favorite_preset_options`, and `save_favorite_preset` in `view.py`.
- [ ] Save favorite options as a list sorted by current module order, excluding global settings.
- [ ] Re-run focused test and confirm it passes.

### Task 3: Apply Favorite Combination

- [ ] Write a test with saved favorite options that calls `toggle_preset("favorite")` and verifies only those rules are checked.
- [ ] Run focused test and confirm it fails.
- [ ] Extend `apply_preset`, `_set_preset_button_state`, `toggle_preset`, and summary text to support `favorite`.
- [ ] If no favorite is saved, clicking `我的常用` shows a warning and leaves current rules unchanged.
- [ ] Re-run focused test and confirm it passes.

### Task 4: Regression Verification

- [ ] Run `python -m pytest tests/test_regressions.py -v`.
- [ ] Inspect `git diff -- view.py controller.py tests/test_regressions.py docs/superpowers/plans/2026-05-29-favorite-rule-combo.md`.

## Self-Review

- Spec coverage: single favorite combination, save/apply buttons, QSettings persistence, global setting exclusion, no multi-profile management.
- Placeholder scan: no incomplete placeholders.
- Type consistency: favorite options are Python lists of option IDs and QSettings string lists.
