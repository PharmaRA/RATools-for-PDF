# Project Structure Refactor Phase 1 and 2 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Turn the current flat script-style layout into a package-based desktop application structure without changing user-visible behavior.

**Architecture:** Keep `main.py` and `main_no_update.py` at the repository root so PyInstaller, Nuitka, and existing user habits stay stable. Move implementation modules into a `ratools_pdf` package behind compatibility shims, then split the three oversized modules (`controller.py`, `view.py`, `pdf_processor.py`) by responsibility in a second phase.

**Tech Stack:** Python 3, PySide6, PyMuPDF (`fitz`), Windows batch build scripts, `unittest`.

---

## Current Structure Problem

The repository currently has a flat application layout:

- `controller.py` is a UI controller, worker host, process orchestration module, IO action helper module, update flow, log export helper, and file-list manager.
- `view.py` contains platform helpers, the frameless dialog base class, all dialogs, custom widgets, settings persistence, and the main window.
- `pdf_processor.py` contains qpdf integration, precheck analysis, font checks, bookmark/link import-export, page resizing, hyperlink style changes, and the full document processing pipeline.
- Smaller modules such as `app_paths.py`, `app_version.py`, `update_checker.py`, and `log_view_model.py` are also flat in the root.

This is workable for an early PySide tool, but it now makes file discovery, ownership, and safe refactoring harder.

## Non-Goals

- Do not change user-facing behavior.
- Do not change the generated executable names.
- Do not move `main.py`, `main_no_update.py`, `build_pyinstaller.bat`, or `build_nuitka.bat` out of the repository root.
- Do not adopt a `src/` layout in these two phases.
- Do not rewrite PDF processing algorithms while moving files.
- Do not remove root-level compatibility modules until a later cleanup phase.
- Do not start the structural refactor from a base branch that is missing the active `feature/ui-redesign` work.

## Phase Gates

Phase 0 is complete when:

- `feature/ui-redesign` has no uncommitted work.
- `main` contains the centralized theme system from that branch: `theme.py`, `tests/test_theme.py`, and the updated `view.py`.
- The full test suite passes after the UI redesign merge.

Phase 1 is complete when:

- `ratools_pdf/` exists and owns the small, low-risk modules.
- Root modules still import successfully as compatibility shims.
- `main.py`, `main_no_update.py`, build scripts, and the full test suite still work.

Phase 2 is complete when:

- The three oversized modules are reduced to compatibility shims or small facade modules.
- Controller helpers, controller workers, UI dialogs/widgets, and PDF processing helpers live in focused package modules.
- Existing tests pass, and new package-structure tests cover the new import paths.

---

## Known Parallel Branch: `feature/ui-redesign`

The `feature/ui-redesign` branch is currently checked out in `.worktrees/ui-redesign`. Its committed branch tip is already an ancestor of `main`, but that worktree has active uncommitted UI work:

```text
 M view.py
?? theme.py
?? tests/test_theme.py
```

That work introduces a centralized theme system and removes large local QSS blocks from `view.py`. Because Phase 2 will split `view.py`, the correct order is:

1. Finish and merge `feature/ui-redesign`.
2. Run Phase 1 from the post-redesign `main`.
3. Run Phase 2 after Phase 1 is green.

Phase 1 must avoid unnecessary edits to `view.py`. Root compatibility shims are enough to keep `view.py` working while helper modules move into `ratools_pdf/`. Phase 2 should split the post-redesign `view.py`, not the older pre-theme version.

## Target Structure After Phase 1

```text
RATools-for-PDF/
  main.py
  main_no_update.py
  app_features.py                  # compatibility shim
  app_paths.py                     # compatibility shim
  app_version.py                   # compatibility shim
  font_embedding_providers.py      # compatibility shim
  log_view_model.py                # compatibility shim
  theme.py                         # compatibility shim from UI redesign
  update_checker.py                # compatibility shim
  controller.py
  pdf_processor.py
  view.py
  ratools_pdf/
    __init__.py
    app.py
    config/
      __init__.py
      features.py
      paths.py
      version.py
    pdf/
      __init__.py
      font_embedding_providers.py
    services/
      __init__.py
      update_checker.py
    ui/
      __init__.py
      log_view_model.py
      theme.py
  tests/
    test_package_structure.py
    test_theme.py
```

## Target Structure After Phase 2

```text
RATools-for-PDF/
  main.py
  main_no_update.py
  controller.py                    # compatibility shim
  pdf_processor.py                 # compatibility shim
  theme.py                         # compatibility shim
  view.py                          # compatibility shim
  ratools_pdf/
    __init__.py
    app.py
    config/
      features.py
      paths.py
      version.py
    controllers/
      __init__.py
      io_actions.py
      log_export.py
      workers.py
      main_controller.py
    pdf/
      __init__.py
      font_embedding_providers.py
      processor.py
      qpdf.py
      precheck.py
      bookmarks_links.py
      page_layout.py
      hyperlink_styles.py
    services/
      update_checker.py
    ui/
      __init__.py
      platform.py
      widgets.py
      dialogs.py
      main_window.py
      log_view_model.py
      theme.py
```

---

# Phase 0: Finish and Merge UI Redesign

## Phase 0 File Map

- Modify: `.worktrees/ui-redesign/view.py`
  Responsibility: final centralized-theme UI changes from the redesign branch.
- Create: `.worktrees/ui-redesign/theme.py`
  Responsibility: centralized palette, QSS, and `ThemeManager`.
- Create: `.worktrees/ui-redesign/tests/test_theme.py`
  Responsibility: regression coverage for saved dark-theme startup behavior.

## Phase 0 Task 1: Stabilize and Merge `feature/ui-redesign`

**Files:**
- Modify: `.worktrees/ui-redesign/view.py`
- Create: `.worktrees/ui-redesign/theme.py`
- Create: `.worktrees/ui-redesign/tests/test_theme.py`

- [ ] **Step 1: Confirm the UI redesign worktree status**

Run:

```powershell
git -C .worktrees\ui-redesign status --short
```

Expected:

```text
 M view.py
?? tests/test_theme.py
?? theme.py
```

- [ ] **Step 2: Run the focused theme regression test**

Run:

```powershell
python -m unittest discover -s .worktrees\ui-redesign\tests -t .worktrees\ui-redesign -p "test_theme.py" -v
```

Expected: PASS.

- [ ] **Step 3: Run the full test suite in the UI redesign worktree**

Run:

```powershell
python -m unittest discover -s .worktrees\ui-redesign\tests -t .worktrees\ui-redesign -p "test_*.py" -v
```

Expected: PASS.

- [ ] **Step 4: Commit the UI redesign work**

Run:

```powershell
git -C .worktrees\ui-redesign add view.py theme.py tests/test_theme.py
git -C .worktrees\ui-redesign commit -m "feat: add centralized UI theme system"
```

Expected: commit succeeds on `feature/ui-redesign`.

- [ ] **Step 5: Rebase the UI redesign branch onto current `main`**

Run:

```powershell
git -C .worktrees\ui-redesign rebase main
```

Expected: rebase succeeds. If conflicts occur, resolve them in favor of preserving the centralized `theme.py` system and the latest `main` docs/tests.

- [ ] **Step 6: Run full tests again after the rebase**

Run:

```powershell
python -m unittest discover -s .worktrees\ui-redesign\tests -t .worktrees\ui-redesign -p "test_*.py" -v
```

Expected: PASS.

- [ ] **Step 7: Merge the UI redesign branch into `main`**

Run this from the main worktree after the plan document is committed or otherwise safely saved:

```powershell
git merge --ff-only feature/ui-redesign
```

Expected: fast-forward succeeds, and `main` now contains `theme.py`, `tests/test_theme.py`, and the redesigned `view.py`.

- [ ] **Step 8: Verify `main` after the merge**

Run:

```powershell
python -m unittest discover -s tests -p "test_*.py" -v
```

Expected: PASS.

- [ ] **Step 9: Commit state**

No extra commit is needed if Step 7 fast-forwarded the UI redesign commit into `main`. If the merge required a merge commit, use:

```powershell
git commit -m "merge: integrate UI redesign"
```

# Phase 1: Establish the Package Boundary

## Phase 1 File Map

- Create: `ratools_pdf/__init__.py`
  Responsibility: package marker and package metadata surface.
- Create: `ratools_pdf/app.py`
  Responsibility: runtime setup and QApplication startup currently in `main.py`.
- Create: `ratools_pdf/config/__init__.py`
  Responsibility: config package marker.
- Create: `ratools_pdf/config/features.py`
  Responsibility: update-check feature flag currently in `app_features.py`.
- Create: `ratools_pdf/config/paths.py`
  Responsibility: app/resource path helpers currently in `app_paths.py`.
- Create: `ratools_pdf/config/version.py`
  Responsibility: application version constants currently in `app_version.py`.
- Create: `ratools_pdf/pdf/__init__.py`
  Responsibility: PDF package marker.
- Create: `ratools_pdf/pdf/font_embedding_providers.py`
  Responsibility: font embedding provider implementations currently in `font_embedding_providers.py`.
- Create: `ratools_pdf/services/__init__.py`
  Responsibility: service package marker.
- Create: `ratools_pdf/services/update_checker.py`
  Responsibility: GitHub release update checks currently in `update_checker.py`.
- Create: `ratools_pdf/ui/__init__.py`
  Responsibility: UI package marker.
- Create: `ratools_pdf/ui/log_view_model.py`
  Responsibility: log summary parsing currently in `log_view_model.py`.
- Create: `ratools_pdf/ui/theme.py`
  Responsibility: centralized theme system introduced by `feature/ui-redesign`.
- Modify: `main.py`
  Responsibility: root executable entry point that delegates to `ratools_pdf.app.run`.
- Modify: `app_features.py`
  Responsibility: compatibility re-export for old imports.
- Modify: `app_paths.py`
  Responsibility: compatibility re-export for old imports.
- Modify: `app_version.py`
  Responsibility: compatibility re-export for old imports.
- Modify: `font_embedding_providers.py`
  Responsibility: compatibility re-export for old imports.
- Modify: `log_view_model.py`
  Responsibility: compatibility re-export for old imports.
- Modify: `theme.py`
  Responsibility: compatibility re-export for old UI redesign imports.
- Modify: `update_checker.py`
  Responsibility: compatibility re-export for old imports.
- Modify: `controller.py`
  Responsibility: import package modules where low-risk.
- Modify: `pdf_processor.py`
  Responsibility: import package modules where low-risk.
- Leave unchanged in Phase 1: `view.py`
  Responsibility: preserve the post-redesign UI file until the Phase 2 UI split.
- Create: `tests/test_package_structure.py`
  Responsibility: prove root compatibility imports and package imports both work.
- Modify: `tests/test_theme.py`
  Responsibility: keep UI theme regression coverage green after moving `theme.py`.
- Modify: `README.md`
  Responsibility: update the project structure section after the code move.

## Phase 1 Task 1: Create Package Directories

**Files:**
- Create: `ratools_pdf/__init__.py`
- Create: `ratools_pdf/config/__init__.py`
- Create: `ratools_pdf/pdf/__init__.py`
- Create: `ratools_pdf/services/__init__.py`
- Create: `ratools_pdf/ui/__init__.py`

- [ ] **Step 1: Create the package directories**

Run:

```powershell
New-Item -ItemType Directory -Force ratools_pdf, ratools_pdf\config, ratools_pdf\pdf, ratools_pdf\services, ratools_pdf\ui
```

Expected: PowerShell reports the directories, or exits successfully if they already exist.

- [ ] **Step 2: Add package marker files**

Create `ratools_pdf/__init__.py`:

```python
"""Application package for RATools for PDF."""

from ratools_pdf.config.version import APP_VERSION_STR

__all__ = ["APP_VERSION_STR"]
```

Create `ratools_pdf/config/__init__.py`:

```python
"""Configuration helpers for RATools for PDF."""
```

Create `ratools_pdf/pdf/__init__.py`:

```python
"""PDF processing helpers for RATools for PDF."""
```

Create `ratools_pdf/services/__init__.py`:

```python
"""External service integrations for RATools for PDF."""
```

Create `ratools_pdf/ui/__init__.py`:

```python
"""User interface helpers for RATools for PDF."""
```

- [ ] **Step 3: Verify empty package imports**

Run:

```powershell
python -c "import ratools_pdf, ratools_pdf.config, ratools_pdf.pdf, ratools_pdf.services, ratools_pdf.ui; print('ok')"
```

Expected:

```text
ok
```

- [ ] **Step 4: Commit**

Run:

```powershell
git add ratools_pdf
git commit -m "refactor: add application package skeleton"
```

## Phase 1 Task 2: Move Runtime Startup Into `ratools_pdf.app`

**Files:**
- Create: `ratools_pdf/app.py`
- Modify: `main.py`
- Test: `tests/test_package_structure.py`

- [ ] **Step 1: Create a failing package startup import test**

Create `tests/test_package_structure.py`:

```python
import unittest


class PackageStructureTests(unittest.TestCase):
    def test_app_run_is_importable_from_package_and_root_entrypoint(self):
        import main
        from ratools_pdf import app

        self.assertIs(main.run, app.run)


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Run the new test and verify it fails**

Run:

```powershell
python -m unittest tests.test_package_structure -v
```

Expected: FAIL or ERROR because `ratools_pdf.app` does not exist yet.

- [ ] **Step 3: Move the current runtime code into `ratools_pdf/app.py`**

Create `ratools_pdf/app.py` from the current `main.py` logic, with package imports:

```python
import ctypes
import multiprocessing as mp
import sys

from PySide6.QtCore import Qt, QTimer
from PySide6.QtWidgets import QApplication

from ratools_pdf.config.features import ENABLE_UPDATE_CHECK
from controller import MainController
from view import MainWindow


def detach_console_if_needed():
    if sys.platform != "win32" or not getattr(sys, "frozen", False):
        return

    try:
        kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
        if kernel32.GetConsoleWindow():
            kernel32.FreeConsole()
    except Exception:
        pass


def configure_runtime():
    mp.freeze_support()
    detach_console_if_needed()
    QApplication.setHighDpiScaleFactorRoundingPolicy(
        Qt.HighDpiScaleFactorRoundingPolicy.PassThrough
    )


def run():
    configure_runtime()

    app = QApplication(sys.argv)
    view = MainWindow()
    controller = MainController(view)

    view.show()
    if ENABLE_UPDATE_CHECK:
        QTimer.singleShot(0, controller.check_updates_on_startup)
    sys.exit(app.exec())
```

- [ ] **Step 4: Replace `main.py` with a root entry shim**

Replace `main.py` with:

```python
from ratools_pdf.app import run


if __name__ == "__main__":
    run()
```

- [ ] **Step 5: Run the startup import test**

Run:

```powershell
python -m unittest tests.test_package_structure -v
```

Expected: PASS.

- [ ] **Step 6: Commit**

Run:

```powershell
git add main.py ratools_pdf/app.py tests/test_package_structure.py
git commit -m "refactor: move application startup into package"
```

## Phase 1 Task 3: Move Config Modules Behind Compatibility Shims

**Files:**
- Move: `app_features.py` to `ratools_pdf/config/features.py`
- Move: `app_paths.py` to `ratools_pdf/config/paths.py`
- Move: `app_version.py` to `ratools_pdf/config/version.py`
- Recreate: `app_features.py`
- Recreate: `app_paths.py`
- Recreate: `app_version.py`
- Modify: `ratools_pdf/app.py`
- Modify: `update_checker.py`
- Modify: `build_pyinstaller.bat`
- Modify: `build_nuitka.bat`
- Test: `tests/test_package_structure.py`

- [ ] **Step 1: Extend the package import test before moving files**

Add these methods to `tests/test_package_structure.py`:

```python
    def test_config_modules_are_available_from_package(self):
        from ratools_pdf.config import features, paths, version

        self.assertIsInstance(features.ENABLE_UPDATE_CHECK, bool)
        self.assertTrue(callable(paths.get_app_dir))
        self.assertTrue(callable(paths.get_resource_path))
        self.assertTrue(version.APP_VERSION_STR)

    def test_root_config_modules_remain_compatible(self):
        import app_features
        import app_paths
        import app_version
        from ratools_pdf.config import features, paths, version

        self.assertEqual(app_features.ENABLE_UPDATE_CHECK, features.ENABLE_UPDATE_CHECK)
        self.assertIs(app_paths.get_app_dir, paths.get_app_dir)
        self.assertIs(app_paths.get_resource_path, paths.get_resource_path)
        self.assertEqual(app_version.APP_VERSION_STR, version.APP_VERSION_STR)
```

- [ ] **Step 2: Run the package test and verify it fails**

Run:

```powershell
python -m unittest tests.test_package_structure -v
```

Expected: FAIL or ERROR because config modules have not moved yet.

- [ ] **Step 3: Move config files**

Run:

```powershell
git mv app_features.py ratools_pdf\config\features.py
git mv app_paths.py ratools_pdf\config\paths.py
git mv app_version.py ratools_pdf\config\version.py
```

- [ ] **Step 4: Recreate root compatibility shims**

Create `app_features.py`:

```python
from ratools_pdf.config.features import *  # noqa: F401,F403
```

Create `app_paths.py`:

```python
from ratools_pdf.config.paths import *  # noqa: F401,F403
```

Create `app_version.py`:

```python
from ratools_pdf.config.version import *  # noqa: F401,F403
```

- [ ] **Step 5: Update low-risk imports**

In `ratools_pdf/app.py`, keep this import:

```python
from ratools_pdf.config.features import ENABLE_UPDATE_CHECK
```

Do not edit `view.py` in this task. The post-redesign `view.py` can continue importing `app_features`, `app_version`, and `app_paths` through the root compatibility shims until the Phase 2 UI split.

In `update_checker.py`, replace:

```python
from app_version import APP_REPOSITORY_NAME, APP_REPOSITORY_OWNER, APP_VERSION
```

with:

```python
from ratools_pdf.config.version import APP_REPOSITORY_NAME, APP_REPOSITORY_OWNER, APP_VERSION
```

- [ ] **Step 6: Update build scripts to read package version metadata**

In `build_pyinstaller.bat`, replace each `python -c "from app_version import ..."` command with `python -c "from ratools_pdf.config.version import ..."`.

In `build_nuitka.bat`, replace each `python -c "from app_version import ..."` command with `python -c "from ratools_pdf.config.version import ..."`.

- [ ] **Step 7: Run focused tests**

Run:

```powershell
python -m unittest tests.test_package_structure tests.test_app_paths tests.test_app_version -v
```

Expected: PASS.

- [ ] **Step 8: Run all tests**

Run:

```powershell
python -m unittest discover -s tests -p "test_*.py" -v
```

Expected: PASS.

- [ ] **Step 9: Commit**

Run:

```powershell
git add app_features.py app_paths.py app_version.py build_pyinstaller.bat build_nuitka.bat ratools_pdf tests/test_package_structure.py update_checker.py
git commit -m "refactor: move config modules into package"
```

## Phase 1 Task 4: Move Service and UI Helper Modules

**Files:**
- Move: `update_checker.py` to `ratools_pdf/services/update_checker.py`
- Move: `log_view_model.py` to `ratools_pdf/ui/log_view_model.py`
- Move: `theme.py` to `ratools_pdf/ui/theme.py`
- Recreate: `update_checker.py`
- Recreate: `log_view_model.py`
- Recreate: `theme.py`
- Modify: `controller.py`
- Test: `tests/test_package_structure.py`
- Test: `tests/test_update_checker.py`
- Test: `tests/test_log_view_model.py`
- Test: `tests/test_theme.py`

- [ ] **Step 1: Extend the package import test**

Add these methods to `tests/test_package_structure.py`:

```python
    def test_service_modules_are_available_from_package(self):
        import update_checker
        from ratools_pdf.services import update_checker as package_update_checker

        self.assertIs(update_checker.check_for_updates, package_update_checker.check_for_updates)
        self.assertIs(update_checker.release_from_github_payload, package_update_checker.release_from_github_payload)

    def test_ui_helper_modules_are_available_from_package(self):
        import log_view_model
        import theme
        from ratools_pdf.ui import log_view_model as package_log_view_model
        from ratools_pdf.ui import theme as package_theme

        self.assertIs(log_view_model.build_log_summary_items, package_log_view_model.build_log_summary_items)
        self.assertIs(log_view_model.filter_log_summary_items, package_log_view_model.filter_log_summary_items)
        self.assertIs(theme.ThemeManager, package_theme.ThemeManager)
        self.assertIs(theme.active_palette, package_theme.active_palette)
```

- [ ] **Step 2: Run the package test and verify it fails**

Run:

```powershell
python -m unittest tests.test_package_structure -v
```

Expected: FAIL or ERROR because the package modules do not exist yet.

- [ ] **Step 3: Move files**

Run:

```powershell
git mv update_checker.py ratools_pdf\services\update_checker.py
git mv log_view_model.py ratools_pdf\ui\log_view_model.py
git mv theme.py ratools_pdf\ui\theme.py
```

- [ ] **Step 4: Recreate root compatibility shims**

Create `update_checker.py`:

```python
from ratools_pdf.services.update_checker import *  # noqa: F401,F403
```

Create `log_view_model.py`:

```python
from ratools_pdf.ui.log_view_model import *  # noqa: F401,F403
```

Create `theme.py`:

```python
from ratools_pdf.ui.theme import *  # noqa: F401,F403
```

- [ ] **Step 5: Update package-internal imports**

In `controller.py`, replace:

```python
if ENABLE_UPDATE_CHECK:
    import update_checker
else:
    update_checker = None
```

with:

```python
if ENABLE_UPDATE_CHECK:
    from ratools_pdf.services import update_checker
else:
    update_checker = None
```

Do not edit `view.py` in this task. The post-redesign `view.py` can continue importing `log_view_model` and `theme` through root compatibility shims until Phase 2.

- [ ] **Step 6: Run focused tests**

Run:

```powershell
python -m unittest tests.test_package_structure tests.test_update_checker tests.test_log_view_model tests.test_theme -v
```

Expected: PASS.

- [ ] **Step 7: Run all tests**

Run:

```powershell
python -m unittest discover -s tests -p "test_*.py" -v
```

Expected: PASS.

- [ ] **Step 8: Commit**

Run:

```powershell
git add controller.py log_view_model.py theme.py ratools_pdf tests/test_package_structure.py tests/test_theme.py update_checker.py
git commit -m "refactor: move service and ui helper modules into package"
```

## Phase 1 Task 5: Move Font Embedding Provider Module

**Files:**
- Move: `font_embedding_providers.py` to `ratools_pdf/pdf/font_embedding_providers.py`
- Recreate: `font_embedding_providers.py`
- Modify: `pdf_processor.py`
- Test: `tests/test_package_structure.py`

- [ ] **Step 1: Extend the package import test**

Add this method to `tests/test_package_structure.py`:

```python
    def test_pdf_helper_modules_are_available_from_package(self):
        import font_embedding_providers
        from ratools_pdf.pdf import font_embedding_providers as package_font_embedding_providers

        self.assertIs(
            font_embedding_providers.get_font_embedding_provider,
            package_font_embedding_providers.get_font_embedding_provider,
        )
```

- [ ] **Step 2: Run the package test and verify it fails**

Run:

```powershell
python -m unittest tests.test_package_structure -v
```

Expected: FAIL or ERROR because `ratools_pdf.pdf.font_embedding_providers` does not exist yet.

- [ ] **Step 3: Move the file**

Run:

```powershell
git mv font_embedding_providers.py ratools_pdf\pdf\font_embedding_providers.py
```

- [ ] **Step 4: Recreate the root compatibility shim**

Create `font_embedding_providers.py`:

```python
from ratools_pdf.pdf.font_embedding_providers import *  # noqa: F401,F403
```

- [ ] **Step 5: Update package-internal imports**

In `pdf_processor.py`, replace:

```python
from app_paths import get_resource_path
from font_embedding_providers import get_font_embedding_provider
```

with:

```python
from ratools_pdf.config.paths import get_resource_path
from ratools_pdf.pdf.font_embedding_providers import get_font_embedding_provider
```

- [ ] **Step 6: Run focused tests**

Run:

```powershell
python -m unittest tests.test_package_structure tests.test_pdf_processor_roundtrip -v
```

Expected: PASS.

- [ ] **Step 7: Run all tests**

Run:

```powershell
python -m unittest discover -s tests -p "test_*.py" -v
```

Expected: PASS.

- [ ] **Step 8: Commit**

Run:

```powershell
git add font_embedding_providers.py pdf_processor.py ratools_pdf tests/test_package_structure.py
git commit -m "refactor: move font embedding providers into package"
```

## Phase 1 Task 6: Update Documentation and Build Smoke Checks

**Files:**
- Modify: `README.md`

- [ ] **Step 1: Update the project structure section**

Update the `README.md` structure block so it shows:

```text
RATools-for-PDF/
  main.py
  main_no_update.py
  ratools_pdf/
    app.py
    config/
    pdf/
    services/
    ui/
  controller.py
  pdf_processor.py
  view.py
  tests/
  docs/
  plugins/
```

Add one short note below the structure block:

```markdown
Root-level modules such as `controller.py`, `view.py`, and `pdf_processor.py` remain importable for compatibility while implementation code is gradually moved into `ratools_pdf/`.
```

- [ ] **Step 2: Run import smoke checks**

Run:

```powershell
python -c "import main, main_no_update, controller, view, pdf_processor; print('root imports ok')"
python -c "from ratools_pdf.config.version import APP_VERSION_STR; from ratools_pdf.services.update_checker import check_for_updates; print(APP_VERSION_STR, callable(check_for_updates))"
```

Expected:

```text
root imports ok
<version> True
```

- [ ] **Step 3: Verify build metadata commands still work**

Run:

```powershell
python -c "from ratools_pdf.config.version import APP_COMPANY; print(APP_COMPANY)"
python -c "from ratools_pdf.config.version import APP_NAME; print(APP_NAME)"
python -c "from ratools_pdf.config.version import APP_VERSION_STR; print(APP_VERSION_STR)"
python -c "from ratools_pdf.config.version import APP_WINDOWS_VERSION; print(*APP_WINDOWS_VERSION, sep=chr(44)+chr(32))"
```

Expected: each command prints a non-empty value.

- [ ] **Step 4: Run all tests**

Run:

```powershell
python -m unittest discover -s tests -p "test_*.py" -v
```

Expected: PASS.

- [ ] **Step 5: Commit**

Run:

```powershell
git add README.md
git commit -m "docs: document package-based project structure"
```

---

# Phase 2: Split Oversized Modules by Responsibility

## Phase 2 File Map

- Create: `ratools_pdf/controllers/__init__.py`
  Responsibility: controller package exports.
- Create: `ratools_pdf/controllers/io_actions.py`
  Responsibility: IO data action metadata, path building, eCTD rename planning, preview rows.
- Create: `ratools_pdf/controllers/log_export.py`
  Responsibility: text log parsing and structured log export row selection.
- Create: `ratools_pdf/controllers/workers.py`
  Responsibility: `ProcessWorker`, `PreCheckWorker`, `IOActionWorker`, `UpdateCheckWorker`, and process task helpers.
- Create: `ratools_pdf/controllers/main_controller.py`
  Responsibility: `MainController` only.
- Modify: `controller.py`
  Responsibility: compatibility re-export for controller package.
- Create: `ratools_pdf/ui/platform.py`
  Responsibility: Windows version and dialog shadow helpers.
- Create: `ratools_pdf/ui/widgets.py`
  Responsibility: small reusable widgets such as `DropZoneLabel`.
- Create: `ratools_pdf/ui/dialogs.py`
  Responsibility: frameless dialog base class and application dialogs.
- Create: `ratools_pdf/ui/main_window.py`
  Responsibility: `MainWindow` only.
- Modify: `view.py`
  Responsibility: compatibility re-export for UI package.
- Create: `ratools_pdf/pdf/processor.py`
  Responsibility: public `PDFProcessor` facade.
- Create: `ratools_pdf/pdf/qpdf.py`
  Responsibility: qpdf path lookup, rewrite, error formatting, PDF version/encryption checks.
- Create: `ratools_pdf/pdf/precheck.py`
  Responsibility: precheck option filtering, font/link/security report building.
- Create: `ratools_pdf/pdf/bookmarks_links.py`
  Responsibility: bookmark and link import-export helpers.
- Create: `ratools_pdf/pdf/page_layout.py`
  Responsibility: paper rectangles, page resizing, geometry transforms.
- Create: `ratools_pdf/pdf/hyperlink_styles.py`
  Responsibility: hyperlink annotation and visible text styling helpers.
- Modify: `pdf_processor.py`
  Responsibility: compatibility re-export for PDF processor package.
- Modify: `ratools_pdf/app.py`
  Responsibility: import `MainWindow` and `MainController` from package modules.
- Modify: tests under `tests/`
  Responsibility: prefer package imports for new structure while preserving root compatibility coverage.
- Modify: `README.md`
  Responsibility: document final Phase 2 structure.

## Phase 2 Task 1: Extract Controller Pure Helpers

**Files:**
- Create: `ratools_pdf/controllers/__init__.py`
- Create: `ratools_pdf/controllers/io_actions.py`
- Create: `ratools_pdf/controllers/log_export.py`
- Modify: `controller.py`
- Modify: `tests/test_controller_guards.py`
- Test: `tests/test_controller_guards.py`

- [ ] **Step 1: Add package import coverage for controller helpers**

In `tests/test_controller_guards.py`, add imports:

```python
from ratools_pdf.controllers.io_actions import (
    _build_io_paths_for_file as package_build_io_paths_for_file,
    _collect_ectd_rename_plan as package_collect_ectd_rename_plan,
    _normalized_ectd_name as package_normalized_ectd_name,
)
from ratools_pdf.controllers.log_export import (
    _render_logs_as_csv_rows as package_render_logs_as_csv_rows,
    _select_log_rows_for_export as package_select_log_rows_for_export,
)
```

Add this test method:

```python
    def test_controller_helper_package_exports_match_root_exports(self):
        self.assertIs(package_build_io_paths_for_file, _build_io_paths_for_file)
        self.assertIs(package_collect_ectd_rename_plan, _collect_ectd_rename_plan)
        self.assertIs(package_normalized_ectd_name, _normalized_ectd_name)
        self.assertIs(package_render_logs_as_csv_rows, _render_logs_as_csv_rows)
        self.assertIs(package_select_log_rows_for_export, _select_log_rows_for_export)
```

- [ ] **Step 2: Run the focused test and verify it fails**

Run:

```powershell
python -m unittest tests.test_controller_guards.ControllerGuardTests.test_controller_helper_package_exports_match_root_exports -v
```

Expected: ERROR because `ratools_pdf.controllers.io_actions` and `ratools_pdf.controllers.log_export` do not exist.

- [ ] **Step 3: Create `ratools_pdf/controllers/__init__.py`**

Create:

```python
"""Controller helpers and orchestration for RATools for PDF."""
```

- [ ] **Step 4: Move IO helper functions into `io_actions.py`**

Move these existing functions from `controller.py` into `ratools_pdf/controllers/io_actions.py`:

```text
_safe_relative_subdir
_normalized_ectd_name
_collect_ectd_rename_plan
_build_io_paths_for_file
_io_action_metadata
_normalize_io_action_types
_build_io_preview_rows
```

Keep their implementations unchanged. Add the imports they already require:

```python
import os
import re
from pathlib import Path
```

- [ ] **Step 5: Move log export helper functions into `log_export.py`**

Move these existing functions from `controller.py` into `ratools_pdf/controllers/log_export.py`:

```text
_render_logs_as_csv_rows
_log_time_to_seconds
_structured_log_row_from_event
_select_log_rows_for_export
```

Keep their implementations unchanged. Add the imports they already require:

```python
import csv
import re
from datetime import date, datetime
```

- [ ] **Step 6: Import helpers back into `controller.py`**

At the top of `controller.py`, add:

```python
from ratools_pdf.controllers.io_actions import (
    _build_io_paths_for_file,
    _build_io_preview_rows,
    _collect_ectd_rename_plan,
    _io_action_metadata,
    _normalize_io_action_types,
    _normalized_ectd_name,
    _safe_relative_subdir,
)
from ratools_pdf.controllers.log_export import (
    _log_time_to_seconds,
    _render_logs_as_csv_rows,
    _select_log_rows_for_export,
    _structured_log_row_from_event,
)
```

Remove the moved function definitions from `controller.py`.

- [ ] **Step 7: Run focused controller tests**

Run:

```powershell
python -m unittest tests.test_controller_guards -v
```

Expected: PASS.

- [ ] **Step 8: Run all tests**

Run:

```powershell
python -m unittest discover -s tests -p "test_*.py" -v
```

Expected: PASS.

- [ ] **Step 9: Commit**

Run:

```powershell
git add controller.py ratools_pdf/controllers tests/test_controller_guards.py
git commit -m "refactor: extract controller helper modules"
```

## Phase 2 Task 2: Extract Controller Workers

**Files:**
- Create: `ratools_pdf/controllers/workers.py`
- Modify: `controller.py`
- Modify: `tests/test_controller_guards.py`

- [ ] **Step 1: Update worker patch paths in tests**

In `tests/test_controller_guards.py`, replace worker patch paths:

```python
with patch("controller.PDFProcessor.export_bookmarks") as export_bookmarks, \
        patch("controller.PDFProcessor.export_links") as export_links:
```

with:

```python
with patch("ratools_pdf.controllers.workers.PDFProcessor.export_bookmarks") as export_bookmarks, \
        patch("ratools_pdf.controllers.workers.PDFProcessor.export_links") as export_links:
```

Replace:

```python
with patch("controller.PDFProcessor.import_bookmarks") as import_bookmarks, \
        patch("controller.PDFProcessor.import_links") as import_links:
```

with:

```python
with patch("ratools_pdf.controllers.workers.PDFProcessor.import_bookmarks") as import_bookmarks, \
        patch("ratools_pdf.controllers.workers.PDFProcessor.import_links") as import_links:
```

- [ ] **Step 2: Add package import coverage for workers**

In `tests/test_controller_guards.py`, add:

```python
from ratools_pdf.controllers.workers import IOActionWorker as PackageIOActionWorker
```

Add this test method:

```python
    def test_controller_worker_package_export_matches_root_export(self):
        self.assertIs(PackageIOActionWorker, controller.IOActionWorker)
```

- [ ] **Step 3: Run the focused test and verify it fails**

Run:

```powershell
python -m unittest tests.test_controller_guards.ControllerGuardTests.test_controller_worker_package_export_matches_root_export -v
```

Expected: ERROR because `ratools_pdf.controllers.workers` does not exist.

- [ ] **Step 4: Move worker code into `workers.py`**

Move these definitions from `controller.py` into `ratools_pdf/controllers/workers.py`:

```text
_process_document_task
_process_document_task_pipe
ProcessWorker
PreCheckWorker
IOActionWorker
UpdateCheckWorker
```

Add these imports to `workers.py`, adjusted from the old `controller.py` imports:

```python
import multiprocessing as mp
import os
import tempfile
from pathlib import Path

from PySide6.QtCore import QCoreApplication, QThread, Signal

from pdf_processor import PDFProcessor
from ratools_pdf.controllers.io_actions import (
    _build_io_paths_for_file,
    _collect_ectd_rename_plan,
    _normalize_io_action_types,
)
from ratools_pdf.services.update_checker import check_for_updates
```

Keep worker implementations unchanged except for import paths.

- [ ] **Step 5: Import workers back into `controller.py`**

Add:

```python
from ratools_pdf.controllers.workers import (
    IOActionWorker,
    PreCheckWorker,
    ProcessWorker,
    UpdateCheckWorker,
)
```

Remove the moved worker definitions from `controller.py`.

- [ ] **Step 6: Run focused controller tests**

Run:

```powershell
python -m unittest tests.test_controller_guards -v
```

Expected: PASS.

- [ ] **Step 7: Run all tests**

Run:

```powershell
python -m unittest discover -s tests -p "test_*.py" -v
```

Expected: PASS.

- [ ] **Step 8: Commit**

Run:

```powershell
git add controller.py ratools_pdf/controllers/workers.py tests/test_controller_guards.py
git commit -m "refactor: extract controller workers"
```

## Phase 2 Task 3: Move `MainController` Into the Package

**Files:**
- Create: `ratools_pdf/controllers/main_controller.py`
- Modify: `ratools_pdf/controllers/__init__.py`
- Modify: `controller.py`
- Modify: `ratools_pdf/app.py`
- Modify: `tests/test_package_structure.py`

- [ ] **Step 1: Add package import coverage**

Add this method to `tests/test_package_structure.py`:

```python
    def test_main_controller_is_available_from_package_and_root(self):
        import controller
        from ratools_pdf.controllers.main_controller import MainController

        self.assertIs(controller.MainController, MainController)
```

- [ ] **Step 2: Run the focused test and verify it fails**

Run:

```powershell
python -m unittest tests.test_package_structure.PackageStructureTests.test_main_controller_is_available_from_package_and_root -v
```

Expected: ERROR because `ratools_pdf.controllers.main_controller` does not exist.

- [ ] **Step 3: Move `MainController` into `main_controller.py`**

Move only this class from `controller.py` into `ratools_pdf/controllers/main_controller.py`:

```text
MainController
```

Add the imports used by that class. Start from the remaining imports in `controller.py` and keep only the ones referenced by `MainController`:

```python
import os
import platform
import re
import subprocess
import webbrowser
from pathlib import Path

from PySide6.QtCore import QObject, Qt, QTimer
from PySide6.QtGui import QColor
from PySide6.QtWidgets import QFileDialog, QMenu, QTreeWidgetItem

from ratools_pdf.config.features import ENABLE_UPDATE_CHECK
from ratools_pdf.controllers.io_actions import (
    _build_io_preview_rows,
    _io_action_metadata,
    _normalize_io_action_types,
)
from ratools_pdf.controllers.log_export import (
    _select_log_rows_for_export,
    _structured_log_row_from_event,
)
from ratools_pdf.controllers.workers import (
    IOActionWorker,
    PreCheckWorker,
    ProcessWorker,
    UpdateCheckWorker,
)
from view import IODataWizardDialog, LogDialog
```

- [ ] **Step 4: Replace `controller.py` with compatibility exports**

Replace `controller.py` with:

```python
from ratools_pdf.controllers.io_actions import *  # noqa: F401,F403
from ratools_pdf.controllers.log_export import *  # noqa: F401,F403
from ratools_pdf.controllers.main_controller import MainController
from ratools_pdf.controllers.workers import *  # noqa: F401,F403

__all__ = [
    "MainController",
    "ProcessWorker",
    "PreCheckWorker",
    "IOActionWorker",
    "UpdateCheckWorker",
    "_process_document_task",
    "_process_document_task_pipe",
    "_safe_relative_subdir",
    "_normalized_ectd_name",
    "_collect_ectd_rename_plan",
    "_build_io_paths_for_file",
    "_io_action_metadata",
    "_normalize_io_action_types",
    "_build_io_preview_rows",
    "_render_logs_as_csv_rows",
    "_log_time_to_seconds",
    "_structured_log_row_from_event",
    "_select_log_rows_for_export",
]
```

- [ ] **Step 5: Export `MainController` from the controller package**

Update `ratools_pdf/controllers/__init__.py`:

```python
"""Controller helpers and orchestration for RATools for PDF."""

from ratools_pdf.controllers.main_controller import MainController

__all__ = ["MainController"]
```

- [ ] **Step 6: Update the application startup import**

In `ratools_pdf/app.py`, replace:

```python
from controller import MainController
```

with:

```python
from ratools_pdf.controllers.main_controller import MainController
```

- [ ] **Step 7: Run focused tests**

Run:

```powershell
python -m unittest tests.test_package_structure tests.test_controller_guards -v
```

Expected: PASS.

- [ ] **Step 8: Run all tests**

Run:

```powershell
python -m unittest discover -s tests -p "test_*.py" -v
```

Expected: PASS.

- [ ] **Step 9: Commit**

Run:

```powershell
git add controller.py ratools_pdf/app.py ratools_pdf/controllers tests/test_package_structure.py
git commit -m "refactor: move main controller into package"
```

## Phase 2 Task 4: Split UI Module

**Files:**
- Create: `ratools_pdf/ui/platform.py`
- Create: `ratools_pdf/ui/widgets.py`
- Create: `ratools_pdf/ui/dialogs.py`
- Create: `ratools_pdf/ui/main_window.py`
- Modify: `ratools_pdf/ui/theme.py`
- Modify: `view.py`
- Modify: `ratools_pdf/app.py`
- Modify: `ratools_pdf/controllers/main_controller.py`
- Modify: `tests/test_controller_guards.py`
- Modify: `tests/test_package_structure.py`

- [ ] **Step 1: Add package import coverage for UI classes**

Add this method to `tests/test_package_structure.py`:

```python
    def test_ui_classes_are_available_from_package_and_root(self):
        import view
        import theme
        from ratools_pdf.ui.dialogs import IODataWizardDialog, LogDialog
        from ratools_pdf.ui.main_window import MainWindow
        from ratools_pdf.ui.theme import ThemeManager

        self.assertIs(view.IODataWizardDialog, IODataWizardDialog)
        self.assertIs(view.LogDialog, LogDialog)
        self.assertIs(view.MainWindow, MainWindow)
        self.assertIs(theme.ThemeManager, ThemeManager)
```

- [ ] **Step 2: Run the focused test and verify it fails**

Run:

```powershell
python -m unittest tests.test_package_structure.PackageStructureTests.test_ui_classes_are_available_from_package_and_root -v
```

Expected: ERROR because `ratools_pdf.ui.dialogs` and `ratools_pdf.ui.main_window` do not exist.

- [ ] **Step 3: Extract platform helpers**

Move these functions from `view.py` into `ratools_pdf/ui/platform.py`:

```text
is_win11
should_use_manual_dialog_shadow
```

Add imports:

```python
import platform
```

- [ ] **Step 4: Extract reusable widgets**

Move this class from `view.py` into `ratools_pdf/ui/widgets.py`:

```text
DropZoneLabel
```

Add imports:

```python
from PySide6.QtCore import Signal
from PySide6.QtWidgets import QLabel
```

- [ ] **Step 5: Extract dialogs**

Move these classes from `view.py` into `ratools_pdf/ui/dialogs.py`:

```text
FramelessDraggableDialog
CustomMessageBox
ManualFontEmbeddingDialog
IODataWizardDialog
LogDialog
SettingsDialog
AboutDialog
```

Add imports by copying the imports used by those classes from the original `view.py`, then remove unused imports after tests pass. The module must import package helpers directly:

```python
from ratools_pdf.config.paths import get_app_dir, get_resource_path
from ratools_pdf.config.version import get_display_version
from ratools_pdf.ui.log_view_model import (
    build_log_summary_items,
    filter_log_summary_items,
    format_duration,
    log_status_tags,
)
from ratools_pdf.ui.platform import should_use_manual_dialog_shadow
from ratools_pdf.ui.theme import active_palette, log_status_colors
```

- [ ] **Step 6: Extract main window**

Move `MainWindow` from `view.py` into `ratools_pdf/ui/main_window.py`.

Add package imports:

```python
from ratools_pdf.config.features import ENABLE_UPDATE_CHECK
from ratools_pdf.config.paths import get_app_dir, get_resource_path
from ratools_pdf.ui.dialogs import (
    AboutDialog,
    CustomMessageBox,
    IODataWizardDialog,
    LogDialog,
    ManualFontEmbeddingDialog,
    SettingsDialog,
)
from ratools_pdf.ui.theme import active_palette, build_app_qss, log_status_colors, ThemeManager
from ratools_pdf.ui.widgets import DropZoneLabel
```

- [ ] **Step 7: Replace `view.py` with compatibility exports**

Replace `view.py` with:

```python
from ratools_pdf.ui.dialogs import *  # noqa: F401,F403
from ratools_pdf.ui.main_window import MainWindow
from ratools_pdf.ui.platform import *  # noqa: F401,F403
from ratools_pdf.ui.theme import *  # noqa: F401,F403
from ratools_pdf.ui.widgets import DropZoneLabel

__all__ = [
    "is_win11",
    "should_use_manual_dialog_shadow",
    "FramelessDraggableDialog",
    "CustomMessageBox",
    "ManualFontEmbeddingDialog",
    "IODataWizardDialog",
    "LogDialog",
    "SettingsDialog",
    "AboutDialog",
    "DropZoneLabel",
    "MainWindow",
    "Palette",
    "LIGHT",
    "DARK",
    "active_palette",
    "build_app_qss",
    "log_status_colors",
    "ThemeManager",
]
```

- [ ] **Step 8: Update package-internal imports**

In `ratools_pdf/app.py`, replace:

```python
from view import MainWindow
```

with:

```python
from ratools_pdf.ui.main_window import MainWindow
```

In `ratools_pdf/controllers/main_controller.py`, replace:

```python
from view import IODataWizardDialog, LogDialog
```

with:

```python
from ratools_pdf.ui.dialogs import IODataWizardDialog, LogDialog
```

In `tests/test_controller_guards.py`, replace:

```python
from view import IODataWizardDialog
```

with:

```python
from ratools_pdf.ui.dialogs import IODataWizardDialog
```

- [ ] **Step 9: Run focused UI tests**

Run:

```powershell
python -m unittest tests.test_package_structure tests.test_controller_guards -v
```

Expected: PASS.

- [ ] **Step 10: Run all tests**

Run:

```powershell
python -m unittest discover -s tests -p "test_*.py" -v
```

Expected: PASS.

- [ ] **Step 11: Commit**

Run:

```powershell
git add view.py ratools_pdf/app.py ratools_pdf/controllers/main_controller.py ratools_pdf/ui tests/test_controller_guards.py tests/test_package_structure.py
git commit -m "refactor: split ui module"
```

## Phase 2 Task 5: Move `PDFProcessor` Into the Package

**Files:**
- Create: `ratools_pdf/pdf/processor.py`
- Modify: `pdf_processor.py`
- Modify: `ratools_pdf/controllers/workers.py`
- Modify: `tests/test_package_structure.py`
- Modify: `tests/test_pdf_processor_roundtrip.py`

- [ ] **Step 1: Add package import coverage for `PDFProcessor`**

Add this method to `tests/test_package_structure.py`:

```python
    def test_pdf_processor_is_available_from_package_and_root(self):
        import pdf_processor
        from ratools_pdf.pdf.processor import PDFProcessor

        self.assertIs(pdf_processor.PDFProcessor, PDFProcessor)
```

- [ ] **Step 2: Run the focused test and verify it fails**

Run:

```powershell
python -m unittest tests.test_package_structure.PackageStructureTests.test_pdf_processor_is_available_from_package_and_root -v
```

Expected: ERROR because `ratools_pdf.pdf.processor` does not exist.

- [ ] **Step 3: Move current `pdf_processor.py` implementation into the package**

Run:

```powershell
git mv pdf_processor.py ratools_pdf\pdf\processor.py
```

- [ ] **Step 4: Recreate the root compatibility shim**

Create `pdf_processor.py`:

```python
from ratools_pdf.pdf.processor import *  # noqa: F401,F403
```

- [ ] **Step 5: Update package-internal imports**

In `ratools_pdf/controllers/workers.py`, replace:

```python
from pdf_processor import PDFProcessor
```

with:

```python
from ratools_pdf.pdf.processor import PDFProcessor
```

In `tests/test_pdf_processor_roundtrip.py`, replace:

```python
from pdf_processor import PDFProcessor
```

with:

```python
from ratools_pdf.pdf.processor import PDFProcessor
```

- [ ] **Step 6: Run focused PDF tests**

Run:

```powershell
python -m unittest tests.test_package_structure tests.test_pdf_processor_roundtrip -v
```

Expected: PASS.

- [ ] **Step 7: Run all tests**

Run:

```powershell
python -m unittest discover -s tests -p "test_*.py" -v
```

Expected: PASS.

- [ ] **Step 8: Commit**

Run:

```powershell
git add pdf_processor.py ratools_pdf/pdf tests/test_package_structure.py tests/test_pdf_processor_roundtrip.py
git commit -m "refactor: move pdf processor into package"
```

## Phase 2 Task 6: Extract PDF Helper Modules Incrementally

**Files:**
- Create: `ratools_pdf/pdf/qpdf.py`
- Create: `ratools_pdf/pdf/precheck.py`
- Create: `ratools_pdf/pdf/bookmarks_links.py`
- Create: `ratools_pdf/pdf/page_layout.py`
- Create: `ratools_pdf/pdf/hyperlink_styles.py`
- Modify: `ratools_pdf/pdf/processor.py`
- Test: `tests/test_pdf_processor_roundtrip.py`

- [ ] **Step 1: Extract qpdf helpers first**

Move these static helper bodies out of `PDFProcessor` into `ratools_pdf/pdf/qpdf.py`:

```text
_get_qpdf_path
_rewrite_with_qpdf
_format_qpdf_error
_read_pdf_header_version
_is_pdf_linearized
_qpdf_encryption_info
_qpdf_reports_restrictions
```

Expose package functions with non-class names:

```python
def get_qpdf_path():
    ...


def rewrite_with_qpdf(input_pdf, output_pdf, force_version=None, linearize=False, decrypt_restrictions=False):
    ...


def format_qpdf_error(error):
    ...
```

Keep compatibility wrappers on `PDFProcessor`:

```python
    @staticmethod
    def _get_qpdf_path():
        return qpdf.get_qpdf_path()

    @staticmethod
    def _rewrite_with_qpdf(input_pdf, output_pdf, force_version=None, linearize=False, decrypt_restrictions=False):
        return qpdf.rewrite_with_qpdf(input_pdf, output_pdf, force_version, linearize, decrypt_restrictions)

    @staticmethod
    def _format_qpdf_error(error):
        return qpdf.format_qpdf_error(error)
```

- [ ] **Step 2: Run PDF tests after qpdf extraction**

Run:

```powershell
python -m unittest tests.test_pdf_processor_roundtrip -v
```

Expected: PASS.

- [ ] **Step 3: Extract bookmark and link import-export helpers**

Move the implementation bodies for these methods into `ratools_pdf/pdf/bookmarks_links.py`:

```text
export_bookmarks
import_bookmarks
export_links
import_links
_resolve_external_file_target
_read_target_pdf_page_count
_add_link_target_integrity_findings
```

Keep compatibility wrappers on `PDFProcessor`:

```python
    @staticmethod
    def export_bookmarks(pdf_path, csv_path):
        return bookmarks_links.export_bookmarks(pdf_path, csv_path)

    @staticmethod
    def import_bookmarks(pdf_path, csv_path, output_path):
        return bookmarks_links.import_bookmarks(pdf_path, csv_path, output_path)

    @staticmethod
    def export_links(pdf_path, json_path):
        return bookmarks_links.export_links(pdf_path, json_path)

    @staticmethod
    def import_links(pdf_path, json_path, output_path):
        return bookmarks_links.import_links(pdf_path, json_path, output_path)
```

- [ ] **Step 4: Run PDF round-trip tests after bookmark/link extraction**

Run:

```powershell
python -m unittest tests.test_pdf_processor_roundtrip -v
```

Expected: PASS.

- [ ] **Step 5: Extract page layout helpers**

Move these helpers into `ratools_pdf/pdf/page_layout.py`:

```text
_transform_rect
_transform_point
_get_oriented_target_rect
_paper_rect_exact
_resize_pages_with_padding
```

Keep compatibility wrappers on `PDFProcessor` for each static method.

- [ ] **Step 6: Run all PDF tests after page layout extraction**

Run:

```powershell
python -m unittest tests.test_pdf_processor_roundtrip -v
```

Expected: PASS.

- [ ] **Step 7: Extract hyperlink style helpers**

Move these helpers into `ratools_pdf/pdf/hyperlink_styles.py`:

```text
_is_text_blue
_overlay_text_color_in_rect
_link_has_visible_border
_force_link_new_window
_rects_intersect
_point_in_any_rect
_make_text_block_blue
_make_text_block_color
_apply_text_color_via_content_stream
_collect_page_state
_apply_blue_text_via_content_stream
_apply_hyperlink_actions
_apply_hyperlink_styles
```

Keep compatibility wrappers on `PDFProcessor` for each static method.

- [ ] **Step 8: Run all PDF tests after hyperlink extraction**

Run:

```powershell
python -m unittest tests.test_pdf_processor_roundtrip -v
```

Expected: PASS.

- [ ] **Step 9: Extract precheck helpers last**

Move these helpers into `ratools_pdf/pdf/precheck.py`:

```text
_option_title
_precheck_option_matches_selected
_selected_precheck_option_id
_filtered_precheck_options
_normalize_font_name
_is_base14_font
_format_font_page_numbers
_font_object_has_embedded_file
_font_tuple_value
_font_tuple_embedded_fallback
_collect_font_precheck_findings
_font_precheck_has_embedding_risk
_collect_font_precheck_for_path
_add_precheck_suggestion
_add_precheck_report_finding
_catalog_key
_catalog_key_is_present
_dereference_xref_value
_catalog_key_resolved_value
build_precheck_report
resolve_processing_options
```

Keep public compatibility wrappers on `PDFProcessor`:

```python
    @staticmethod
    def build_precheck_report(input_path, selected_options=None):
        return precheck.build_precheck_report(input_path, selected_options)

    @staticmethod
    def resolve_processing_options(input_path, options, processing_mode="smart"):
        return precheck.resolve_processing_options(input_path, options, processing_mode)
```

Keep private static wrappers only where existing tests or same-module code still call them. Remove unused private wrappers after all tests pass.

- [ ] **Step 10: Run all tests after precheck extraction**

Run:

```powershell
python -m unittest discover -s tests -p "test_*.py" -v
```

Expected: PASS.

- [ ] **Step 11: Commit**

Run:

```powershell
git add pdf_processor.py ratools_pdf/pdf tests
git commit -m "refactor: extract pdf helper modules"
```

## Phase 2 Task 7: Update Documentation and Final Verification

**Files:**
- Modify: `README.md`

- [ ] **Step 1: Update the README structure block to the Phase 2 target**

Use this structure:

```text
RATools-for-PDF/
  main.py
  main_no_update.py
  ratools_pdf/
    app.py
    config/
    controllers/
    pdf/
    services/
    ui/
  controller.py        # compatibility shim
  pdf_processor.py     # compatibility shim
  view.py              # compatibility shim
  tests/
  docs/
  plugins/
```

- [ ] **Step 2: Add a short maintainer note**

Add:

```markdown
New code should import from `ratools_pdf.*` modules. Root-level modules remain for backward compatibility with existing tests, scripts, and user workflows.
```

- [ ] **Step 3: Run root import smoke checks**

Run:

```powershell
python -c "import main, main_no_update, controller, view, pdf_processor; print('root imports ok')"
```

Expected:

```text
root imports ok
```

- [ ] **Step 4: Run package import smoke checks**

Run:

```powershell
python -c "from ratools_pdf.controllers.main_controller import MainController; from ratools_pdf.ui.main_window import MainWindow; from ratools_pdf.pdf.processor import PDFProcessor; print(MainController.__name__, MainWindow.__name__, PDFProcessor.__name__)"
```

Expected:

```text
MainController MainWindow PDFProcessor
```

- [ ] **Step 5: Run all tests**

Run:

```powershell
python -m unittest discover -s tests -p "test_*.py" -v
```

Expected: PASS.

- [ ] **Step 6: Check changed files**

Run:

```powershell
git status --short
```

Expected: only planned files changed.

- [ ] **Step 7: Commit**

Run:

```powershell
git add README.md
git commit -m "docs: document refactored package structure"
```

---

## Risk Notes

- `feature/ui-redesign` must land before Phase 1. Its active worktree currently changes `view.py` and adds `theme.py` plus `tests/test_theme.py`; splitting the older `view.py` would create avoidable conflicts.
- `main.py` must stay in the root because both build scripts currently point at it.
- Root compatibility shims reduce risk for tests and build scripts, but they should not become the preferred import style.
- Phase 1 intentionally avoids unnecessary `view.py` edits. Root shims keep post-redesign imports such as `theme`, `log_view_model`, `app_paths`, and `app_version` working until Phase 2.
- Worker tests that patch `controller.PDFProcessor` must be updated when workers move, because patching the old root alias will no longer intercept package-local imports.
- `view.py` extraction is mostly copy/move work, but it is PySide-heavy, theme-aware, and import-order-sensitive. Run focused UI tests after each extraction.
- `pdf_processor.py` extraction should happen after controller and UI moves. It has the broadest behavioral blast radius.
- The first PDF extraction should be a whole-file move to `ratools_pdf/pdf/processor.py`; helper extraction should happen only after that move is green.

## Verification Checklist

- [ ] `git -C .worktrees\ui-redesign status --short` shows a clean worktree before Phase 1 starts.
- [ ] `python -m unittest discover -s tests -p "test_*.py" -v` passes after every task.
- [ ] `python -m unittest tests.test_theme -v` passes after the UI redesign merge and after moving `theme.py`.
- [ ] `python -c "import main, main_no_update, controller, view, pdf_processor; print('root imports ok')"` passes after every phase.
- [ ] `python -c "import theme; from ratools_pdf.ui.theme import ThemeManager; print(theme.ThemeManager is ThemeManager)"` prints `True` after Phase 1 Task 4.
- [ ] `python -c "from ratools_pdf.pdf.processor import PDFProcessor; print(PDFProcessor.__name__)"` passes after Phase 2 Task 5.
- [ ] `build_pyinstaller.bat` still finds `main.py`, `main_no_update.py`, `icon.ico`, `plugins`, and version metadata.
- [ ] `build_nuitka.bat` still finds `main.py`, `icon.ico`, `plugins`, and version metadata.

## Deferred Phase 3 Ideas

- Remove root compatibility shims after one release cycle, if no scripts or docs depend on them.
- Add a `pyproject.toml` with test and packaging metadata.
- Consider a `src/` layout only after package imports are stable.
- Split `ratools_pdf/ui/dialogs.py` further into one module per dialog if it remains too large.
- Replace private static wrapper methods in `PDFProcessor` with direct module functions where no external compatibility is needed.

## Self-Review

- Spec coverage: covers the requested first and second phases, root cleanup, package structure, large-module splitting, tests, docs, and build-script safety.
- Placeholder scan: no placeholder markers or unspecified follow-up steps remain in the executable phase tasks.
- Type consistency: all referenced module names, class names, and test names match the current repository names or files created by earlier tasks.
- Scope check: Phase 1 is intentionally low-risk packaging work; Phase 2 is the larger structural split. Behavior changes are explicitly out of scope.

## Execution Handoff

Plan complete. Recommended implementation mode: complete Phase 1 first, verify and optionally release it, then start Phase 2 in a separate branch or worktree.
