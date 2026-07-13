# Build Path Merge Readiness Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Correct the remaining legacy module paths, add regression coverage, and fast-forward the verified refactor into the local `main` branch without pushing.

**Architecture:** Keep all application imports package-qualified under `ratools_pdf`. Extend the existing package-structure test to inspect active release and build configuration, then make the two minimal configuration edits required for that test to pass. Verify on the feature branch and again after a local fast-forward merge.

**Tech Stack:** Python 3, `unittest`, GitHub Actions YAML, Windows batch, Git.

---

## File Map

- Modify `.github/workflows/build.yml`: use the packaged version module in the release workflow.
- Modify `build_nuitka.bat`: include the packaged paths module in Nuitka builds.
- Modify `tests/test_package_structure.py`: protect both build paths from legacy root-module regressions.
- Create `docs/superpowers/plans/2026-07-13-build-path-merge-readiness.md`: record this implementation sequence.

### Task 1: Add failing release and Nuitka path regression checks

**Files:**
- Modify: `tests/test_package_structure.py:86-100`
- Test: `tests/test_package_structure.py`

- [ ] **Step 1: Extend the existing build-script test**

Replace `test_build_scripts_use_package_module_paths` with:

```python
def test_build_scripts_use_package_module_paths(self):
    repo_root = Path(__file__).resolve().parents[1]
    pyinstaller_script = (repo_root / "build_pyinstaller.bat").read_text(encoding="utf-8")
    nuitka_script = (repo_root / "build_nuitka.bat").read_text(encoding="utf-8")
    release_workflow = (repo_root / ".github" / "workflows" / "build.yml").read_text(
        encoding="utf-8"
    )

    package_version_import = "from ratools_pdf.config.version import APP_VERSION_STR"

    self.assertIn(package_version_import, pyinstaller_script)
    self.assertIn(package_version_import, nuitka_script)
    self.assertIn(package_version_import, release_workflow)
    self.assertIn("--include-module=ratools_pdf.config.paths", nuitka_script)
    self.assertIn("--exclude-module ratools_pdf.services.update_checker", pyinstaller_script)
    self.assertNotIn("from app_version import", pyinstaller_script)
    self.assertNotIn("from app_version import", nuitka_script)
    self.assertNotIn("from app_version import", release_workflow)
    self.assertNotIn("--include-module=app_paths", nuitka_script)
    self.assertNotIn("--exclude-module update_checker", pyinstaller_script)
```

- [ ] **Step 2: Run the focused test and confirm the regression is exposed**

Run from the feature worktree:

```powershell
python -m unittest tests.test_package_structure.PackageStructureTests.test_build_scripts_use_package_module_paths -v
```

Expected: FAIL because `.github/workflows/build.yml` still imports `app_version` and `build_nuitka.bat` still contains `--include-module=app_paths`.

### Task 2: Correct the release workflow and Nuitka module paths

**Files:**
- Modify: `.github/workflows/build.yml:42`
- Modify: `build_nuitka.bat:57`
- Test: `tests/test_package_structure.py`

- [ ] **Step 1: Update the release workflow version import**

Change the `Resolve build version` command to:

```yaml
$version = python -c "from ratools_pdf.config.version import APP_VERSION_STR; print(APP_VERSION_STR)"
```

- [ ] **Step 2: Update the Nuitka explicit module include**

Change the Nuitka option to:

```bat
--include-module=ratools_pdf.config.paths ^
```

- [ ] **Step 3: Run the focused test and confirm it passes**

```powershell
python -m unittest tests.test_package_structure.PackageStructureTests.test_build_scripts_use_package_module_paths -v
```

Expected: one test passes with `OK`.

- [ ] **Step 4: Run the release-workflow import directly**

```powershell
python -c "from ratools_pdf.config.version import APP_VERSION_STR; print(APP_VERSION_STR)"
```

Expected: exits 0 and prints the current application version.

- [ ] **Step 5: Confirm active build configuration has no legacy paths**

```powershell
git grep -n -E 'from app_version import|--include-module=app_paths' HEAD -- '.github/workflows/build.yml' 'build_nuitka.bat' 'build_pyinstaller.bat'
```

Expected: exit 1 with no matches, because `git grep` returns 1 when the requested legacy text is absent.

- [ ] **Step 6: Commit the regression test and fixes**

```powershell
git add -- tests/test_package_structure.py .github/workflows/build.yml build_nuitka.bat
git diff --cached --check
git commit -m "fix: update packaged build module paths"
```

Expected: the staged diff check exits 0 and the commit is created on `codex/project-structure-refactor`.

### Task 3: Verify the completed feature branch

**Files:**
- Verify: `main.py`
- Verify: `main_no_update.py`
- Verify: `ratools_pdf/`
- Verify: `tests/`

- [ ] **Step 1: Run the complete test suite**

```powershell
python -m unittest discover -s tests -p "test_*.py" -v
```

Expected: all tests pass with `OK`.

- [ ] **Step 2: Compile application and test modules**

```powershell
python -m compileall -q main.py main_no_update.py ratools_pdf tests
```

Expected: exit 0 with no syntax errors.

- [ ] **Step 3: Import the entrypoints and core package modules**

```powershell
python -c "import main; import main_no_update; import ratools_pdf.app; import ratools_pdf.controllers.main_controller; import ratools_pdf.pdf.processor; print('ENTRY_IMPORTS_OK')"
```

Expected: exit 0 and print `ENTRY_IMPORTS_OK`.

- [ ] **Step 4: Confirm branch cleanliness and commit relationship**

```powershell
git status --short --branch
git merge-base --is-ancestor main codex/project-structure-refactor
```

Expected: the feature worktree is clean and the ancestry check exits 0.

### Task 4: Fast-forward and verify the local main branch

**Files:**
- Update local branch ref: `main`
- Do not update: `origin/main`

- [ ] **Step 1: Confirm the main worktree is clean**

```powershell
git -C E:\02_GitHub\RATools-for-PDF status --short --branch
```

Expected: no modified or untracked files.

- [ ] **Step 2: Fast-forward local main**

```powershell
git -C E:\02_GitHub\RATools-for-PDF merge --ff-only codex/project-structure-refactor
```

Expected: Git reports a fast-forward and local `main` points to the feature branch tip.

- [ ] **Step 3: Re-run the full tests from main**

```powershell
python -m unittest discover -s tests -p "test_*.py" -v
```

Expected: all tests pass with `OK`.

- [ ] **Step 4: Re-run compile and entrypoint checks from main**

```powershell
python -m compileall -q main.py main_no_update.py ratools_pdf tests
python -c "import main; import main_no_update; import ratools_pdf.app; import ratools_pdf.controllers.main_controller; import ratools_pdf.pdf.processor; print('ENTRY_IMPORTS_OK')"
```

Expected: both commands exit 0 and the import check prints `ENTRY_IMPORTS_OK`.

- [ ] **Step 5: Confirm local-only merge state**

```powershell
git status --short --branch
git rev-parse main
git rev-parse codex/project-structure-refactor
git rev-parse origin/main
```

Expected: `main` and `codex/project-structure-refactor` have the same commit; `origin/main` remains at its earlier commit; the main worktree is clean and reports it is ahead of `origin/main`.
