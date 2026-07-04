# PyInstaller Version Info Generation Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Remove the static `build_version_info.txt` file and make `build_pyinstaller.bat` generate a temporary version info file from `app_version.py` at build time.

**Architecture:** Keep `app_version.py` as the single source of truth for version metadata. Update `build_pyinstaller.bat` to ask Python for the current app metadata, write a temporary version-info file into the build directory, pass that file to PyInstaller, and delete it after each variant build. Update tests to verify the batch script now generates version info dynamically and that the repository no longer depends on a checked-in `build_version_info.txt`.

**Tech Stack:** Windows batch, Python standard library, existing `unittest` tests, `app_version.py`, PyInstaller.

---

### Task 1: Replace Static Version File Usage

**Files:**
- Modify: `build_pyinstaller.bat`
- Modify: `test_build_pyinstaller.py`
- Delete: `build_version_info.txt`

- [ ] **Step 1: Write the failing test**

Update `test_build_pyinstaller.py` so it checks that:

```python
self.assertIn('from app_version import APP_COMPANY, APP_NAME, APP_VERSION, APP_VERSION_STR', script)
self.assertIn('set "VERSION_INFO_FILE=%BUILD_DIR%\\build_version_info_generated.txt"', script)
self.assertIn('--version-file "%VERSION_INFO_FILE%"', script)
self.assertFalse(Path("build_version_info.txt").exists())
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m unittest test_build_pyinstaller.py -v`
Expected: FAIL because the script still references `build_version_info.txt` and the file still exists.

- [ ] **Step 3: Write minimal implementation**

Update `build_pyinstaller.bat` to:

```bat
set "VERSION_INFO_FILE=%BUILD_DIR%\build_version_info_generated.txt"
python -c "from pathlib import Path; from app_version import APP_COMPANY, APP_NAME, APP_VERSION, APP_VERSION_STR; Path(r'%VERSION_INFO_FILE%').write_text(...)"
```

and use:

```bat
  --version-file "%VERSION_INFO_FILE%" ^
```

Then delete `build_version_info.txt` from the repository.

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m unittest test_build_pyinstaller.py -v`
Expected: PASS.

- [ ] **Step 5: Run broader verification**

Run: `python -m unittest discover -v`
Expected: full suite passes.

## Self-Review

- Spec coverage: plan covers removing the checked-in version file, generating a temporary PyInstaller version file from `app_version.py`, and updating tests.
- Placeholder scan: no incomplete placeholders remain.
- Type consistency: `APP_COMPANY`, `APP_NAME`, `APP_VERSION`, and `APP_VERSION_STR` are the only required metadata inputs across script and tests.
