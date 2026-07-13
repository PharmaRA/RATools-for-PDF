# Build Path Merge Readiness Design

## Goal

Make `codex/project-structure-refactor` safe to merge into the local `main` branch by removing the two remaining references to deleted root-level compatibility modules and preventing those references from returning.

## Scope

The change is intentionally limited to release and build path correctness:

- Update `.github/workflows/build.yml` to import version metadata from `ratools_pdf.config.version`.
- Update `build_nuitka.bat` to include `ratools_pdf.config.paths` instead of the removed `app_paths` module.
- Extend `tests/test_package_structure.py` so release and build configuration cannot silently return to removed root-module paths.
- Do not restore compatibility shim modules.
- Do not refactor unrelated build, application, PDF, controller, or UI code.
- Do not push any branch to the remote repository.

## Implementation Approach

Use the existing package layout as the sole source of application modules. The release workflow and Nuitka script will reference package-qualified module names, matching the already-updated PyInstaller script and application imports.

Regression coverage will inspect the build files as text. It will assert that the package-qualified paths are present and that the removed root-module references are absent. This keeps the tests fast and allows the normal CI job to catch the exact configuration regression that the current test suite missed.

## Verification

Verification will run first on `codex/project-structure-refactor` and again after the local fast-forward merge into `main`:

1. Run the focused package-structure tests.
2. Run the complete `unittest` suite.
3. Compile the application, package, and tests with `compileall`.
4. Import the root entrypoints and core package modules.
5. Execute the same version-import command used by the release workflow.
6. Confirm the legacy module names no longer appear in active build or release configuration.
7. Confirm both worktrees are clean and `main` contains the feature branch commit.

Full PyInstaller and Nuitka binary builds are outside this minimal merge-readiness change. The checks will validate their corrected module paths, while a future release should still run the normal packaging workflow and executable smoke tests.

## Git Procedure

1. Commit the regression test and build-path fixes on `codex/project-structure-refactor`.
2. Verify the feature branch.
3. Fast-forward the local `main` branch to the feature branch.
4. Verify the merged local `main` branch.
5. Leave `origin/main` and all other remote refs unchanged.

## Success Criteria

- The release workflow no longer imports `app_version`.
- The Nuitka script no longer includes `app_paths`.
- Regression tests cover both corrected paths.
- All local verification commands exit successfully.
- Local `main` contains the completed refactor and fix commit.
- No remote push occurs.
