# Open Source Baseline Phase 1 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a trustworthy minimum open-source collaboration and validation baseline to `RATools-for-PDF`.

**Architecture:** Keep the implementation lightweight and repository-focused. Add short governance docs, add a minimal deterministic `unittest` suite, wire that suite into a Windows GitHub Actions workflow, and align `README.md` plus `.gitignore` with the repository's actual structure.

**Tech Stack:** Markdown, GitHub Actions YAML, Python 3 `unittest`, `unittest.mock`.

---

## File Map

- Create: `CONTRIBUTING.md`
  Responsibility: contributor entry point and maintainer expectations.
- Create: `SECURITY.md`
  Responsibility: security disclosure process and expectations.
- Create: `CODE_OF_CONDUCT.md`
  Responsibility: collaboration norms and enforcement posture.
- Create: `.github/workflows/ci.yml`
  Responsibility: install dependencies and run the canonical regression command.
- Create: `tests/test_app_paths.py`
  Responsibility: resource/app path behavior coverage.
- Create: `tests/test_app_version.py`
  Responsibility: version display coverage.
- Create: `tests/test_update_checker.py`
  Responsibility: update-checker parsing and version decision coverage.
- Modify: `README.md`
  Responsibility: accurate testing/collaboration/structure documentation.
- Modify: `.gitignore`
  Responsibility: stop conflicting with tracked docs/tests content.

## Task Order

### Task 1: Add the first regression tests

- Create `tests/test_app_paths.py` for app/resource path resolution.
- Create `tests/test_app_version.py` for version text formatting.
- Create `tests/test_update_checker.py` for release parsing and update decision rules.
- Run `python -m unittest discover -s tests -p "test_*.py" -v`.

### Task 2: Add governance documents

- Create `CONTRIBUTING.md` with short bilingual maintainer guidance.
- Create `SECURITY.md` with private-report guidance.
- Create `CODE_OF_CONDUCT.md` with a short custom collaboration standard.

### Task 3: Add the minimum CI workflow

- Create `.github/workflows/ci.yml`.
- Use `windows-latest`.
- Install `requirements.txt` and run the same `unittest` command as local docs.

### Task 4: Align README and ignore rules with reality

- Update the tech stack entry from `pytest / unittest` to the actual minimal `unittest` suite.
- Update the project structure block to include `tests/`, `docs/`, and `.github/workflows/` accurately.
- Rewrite the testing section to match the new suite and workflow.
- Add a short `Contributing` section that points to repository governance files.
- Remove `.gitignore` rules that conflict with tracked `docs/` and `tests/` content.

### Task 5: Verify the new baseline

- Run `python -m unittest discover -s tests -p "test_*.py" -v`.
- Run `git status --short` to confirm only planned files changed.
- Summarize deferred Phase 2 and Phase 3 work.

## Self-Review

- Spec coverage: covers governance docs, tests, CI, README alignment, and ignore-rule cleanup from the approved Phase 1 scope.
- Placeholder scan: no `TODO` or `TBD` placeholders remain.
- Type consistency: all referenced files and commands exist or are created by this plan.

## Execution Handoff

The user already requested continuous execution, so proceed inline with the plan.
