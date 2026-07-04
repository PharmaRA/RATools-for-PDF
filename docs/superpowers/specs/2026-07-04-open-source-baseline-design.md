# Open Source Baseline Design

## Context

`RATools-for-PDF` already has a strong product-facing `README.md`, an AGPL license, and third-party notices. The current gap is not product description but open-source project infrastructure: external contributors do not yet have a clear collaboration path, the repository does not expose a minimal CI signal, and some repository-level files describe structure or automation that is not actually present.

The desired project shape is a personal-maintainer-led open source repository. That means the project should welcome issues and focused pull requests, while keeping long-term maintenance cost low for a single primary maintainer.

The preferred language strategy is bilingual with minimal overhead: keep the main explanations Chinese-first where that best matches the current audience, but make the collaboration surface understandable to English-speaking contributors with concise English equivalents.

## Goals

- Establish a trustworthy minimum open-source baseline for a personal-maintained project.
- Make the repository self-explanatory for users, issue reporters, and focused contributors.
- Add a minimal automated validation signal that proves the repository can run at least a small regression suite.
- Align `README.md`, repository structure, and ignore rules with the files that actually exist.
- Keep every new governance file short, explicit, and cheap to maintain.

## Non-Goals

- Do not turn the repository into a high-process community-governed project.
- Do not introduce heavy contribution bureaucracy, multi-role governance, or long policy documents.
- Do not require complex cross-platform CI in the first phase when the project is primarily Windows desktop oriented.
- Do not add screenshot galleries, roadmap pages, issue templates, or PR templates in the first phase unless they are required to support the minimum trust baseline.

## Maintenance Principles

### 1. Personal-maintainer-led

The maintainer remains the final decision-maker on scope, release timing, and contribution acceptance. Repository documents should explicitly welcome contributions without creating an expectation of rapid support or guaranteed review.

### 2. Minimum trust before maximum polish

The first phase should prioritize artifacts that increase credibility:

- clear contribution path
- clear security reporting path
- clear expected conduct
- visible automated checks
- accurate documentation

Presentation upgrades such as screenshots and extended templates come later.

### 3. Bilingual, but asymmetrically

To keep maintenance cost low:

- core user-facing product docs stay Chinese-first
- collaboration docs use short bilingual sections
- headings may remain English when that matches GitHub conventions
- duplicated prose should be avoided unless it materially improves contributor comprehension

### 4. Match claims to reality

The repository should not claim tests, CI, or project structure that are absent. If a file mentions automation, tests, or directories, the repo should either contain them or describe them as planned rather than present.

## Phased Rollout

## Phase 1: Minimum Trustworthy Baseline

This phase is the only one implemented immediately.

Deliverables:

- `CONTRIBUTING.md`
- `SECURITY.md`
- `CODE_OF_CONDUCT.md`
- a minimal GitHub Actions workflow that installs dependencies and runs a small regression suite
- a small `tests/` suite focused on stable, non-GUI utility behavior
- `README.md` updates so the testing and collaboration sections match the repository
- `.gitignore` cleanup so `docs/` and `tests/` are no longer awkwardly ignored

### Phase 1 Language Strategy

- Each governance file should open with a short English heading and then use concise bilingual content.
- The wording should be direct and short rather than fully mirrored paragraph-by-paragraph.
- Commands and repository conventions stay language-neutral in code blocks.

### Phase 1 Testing Strategy

The first CI signal should prove lightweight, deterministic behavior without depending on GUI interaction or external network access.

The regression suite should cover:

- app version formatting
- resource/app path resolution logic
- update-checker parsing and decision logic using local mock payloads

This keeps the suite fast and stable while still proving that the repository has a real executable test command.

### Phase 1 CI Strategy

CI should target `windows-latest` first because the project is a Windows-first desktop application and depends on Windows packaging/runtime assumptions in several places.

The workflow should:

- check out the repository
- set up Python
- install `requirements.txt`
- run `python -m unittest discover -s tests -p "test_*.py" -v`

Packaging is intentionally out of scope for the first phase because trust here comes from a green validation signal, not from reproducing the full release pipeline.

## Phase 2: Smoother Collaboration

Phase 2 is planned but not implemented in this pass.

Deliverables:

- GitHub issue templates
- pull request template
- `CHANGELOG.md`
- `README.md` links to issue and pull request guidance

The purpose of Phase 2 is to reduce maintainer triage effort once outside usage increases.

## Phase 3: Better Project Presentation

Phase 3 is planned but not implemented in this pass.

Deliverables:

- screenshots or short GIFs
- FAQ / known limitations
- simple roadmap or support policy

The purpose of Phase 3 is to improve first impressions and reduce repeated user questions.

## File-Level Design For Phase 1

### `CONTRIBUTING.md`

A short bilingual contribution guide describing:

- what kinds of issues and PRs are welcome
- how to discuss larger changes before implementation
- the local setup and test command
- the maintainer's review expectations

It should intentionally avoid detailed branch naming, release-flow, or conventional-commit requirements.

### `SECURITY.md`

A concise bilingual security policy describing:

- which report channels to use
- what information to include in a report
- that public issues are discouraged for undisclosed security problems
- that response time is best-effort, not SLA-backed

### `CODE_OF_CONDUCT.md`

A short custom code of conduct is preferable to a long upstream template in this repository. It should emphasize respectful, professional collaboration and explicitly note that abusive behavior may result in issue/PR closure or participation limits.

### `tests/`

The first test files should stay independent from GUI windows and PDF fixture complexity:

- `tests/test_app_paths.py`
- `tests/test_app_version.py`
- `tests/test_update_checker.py`

These files should use `unittest` and local mocks only.

### `.github/workflows/ci.yml`

This workflow becomes the canonical automated validation entry point for contributors and for `README.md`.

### `README.md`

The README should gain a short collaboration section and should stop overstating current CI/test/build guarantees. Any mention of tests or GitHub Actions must point to the actual workflow and actual command added in Phase 1.

### `.gitignore`

The ignore file should stop ignoring repository-owned `docs/` and `tests/` content. Build outputs, local settings, caches, and temporary artifacts should remain ignored.

## Risks And Mitigations

### Risk: Governance files become too verbose

Mitigation: keep each file intentionally short and focused on the maintainer's actual workflow.

### Risk: CI becomes flaky or too expensive

Mitigation: limit the first suite to deterministic utility tests and a single Windows job.

### Risk: README and repo drift again later

Mitigation: make the Phase 1 workflow and test command the single documented source of truth, and keep future docs changes small and explicit.

## Acceptance Criteria

Phase 1 is complete when:

- the repository contains the three governance files
- the repository contains a real `tests/` directory with runnable tests
- the repository contains a GitHub Actions CI workflow
- `README.md` accurately describes the current collaboration and testing setup
- `.gitignore` no longer conflicts with tracked `docs/` and `tests/`

Phase 2 and Phase 3 remain documented future work and are not required for first completion.
