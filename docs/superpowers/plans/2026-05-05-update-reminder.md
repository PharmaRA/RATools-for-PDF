# Update Reminder Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a GitHub Releases based update checker with manual checks in the About dialog and daily startup reminders only for major updates.

**Architecture:** Keep version parsing, GitHub API parsing, and reminder policy in a focused `update_checker.py` module. Keep application version constants in `app_version.py`. Use `QThread` workers from `controller.py` so update checks never block the PySide6 UI, while `view.py` only owns display widgets and message dialogs.

**Tech Stack:** Python standard library, `unittest`, `unittest.mock`, PySide6, GitHub Releases API, existing `QSettings` persistence.

---

## File Structure

- Create: `app_version.py`
  - Owns canonical app version and GitHub repository constants.
- Create: `update_checker.py`
  - Owns pure version parsing/comparison, major-update policy, GitHub API request, and structured update result.
- Create: `test_update_checker.py`
  - Tests pure update logic and mocked network behavior.
- Modify: `view.py`
  - Replace hardcoded About version. Add About dialog widgets and methods for manual update status. Add a custom major-update prompt method.
- Modify: `controller.py`
  - Add update-check worker thread, manual check trigger, startup silent check trigger, and QSettings throttle/ignore handling.
- Modify: `main.py`
  - Trigger startup silent check after the main window is shown.
- Modify: `test_build_pyinstaller.py`
  - Keep version-info expectations aligned with `app_version.py`.
- Optional modify: `build_version_info.txt`
  - Align static metadata with the canonical version if tests expose mismatch.

Repository rule: do not create git commits unless the user explicitly requests a commit.

## Task 1: Canonical App Version

**Files:**
- Create: `app_version.py`
- Modify: `view.py`
- Test: `test_build_pyinstaller.py`

- [ ] **Step 1: Write the version module**

Create `app_version.py` with:

```python
APP_COMPANY = "PharmaRA"
APP_NAME = "RATools for PDF"
APP_REPOSITORY_OWNER = "PharmaRA"
APP_REPOSITORY_NAME = "RATools-for-PDF"
APP_VERSION = (0, 2, 5, 0)
APP_VERSION_STR = "0.2.5.0"


def get_display_version():
    return f"Version {APP_VERSION_STR}"
```

- [ ] **Step 2: Run the existing version tests**

Run: `python -m unittest test_build_pyinstaller.py -v`

Expected before fixes: tests may fail if `build_version_info.txt` does not match `APP_VERSION_STR` exactly.

- [ ] **Step 3: Replace the hardcoded About version**

In `view.py`, add the import near the other local imports:

```python
from app_version import get_display_version
```

Change `AboutDialog.__init__` from:

```python
version_badge = QLabel("Version 1.0.0")
```

to:

```python
version_badge = QLabel(get_display_version())
```

- [ ] **Step 4: Align static build metadata if needed**

If `test_build_pyinstaller.py` fails because `ProductVersion` is `0.2.50`, change `build_version_info.txt` line 23 to:

```text
          StringStruct('ProductVersion', '0.2.5.0')
```

- [ ] **Step 5: Verify version tests pass**

Run: `python -m unittest test_build_pyinstaller.py -v`

Expected: `OK`.

## Task 2: Update Checker Pure Logic

**Files:**
- Create: `update_checker.py`
- Create: `test_update_checker.py`

- [ ] **Step 1: Write failing tests for version parsing and comparison**

Create `test_update_checker.py` with:

```python
import unittest
from unittest.mock import patch

import update_checker


class UpdateCheckerVersionTests(unittest.TestCase):
    def test_parse_version_accepts_plain_and_v_prefixed_tags(self):
        self.assertEqual(update_checker.parse_version_tag("0.2.5"), (0, 2, 5))
        self.assertEqual(update_checker.parse_version_tag("v0.2.5"), (0, 2, 5))

    def test_parse_version_rejects_invalid_tags(self):
        with self.assertRaises(ValueError):
            update_checker.parse_version_tag("release-latest")

    def test_is_newer_version_compares_tuples(self):
        self.assertTrue(update_checker.is_newer_version((0, 2, 6), (0, 2, 5)))
        self.assertTrue(update_checker.is_newer_version((0, 3, 0), (0, 2, 9)))
        self.assertFalse(update_checker.is_newer_version((0, 2, 5), (0, 2, 5)))
        self.assertFalse(update_checker.is_newer_version((0, 2, 4), (0, 2, 5)))


class UpdateCheckerMajorRuleTests(unittest.TestCase):
    def test_major_update_uses_minor_bump_during_zero_major_stage(self):
        release = update_checker.ReleaseInfo(
            version=(0, 3, 0),
            version_text="0.3.0",
            title="v0.3.0",
            body="",
            html_url="https://github.com/PharmaRA/RATools-for-PDF/releases/tag/v0.3.0",
            published_at="2026-05-05T00:00:00Z",
        )
        self.assertTrue(update_checker.is_major_update((0, 2, 5), release))

    def test_patch_bump_is_not_major_during_zero_major_stage(self):
        release = update_checker.ReleaseInfo(
            version=(0, 2, 6),
            version_text="0.2.6",
            title="v0.2.6",
            body="",
            html_url="https://github.com/PharmaRA/RATools-for-PDF/releases/tag/v0.2.6",
            published_at="2026-05-05T00:00:00Z",
        )
        self.assertFalse(update_checker.is_major_update((0, 2, 5), release))

    def test_major_update_uses_major_bump_after_one_dot_zero(self):
        release = update_checker.ReleaseInfo(
            version=(2, 0, 0),
            version_text="2.0.0",
            title="v2.0.0",
            body="",
            html_url="https://github.com/PharmaRA/RATools-for-PDF/releases/tag/v2.0.0",
            published_at="2026-05-05T00:00:00Z",
        )
        self.assertTrue(update_checker.is_major_update((1, 4, 2), release))

    def test_minor_bump_is_not_major_after_one_dot_zero(self):
        release = update_checker.ReleaseInfo(
            version=(1, 5, 0),
            version_text="1.5.0",
            title="v1.5.0",
            body="",
            html_url="https://github.com/PharmaRA/RATools-for-PDF/releases/tag/v1.5.0",
            published_at="2026-05-05T00:00:00Z",
        )
        self.assertFalse(update_checker.is_major_update((1, 4, 2), release))

    def test_major_update_marker_forces_major_update(self):
        release = update_checker.ReleaseInfo(
            version=(0, 2, 6),
            version_text="0.2.6",
            title="v0.2.6 [major-update]",
            body="",
            html_url="https://github.com/PharmaRA/RATools-for-PDF/releases/tag/v0.2.6",
            published_at="2026-05-05T00:00:00Z",
        )
        self.assertTrue(update_checker.is_major_update((0, 2, 5), release))


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m unittest test_update_checker.py -v`

Expected: fail with `ModuleNotFoundError: No module named 'update_checker'`.

- [ ] **Step 3: Implement minimal pure update logic**

Create `update_checker.py` with:

```python
import json
import urllib.error
import urllib.request
from dataclasses import dataclass

from app_version import APP_REPOSITORY_NAME, APP_REPOSITORY_OWNER, APP_VERSION

MAJOR_UPDATE_MARKER = "[major-update]"
GITHUB_API_URL = (
    f"https://api.github.com/repos/{APP_REPOSITORY_OWNER}/{APP_REPOSITORY_NAME}/releases/latest"
)


@dataclass(frozen=True)
class ReleaseInfo:
    version: tuple[int, int, int]
    version_text: str
    title: str
    body: str
    html_url: str
    published_at: str


@dataclass(frozen=True)
class UpdateCheckResult:
    ok: bool
    current_version: str
    latest_release: ReleaseInfo | None = None
    has_update: bool = False
    is_major: bool = False
    error: str = ""


def normalize_version(version):
    parts = tuple(int(part) for part in version)
    if len(parts) < 3:
        raise ValueError("version must contain at least three numeric parts")
    return parts[:3]


def version_to_text(version):
    return ".".join(str(part) for part in normalize_version(version))


def parse_version_tag(tag_name):
    value = str(tag_name).strip()
    if value.lower().startswith("v"):
        value = value[1:]
    parts = value.split(".")
    if len(parts) < 3:
        raise ValueError(f"invalid version tag: {tag_name}")
    try:
        parsed = tuple(int(part) for part in parts[:3])
    except ValueError as exc:
        raise ValueError(f"invalid version tag: {tag_name}") from exc
    return parsed


def is_newer_version(remote_version, current_version):
    return normalize_version(remote_version) > normalize_version(current_version)


def has_major_update_marker(release):
    marker_source = f"{release.title}\n{release.body}".lower()
    return MAJOR_UPDATE_MARKER in marker_source


def is_major_update(current_version, release):
    current = normalize_version(current_version)
    remote = normalize_version(release.version)
    if not is_newer_version(remote, current):
        return False
    if has_major_update_marker(release):
        return True
    if current[0] == 0:
        return remote[0] > current[0] or remote[1] > current[1]
    return remote[0] > current[0]
```

- [ ] **Step 4: Run pure logic tests**

Run: `python -m unittest test_update_checker.py -v`

Expected: `OK` for the tests added so far.

## Task 3: GitHub Release Fetching

**Files:**
- Modify: `update_checker.py`
- Modify: `test_update_checker.py`

- [ ] **Step 1: Add mocked network tests**

Append to `test_update_checker.py` before the `if __name__ == "__main__"` block:

```python
class FakeResponse:
    def __init__(self, payload):
        self.payload = payload

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        return False

    def read(self):
        return self.payload


class UpdateCheckerFetchTests(unittest.TestCase):
    def test_fetch_latest_release_parses_github_response(self):
        payload = {
            "tag_name": "v0.3.0",
            "name": "v0.3.0 Important Release",
            "body": "[major-update] Important fixes",
            "html_url": "https://github.com/PharmaRA/RATools-for-PDF/releases/tag/v0.3.0",
            "published_at": "2026-05-05T00:00:00Z",
            "draft": False,
            "prerelease": False,
        }
        with patch.object(update_checker.urllib.request, "urlopen", return_value=FakeResponse(json.dumps(payload).encode("utf-8"))):
            release = update_checker.fetch_latest_release(timeout=1)

        self.assertEqual(release.version, (0, 3, 0))
        self.assertEqual(release.version_text, "0.3.0")
        self.assertIn("Important Release", release.title)

    def test_check_for_updates_returns_structured_error_on_network_failure(self):
        with patch.object(update_checker.urllib.request, "urlopen", side_effect=OSError("offline")):
            result = update_checker.check_for_updates(current_version=(0, 2, 5), timeout=1)

        self.assertFalse(result.ok)
        self.assertFalse(result.has_update)
        self.assertIn("offline", result.error)
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m unittest test_update_checker.py -v`

Expected: fail with missing `fetch_latest_release` and `check_for_updates`.

- [ ] **Step 3: Implement GitHub API fetching and structured check**

Append to `update_checker.py`:

```python
def release_from_github_payload(payload):
    if payload.get("draft") or payload.get("prerelease"):
        raise ValueError("latest release is not a stable release")

    version = parse_version_tag(payload["tag_name"])
    title = payload.get("name") or payload.get("tag_name") or version_to_text(version)
    return ReleaseInfo(
        version=version,
        version_text=version_to_text(version),
        title=str(title),
        body=str(payload.get("body") or ""),
        html_url=str(payload.get("html_url") or ""),
        published_at=str(payload.get("published_at") or ""),
    )


def fetch_latest_release(timeout=8):
    request = urllib.request.Request(
        GITHUB_API_URL,
        headers={
            "Accept": "application/vnd.github+json",
            "User-Agent": "RATools-for-PDF",
        },
    )
    with urllib.request.urlopen(request, timeout=timeout) as response:
        payload = json.loads(response.read().decode("utf-8"))
    return release_from_github_payload(payload)


def check_for_updates(current_version=APP_VERSION, timeout=8):
    current = normalize_version(current_version)
    try:
        release = fetch_latest_release(timeout=timeout)
        has_update = is_newer_version(release.version, current)
        return UpdateCheckResult(
            ok=True,
            current_version=version_to_text(current),
            latest_release=release,
            has_update=has_update,
            is_major=is_major_update(current, release) if has_update else False,
        )
    except (OSError, urllib.error.URLError, urllib.error.HTTPError, ValueError, KeyError, json.JSONDecodeError) as exc:
        return UpdateCheckResult(
            ok=False,
            current_version=version_to_text(current),
            error=str(exc),
        )
```

- [ ] **Step 4: Run update checker tests**

Run: `python -m unittest test_update_checker.py -v`

Expected: `OK`.

## Task 4: About Dialog Manual Update UI

**Files:**
- Modify: `view.py`
- Modify: `controller.py`

- [ ] **Step 1: Extend AboutDialog widgets**

In `view.py`, update `AboutDialog.__init__` after the tech card is added and before `self.content_layout.addStretch()`:

```python
        update_card = QFrame()
        update_card.setObjectName("aboutInfoCard")
        update_layout = QVBoxLayout(update_card)
        update_layout.setContentsMargins(16, 14, 16, 14)
        update_layout.setSpacing(8)

        update_title = QLabel("更新")
        update_title.setObjectName("aboutTitle")
        self.update_status_label = QLabel("可手动检查 GitHub Releases 中的最新版本。")
        self.update_status_label.setWordWrap(True)
        self.update_status_label.setObjectName("aboutText")

        update_btn_layout = QHBoxLayout()
        self.btn_check_updates = QPushButton("检查更新")
        self.btn_check_updates.setObjectName("dialogPrimaryBtn")
        self.btn_open_release = QPushButton("打开发布页")
        self.btn_open_release.setObjectName("dialogSecondaryBtn")
        self.btn_open_release.hide()
        update_btn_layout.addWidget(self.btn_check_updates)
        update_btn_layout.addWidget(self.btn_open_release)
        update_btn_layout.addStretch()

        update_layout.addWidget(update_title)
        update_layout.addWidget(self.update_status_label)
        update_layout.addLayout(update_btn_layout)
        self.content_layout.addWidget(update_card)
```

- [ ] **Step 2: Add AboutDialog helper methods**

Add methods inside `AboutDialog`:

```python
    def set_update_checking(self):
        self.btn_check_updates.setEnabled(False)
        self.update_status_label.setText("正在检查更新...")
        self.btn_open_release.hide()

    def set_update_result(self, message, release_url=""):
        self.btn_check_updates.setEnabled(True)
        self.update_status_label.setText(message)
        self.latest_release_url = release_url
        self.btn_open_release.setVisible(bool(release_url))
```

- [ ] **Step 3: Add controller worker skeleton**

In `controller.py`, add imports:

```python
import webbrowser
from PySide6.QtCore import QObject, QThread, Signal, Qt, QTimer
import update_checker
```

Add a worker class before `MainController`:

```python
class UpdateCheckWorker(QThread):
    finished_check = Signal(object, bool)

    def __init__(self, silent=False):
        super().__init__()
        self.silent = silent

    def run(self):
        result = update_checker.check_for_updates()
        self.finished_check.emit(result, self.silent)
```

- [ ] **Step 4: Wire manual check signals**

In `MainController.setup_connections`, add after log button connection:

```python
        self.view.btn_top_about.clicked.connect(self._wire_about_dialog_updates)
```

Add methods to `MainController`:

```python
    def _wire_about_dialog_updates(self):
        if not hasattr(self.view, "about_dialog"):
            return
        dialog = self.view.about_dialog
        if getattr(dialog, "_update_connections_ready", False):
            return
        dialog.btn_check_updates.clicked.connect(self.check_updates_manually)
        dialog.btn_open_release.clicked.connect(lambda: self.open_release_url(getattr(dialog, "latest_release_url", "")))
        dialog._update_connections_ready = True

    def check_updates_manually(self):
        self.view.show_about_dialog()
        self._wire_about_dialog_updates()
        self.view.about_dialog.set_update_checking()
        self._start_update_check(silent=False)

    def _start_update_check(self, silent=False):
        if getattr(self, "update_worker", None) and self.update_worker.isRunning():
            return
        self.update_worker = UpdateCheckWorker(silent=silent)
        self.update_worker.finished_check.connect(self._handle_update_result)
        self.update_worker.start()

    def open_release_url(self, url):
        if not url:
            return
        try:
            webbrowser.open(url)
        except Exception as exc:
            self.view.show_error_message("打开失败", f"无法打开发布页：{exc}")
```

- [ ] **Step 5: Add manual result handling**

Add to `MainController`:

```python
    def _handle_update_result(self, result, silent):
        if silent:
            self._handle_silent_update_result(result)
            return
        dialog = getattr(self.view, "about_dialog", None)
        if dialog is None:
            return
        if not result.ok:
            dialog.set_update_result(f"检查更新失败：{result.error}")
            return
        release = result.latest_release
        if not result.has_update or release is None:
            dialog.set_update_result(f"当前已是最新版本：{result.current_version}")
            return
        message = (
            f"发现新版本 {release.version_text}\n"
            f"当前版本：{result.current_version}\n"
            f"发布标题：{release.title}\n"
            f"发布时间：{release.published_at or '未知'}"
        )
        dialog.set_update_result(message, release.html_url)
```

## Task 5: Startup Silent Major Update Reminder

**Files:**
- Modify: `view.py`
- Modify: `controller.py`
- Modify: `main.py`

- [ ] **Step 1: Add a three-choice major update prompt**

In `view.py`, add to imports:

```python
from PySide6.QtWidgets import QDialogButtonBox
```

If avoiding `QDialogButtonBox`, add this simpler custom prompt method to `MainWindow`:

```python
    def show_major_update_prompt(self, current_version, release):
        dlg = FramelessDraggableDialog("重要更新可用", self)
        dlg.resize(460, 260)
        message = QLabel(
            f"检测到重要更新 {release.version_text}\n"
            f"当前版本：{current_version}\n"
            f"发布标题：{release.title}\n"
            f"发布时间：{release.published_at or '未知'}"
        )
        message.setWordWrap(True)
        message.setObjectName("aboutText")
        dlg.content_layout.addWidget(message)
        dlg.result_action = "later"

        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        btn_open = QPushButton("查看更新")
        btn_open.setObjectName("dialogPrimaryBtn")
        btn_later = QPushButton("稍后提醒")
        btn_later.setObjectName("dialogSecondaryBtn")
        btn_ignore = QPushButton("忽略此版本")
        btn_ignore.setObjectName("dialogSecondaryBtn")

        def choose(action):
            dlg.result_action = action
            dlg.accept()

        btn_open.clicked.connect(lambda: choose("open"))
        btn_later.clicked.connect(lambda: choose("later"))
        btn_ignore.clicked.connect(lambda: choose("ignore"))
        btn_layout.addWidget(btn_ignore)
        btn_layout.addWidget(btn_later)
        btn_layout.addWidget(btn_open)
        dlg.content_layout.addLayout(btn_layout)
        dlg.exec()
        return dlg.result_action
```

- [ ] **Step 2: Add silent-check date helpers**

In `controller.py`, add import:

```python
from datetime import date, datetime
```

Replace the existing `from datetime import datetime` with that combined import.

Add to `MainController`:

```python
    def check_updates_on_startup(self):
        if not hasattr(self.view, "app_settings"):
            return
        today = date.today().isoformat()
        if str(self.view.app_settings.value("Update/LastSilentCheckDate", "")) == today:
            return
        self.view.app_settings.setValue("Update/LastSilentCheckDate", today)
        self._start_update_check(silent=True)

    def _handle_silent_update_result(self, result):
        if not result.ok or not result.has_update or not result.is_major or result.latest_release is None:
            return
        release = result.latest_release
        ignored_version = str(self.view.app_settings.value("Update/IgnoredVersion", "") or "")
        if ignored_version == release.version_text:
            return
        today = date.today().isoformat()
        prompted_key = "Update/LastPromptedVersion"
        prompted_value = str(self.view.app_settings.value(prompted_key, "") or "")
        if prompted_value == f"{today}:{release.version_text}":
            return
        self.view.app_settings.setValue(prompted_key, f"{today}:{release.version_text}")
        action = self.view.show_major_update_prompt(result.current_version, release)
        if action == "open":
            self.open_release_url(release.html_url)
        elif action == "ignore":
            self.view.app_settings.setValue("Update/IgnoredVersion", release.version_text)
```

- [ ] **Step 3: Trigger startup check after window show**

In `main.py`, change the end of the entry block from:

```python
    view.show()
    sys.exit(app.exec())
```

to:

```python
    view.show()
    QTimer.singleShot(0, controller.check_updates_on_startup)
    sys.exit(app.exec())
```

Also update imports:

```python
from PySide6.QtCore import Qt, QTimer
```

## Task 6: Verification

**Files:**
- All touched files

- [ ] **Step 1: Run focused logic tests**

Run: `python -m unittest test_update_checker.py -v`

Expected: `OK`.

- [ ] **Step 2: Run build/version tests**

Run: `python -m unittest test_build_pyinstaller.py -v`

Expected: `OK`.

- [ ] **Step 3: Run bootstrap tests**

Run: `python -m unittest test_main.py -v`

Expected: `OK`.

- [ ] **Step 4: Run the complete test suite**

Run: `python -m unittest discover -v`

Expected: all tests pass.

- [ ] **Step 5: Manual smoke test**

Run: `python main.py`

Expected manual checks:

- App opens without blocking during startup.
- About dialog shows `Version 0.2.5.0`.
- About dialog has `检查更新`.
- Clicking `检查更新` updates status text instead of freezing UI.
- If GitHub is unreachable, the About dialog shows a readable failure message.

## Self-Review

- Spec coverage: version source, GitHub Releases source, major version rule, `[major-update]` marker, manual About check, startup silent check, QSettings throttle, ignore-version behavior, and error handling are all covered by tasks.
- Placeholder scan: no incomplete placeholder markers or undefined future task references remain.
- Type consistency: `ReleaseInfo`, `UpdateCheckResult`, `version_text`, `html_url`, and controller method names are used consistently across tasks.
- Repository policy: plan explicitly avoids commit steps because commits require explicit user request in this environment.
