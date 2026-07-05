from dataclasses import dataclass
import json
import re
from typing import Optional, Tuple
import urllib.error
import urllib.request

from ratools_pdf.config.version import APP_REPOSITORY_NAME, APP_REPOSITORY_OWNER, APP_VERSION


MAJOR_UPDATE_MARKER = "[major-update]"
GITHUB_API_URL = (
    f"https://api.github.com/repos/{APP_REPOSITORY_OWNER}/{APP_REPOSITORY_NAME}/releases/latest"
)


@dataclass(frozen=True)
class ReleaseInfo:
    version: Tuple[int, int, int]
    version_text: str
    title: str
    body: str
    html_url: str
    published_at: str


@dataclass(frozen=True)
class UpdateCheckResult:
    ok: bool
    current_version: str
    latest_release: Optional[ReleaseInfo]
    has_update: bool
    is_major: bool
    error: str


def normalize_version(version):
    return tuple(version[:3])


def version_to_text(version):
    return ".".join(str(part) for part in normalize_version(version))


def parse_version_tag(tag):
    match = re.fullmatch(r"v?(\d+)\.(\d+)\.(\d+)", tag)
    if not match:
        raise ValueError(f"Invalid version tag: {tag}")
    return tuple(int(part) for part in match.groups())


def is_newer_version(candidate_version, current_version):
    return normalize_version(candidate_version) > normalize_version(current_version)


def has_major_update_marker(release):
    return MAJOR_UPDATE_MARKER in release.title or MAJOR_UPDATE_MARKER in release.body


def is_major_update(current_version, latest_release):
    current = normalize_version(current_version)
    latest = normalize_version(latest_release.version)

    if not is_newer_version(latest, current):
        return False

    if has_major_update_marker(latest_release):
        return True

    if current[0] == 0:
        return latest[0] > 0 or latest[1] > current[1]

    return latest[0] > current[0]


def release_from_github_payload(payload):
    if payload["draft"] or payload["prerelease"]:
        raise ValueError("Unstable release")

    version = parse_version_tag(payload["tag_name"])
    version_text = version_to_text(version)

    return ReleaseInfo(
        version=version,
        version_text=version_text,
        title=payload.get("name") or payload.get("tag_name") or version_text,
        body=payload.get("body") or "",
        html_url=payload.get("html_url") or "",
        published_at=payload.get("published_at") or "",
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
    current_text = version_to_text(current)
    try:
        latest_release = fetch_latest_release(timeout=timeout)
        has_update = is_newer_version(latest_release.version, current)
        return UpdateCheckResult(
            ok=True,
            current_version=current_text,
            latest_release=latest_release,
            has_update=has_update,
            is_major=is_major_update(current, latest_release) if has_update else False,
            error="",
        )
    except (
        OSError,
        urllib.error.URLError,
        urllib.error.HTTPError,
        ValueError,
        KeyError,
        json.JSONDecodeError,
    ) as exc:
        return UpdateCheckResult(
            ok=False,
            current_version=current_text,
            latest_release=None,
            has_update=False,
            is_major=False,
            error=str(exc),
        )
