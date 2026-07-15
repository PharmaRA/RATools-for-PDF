from dataclasses import dataclass
import json
import re
from typing import Optional, Tuple
import urllib.error
import urllib.request

from ratools_pdf.config.version import APP_REPOSITORY_NAME, APP_REPOSITORY_OWNER, APP_VERSION


MAJOR_UPDATE_MARKER = "[major-update]"
GITHUB_API_URL = (
    f"https://api.github.com/repos/{APP_REPOSITORY_OWNER}/{APP_REPOSITORY_NAME}/releases"
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


def select_latest_release(payloads):
    """从 releases 列表中挑选语义版本号最大的稳定发布。

    GitHub 的 /releases/latest 端点依据发布时间或手动标记判定 "latest"，
    并不保证返回语义版本号最大的发布，因此改为拉取完整列表后自行比较，
    避免出现「已有 v0.7.1 却仍提示 v0.7.0」的问题。
    """
    latest = None
    for payload in payloads:
        try:
            release = release_from_github_payload(payload)
        except (ValueError, KeyError):
            continue
        if latest is None or is_newer_version(release.version, latest.version):
            latest = release
    if latest is None:
        raise ValueError("No stable release found")
    return latest


def fetch_latest_release(timeout=8):
    request = urllib.request.Request(
        GITHUB_API_URL,
        headers={
            "Accept": "application/vnd.github+json",
            "User-Agent": "RATools-for-PDF",
        },
    )
    with urllib.request.urlopen(request, timeout=timeout) as response:
        payloads = json.loads(response.read().decode("utf-8"))
    return select_latest_release(payloads)


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
