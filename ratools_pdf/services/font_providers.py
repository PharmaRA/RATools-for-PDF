"""Font workflow v1: read-only inspection and truthful manual Acrobat handoff.

The retired COM fixup is deliberately absent. Installation detection is only an
opening hint; it never qualifies an automatic repair engine or a Preflight profile.
"""

from dataclasses import dataclass
import platform
from typing import Mapping, Optional, Protocol

from ratools_pdf.pdf import inspect as pdf_inspect
from ratools_pdf.services import system_shell


AUTOMATION_RETIRED_REASON = (
    "自动字体修复已停用：Acrobat COM/OLE 执行 Preflight fixup 不可靠；"
    "安装 Acrobat 不代表自动修复或本机 profile 已通过验证。"
)


@dataclass(frozen=True)
class FontCapability:
    platform: str
    provider: str
    version: str  # Adapter contract version, not the installed Acrobat version.
    external_version: Optional[str]
    availability: str
    supported_operations: tuple
    unavailable_reasons: Mapping[str, str]
    acrobat_detected: bool
    external_qualification: str


@dataclass(frozen=True)
class FontInspection:
    status: str
    reason: str
    unembedded_count: Optional[int] = None
    substitution_count: Optional[int] = None
    unknown_count: Optional[int] = None
    findings: tuple = ()

    @property
    def verified(self):
        return self.status == "verified"


@dataclass(frozen=True)
class FontOperationResult:
    status: str
    reason: str
    opened: bool = False
    provider_reported_success: bool = False
    inspection: Optional[FontInspection] = None

    @property
    def verified(self):
        return bool(self.status == "verified" and self.inspection and self.inspection.verified)


class FontProvider(Protocol):
    """Backend attempts are untrusted until the workflow independently rechecks."""

    def capability(self) -> FontCapability: ...

    def repair(self, pdf_path) -> FontOperationResult: ...

    def open_manual(self, pdf_path) -> FontOperationResult: ...


class AcrobatManualProvider:
    """Existing shell handoff only; no COM, fixup, rasterization or PDF writes."""

    def __init__(self, *, platform_name=None):
        self.platform = platform_name or platform.system()
        self.acrobat_path = ""
        self._discovery_error = ""
        if self.platform == "Windows":
            try:
                self.acrobat_path = system_shell.find_acrobat_executable()
            except Exception as exc:
                self._discovery_error = type(exc).__name__

    def capability(self):
        supported = self.platform in {"Windows", "Darwin", "Linux"}
        operations = ("inspect", "recheck", "open_manual") if supported else ("inspect", "recheck")
        reasons = {"repair": AUTOMATION_RETIRED_REASON}
        if not supported:
            reasons["open_manual"] = "当前平台不支持 Acrobat/默认 PDF 程序手动交接。"
        if not self.acrobat_path:
            reasons["acrobat"] = "未定位到 Acrobat；默认 PDF 程序可能不具备印前检查修复功能。"
            if self._discovery_error:
                reasons["acrobat"] += f" 探测失败（{self._discovery_error}）。"
        return FontCapability(
            platform=self.platform, provider="acrobat_manual", version="1",
            external_version=None,
            availability=("manual_only" if self.acrobat_path else "default_app_only") if supported else "unsupported",
            supported_operations=operations, unavailable_reasons=reasons,
            acrobat_detected=bool(self.acrobat_path), external_qualification="pending",
        )

    def repair(self, pdf_path):
        return FontOperationResult("blocked", AUTOMATION_RETIRED_REASON)

    def open_manual(self, pdf_path):
        capability = self.capability()
        if "open_manual" not in capability.supported_operations:
            return FontOperationResult("unavailable", capability.unavailable_reasons["open_manual"])
        system_shell.open_pdf_in_acrobat_or_default(str(pdf_path), self.acrobat_path)
        return FontOperationResult(
            "manual_required", "已交接打开请求；请手动修复并保存，再执行字体复检。", opened=True,
        )


class FontWorkflow:
    """Unified capability/inspect/repair/recheck contract consumed by the UI.

    Only the manual provider is shipped. A backend's reported_success is an
    attempt, never verification. Recheck always uses the public PDF inspector.
    """

    def __init__(self, provider=None):
        self.provider = provider if provider is not None else AcrobatManualProvider()

    def capability(self):
        return self.provider.capability()

    def inspect(self, pdf_path):
        try:
            raw = pdf_inspect.collect_font_precheck_for_path(pdf_path)
            if not raw.get("available"):
                return FontInspection("unavailable", "字体预检不可用：PDF 无法读取或需要密码。")
            findings = tuple(raw["font_findings"])
            counts = tuple(sum(bool(item.get(flag)) for item in findings) for flag in (
                "has_unembedded", "substitution_risk", "has_unknown_embedding_status",
            ))
            risky = pdf_inspect.font_precheck_has_embedding_risk(raw)
            return FontInspection(
                "risks_remain" if risky else "verified",
                "字体风险仍存在。" if risky else "本次字体复检未发现三类字体风险。",
                *counts, findings=findings,
            )
        except TimeoutError:
            return FontInspection("timeout", "字体预检超时，未完成复检。")
        except Exception as exc:
            return FontInspection("unavailable", f"字体预检失败（{type(exc).__name__}），未完成复检。")

    def recheck(self, pdf_path):
        return self.inspect(pdf_path)

    def open_manual(self, pdf_path):
        try:
            capability = self.capability()
            if "open_manual" not in capability.supported_operations:
                return FontOperationResult(
                    "unavailable", capability.unavailable_reasons.get("open_manual", "手动交接不可用。"),
                )
            return self.provider.open_manual(pdf_path)
        except TimeoutError:
            return FontOperationResult("timeout", "打开请求超时；未确认 PDF 已打开或修复。")
        except Exception as exc:
            return FontOperationResult("failed", f"打开请求失败（{type(exc).__name__}）；未修复 PDF。")

    def repair(self, pdf_path):
        try:
            capability = self.capability()
            if "repair" not in capability.supported_operations:
                return FontOperationResult(
                    "blocked", capability.unavailable_reasons.get("repair", "此 provider 不支持自动字体修复。"),
                )
            attempt = self.provider.repair(pdf_path)
            if attempt.status != "reported_success":
                return FontOperationResult(attempt.status, attempt.reason)
            inspection = self.recheck(pdf_path)
            return FontOperationResult(
                "verified" if inspection.verified else "recheck_failed",
                inspection.reason, provider_reported_success=True, inspection=inspection,
            )
        except TimeoutError:
            return FontOperationResult("timeout", "字体 provider 超时，未验证修复结果。")
        except Exception as exc:
            return FontOperationResult("failed", f"字体 provider 失败（{type(exc).__name__}），未验证修复结果。")
