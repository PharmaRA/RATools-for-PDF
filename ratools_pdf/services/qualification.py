"""Versioned, read-only PDF qualification evidence. No document contents are saved."""

import hashlib
import json
from pathlib import Path, PureWindowsPath
import platform
import re
from datetime import datetime, timezone

import fitz

from ratools_pdf.pdf.inspect import read_pdf_header_version, pdf_has_signature
from ratools_pdf.pdf.qpdf import check_pdf_syntax, run_qpdf_command
from ratools_pdf.services.file_safety import check_unlinked_path

MANIFEST_SCHEMA_VERSION = "1.0"
REPORT_SCHEMA_VERSION = "1.0"
STATUSES = ("passed", "failed", "pending", "blocked")
EVIDENCE_LEVELS = (
    "automated_regression", "windows_release", "real_business_sample", "external_reader",
)


def _sha256(path):
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _status(statuses):
    values = set(statuses)
    for value in ("failed", "blocked", "pending"):
        if value in values:
            return value
    return "passed" if values else "pending"


def _unique_object(pairs):
    result = {}
    for key, value in pairs:
        if key in result:
            raise ValueError("Duplicate qualification JSON key")
        result[key] = value
    return result


def _validate_sample_history(manifest):
    history = manifest.get("sample_history")
    if history is None:
        if any("sample_version" in sample or "expected_metrics_version" in sample for sample in manifest["samples"]):
            raise ValueError("Versioned samples require sample_history")
        return
    if not isinstance(history, dict) or history.get("schema_version") != "1.0":
        raise ValueError("Unsupported sample history version")
    records = history.get("records")
    if not isinstance(records, list) or not records:
        raise ValueError("Sample history records required")
    samples = {sample["id"]: sample for sample in manifest["samples"]}
    revisions, metrics = {}, {}
    for record in records:
        if not isinstance(record, dict) or record.get("sample_id") not in samples:
            raise ValueError("History must reference declared sample IDs")
        for field in ("sample_version", "expected_metrics_version", "reason", "source_authorization"):
            if not isinstance(record.get(field), str) or not record[field].strip():
                raise ValueError("History versions, reason and source authorization required")
        digest = record.get("sha256")
        if digest is not None and (not isinstance(digest, str) or not re.fullmatch(r"[0-9a-f]{64}", digest)):
            raise ValueError("History requires a SHA256 or null for an unavailable sample")
        if not isinstance(record.get("expected"), dict):
            raise ValueError("History requires an expected metrics snapshot")
        key = (record["sample_id"], record["sample_version"])
        if key in revisions:
            raise ValueError("Duplicate sample history revision")
        revisions[key] = record
        metric_key = (record["sample_id"], record["expected_metrics_version"])
        if metric_key in metrics and metrics[metric_key] != record["expected"]:
            raise ValueError("Changed metrics require a new expected_metrics_version")
        metrics[metric_key] = record["expected"]
        approval = record.get("approval")
        if not isinstance(approval, dict) or approval.get("status") not in {"pending", "approved", "rejected"}:
            raise ValueError("Explicit human approval status required")
        if approval["status"] != "pending" and (not isinstance(approval.get("reference"), str) or not approval["reference"].strip()):
            raise ValueError("Human approval/rejection requires an independent record reference")
    for sample in samples.values():
        if not isinstance(sample.get("sample_version"), str) or not isinstance(sample.get("expected_metrics_version"), str):
            raise ValueError("Each sample requires explicit sample and metric versions")
        current = revisions.get((sample["id"], sample["sample_version"]))
        if current is None or any(current.get(field) != sample.get(field) for field in
                ("sha256", "expected", "expected_metrics_version")) or current["source_authorization"] != sample["authorization"]:
            raise ValueError("Current sample hash, metrics or authorization differ from history")


def _load_manifest(path):
    manifest = json.loads(check_unlinked_path(path).read_text(encoding="utf-8"), object_pairs_hook=_unique_object)
    if not isinstance(manifest, dict) or manifest.get("schema_version") != MANIFEST_SCHEMA_VERSION:
        raise ValueError("Unsupported qualification manifest schema_version")
    if not isinstance(manifest.get("matrix_version"), str) or not manifest["matrix_version"]:
        raise ValueError("matrix_version is required")
    samples = manifest.get("samples")
    matrix = manifest.get("matrix")
    if not isinstance(samples, list) or not samples or not isinstance(matrix, list) or not matrix:
        raise ValueError("A nonempty samples list and matrix are required")
    ids = set()
    for sample in samples:
        if not isinstance(sample, dict):
            raise ValueError("Invalid sample entry")
        sample_id = sample.get("id", "")
        if not isinstance(sample_id, str) or not re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9_.-]{0,99}", sample_id) or sample_id in ids:
            raise ValueError("Sample IDs must be unique portable identifiers")
        ids.add(sample_id)
        source = sample.get("source")
        if not isinstance(source, dict) or source.get("type") not in {"synthetic", "real"}:
            raise ValueError("Each sample requires an honest source type")
        if not isinstance(source.get("tool"), str) or not source["tool"]:
            raise ValueError("Source tool is required")
        if not isinstance(source.get("version"), str):
            raise ValueError("Source version must be a string (pending when unknown)")
        if not isinstance(sample.get("authorization"), str) or not sample["authorization"]:
            raise ValueError("Authorization/non-sensitive declaration is required")
        if not isinstance(sample.get("features"), list) or not isinstance(sample.get("processing_options"), dict):
            raise ValueError("Features and processing_options are required")
        if not isinstance(sample.get("expected"), dict):
            raise ValueError("Expected metrics are required")
        if sample.get("path") is not None and not isinstance(sample["path"], str):
            raise ValueError("Sample path must be relative or null")
    row_ids = set()
    for row in matrix:
        if not isinstance(row, dict) or not isinstance(row.get("id"), str) or row["id"] in row_ids:
            raise ValueError("Matrix IDs must be unique")
        row_ids.add(row["id"])
        if row.get("evidence_level") not in EVIDENCE_LEVELS:
            raise ValueError("Unknown evidence level")
        if not isinstance(row.get("sample_ids"), list) or not row["sample_ids"] or any(s not in ids for s in row["sample_ids"]):
            raise ValueError("Matrix must reference declared samples")
    _validate_sample_history(manifest)
    return manifest


def _sample_path(root, relative):
    # Reject both POSIX and Windows absolute paths, traversal, drive-relative paths and ADS.
    portable = relative.replace("\\", "/")
    if not portable or portable.startswith("/") or PureWindowsPath(relative).drive or ":" in portable:
        raise ValueError("sample_path_containment")
    if ".." in Path(portable).parts:
        raise ValueError("sample_path_containment")
    candidate = check_unlinked_path(root / portable)
    if not candidate.resolve().is_relative_to(root.resolve()):
        raise ValueError("sample_path_containment")
    return candidate


def _inspect(path, expected):
    checks = {}
    with fitz.open(path) as document:
        checks["open"] = {"status": "passed", "repaired": bool(document.is_repaired)}
        if document.needs_pass:
            checks["structure"] = {"status": "blocked", "reason": "password_required"}
            checks["render"] = {"status": "blocked", "reason": "password_required"}
            return checks
        metrics = {
            "page_count": document.page_count,
            "pdf_version": read_pdf_header_version(path),
            "signature_present": bool(pdf_has_signature(path)),
            "image_count": 0, "annotation_count": 0, "widget_count": 0,
            "link_count": 0, "text_char_count": 0,
        }
        for page in document:
            metrics["image_count"] += len(page.get_images(full=True))
            metrics["annotation_count"] += sum(1 for _ in page.annots())
            metrics["widget_count"] += sum(1 for _ in page.widgets())
            metrics["link_count"] += len(page.get_links())
            metrics["text_char_count"] += len(page.get_text())
        mismatches = [key for key in expected if key in metrics and expected[key] != metrics[key]]
        unsupported = [key for key in expected if key not in metrics and key != "render_pages"]
        has_metrics = any(key in metrics for key in expected)
        checks["structure"] = {"status": "failed" if mismatches else "blocked" if unsupported else "passed" if has_metrics else "pending",
            **metrics, "mismatched_metrics": mismatches, "unsupported_metrics": unsupported}
        if not has_metrics:
            checks["structure"]["reason"] = "expected_metrics_required"
        pages = expected.get("render_pages")
        if not isinstance(pages, list) or not pages:
            checks["render"] = {"status": "pending", "reason": "no_render_pages_requested", "pages": []}
        else:
            renders = []
            for number in pages:
                if type(number) is not int or number < 1 or number > document.page_count:
                    checks["render"] = {"status": "failed", "reason": "invalid_render_page", "pages": renders}
                    break
                pixmap = document[number - 1].get_pixmap(matrix=fitz.Matrix(1, 1), alpha=False)
                pixels = pixmap.samples
                if not pixels or pixmap.width < 1 or pixmap.height < 1:
                    raise ValueError("empty_render")
                renders.append({"page": number, "width": pixmap.width, "height": pixmap.height,
                    "pixel_sha256": hashlib.sha256(pixels).hexdigest()})
            else:
                checks["render"] = {"status": "passed", "dpi": 72, "pages": renders}
    if checks["open"]["repaired"]:
        checks["open"].update(status="pending", reason="engine_repaired_input")
    return checks


def _verify_sample(sample, sample_root, qpdf_enabled):
    result = {"id": sample["id"], "source_type": sample["source"]["type"],
        "path": sample.get("path"), "status": "pending", "checks": {}, "warnings": [], "failures": []}
    checks = result["checks"]
    if not sample.get("path"):
        checks["availability"] = {"status": "pending", "reason": "sample_not_supplied"}
        return result
    try:
        path = _sample_path(sample_root, sample["path"])
    except (ValueError, OSError):
        result.update(status="failed", path=None)
        checks["availability"] = {"status": "failed", "reason": "sample_path_containment_or_link"}
        result["failures"].append("sample_path_containment_or_link")
        return result
    if not path.is_file():
        checks["availability"] = {"status": "pending", "reason": "sample_not_supplied"}
        return result
    before = None
    try:
        before = _sha256(path)
        expected_hash = sample.get("sha256")
        if not isinstance(expected_hash, str) or not re.fullmatch(r"[0-9a-fA-F]{64}", expected_hash):
            checks["sha256"] = {"status": "blocked", "reason": "expected_sha256_required", "actual": before}
        elif before != expected_hash.lower():
            checks["sha256"] = {"status": "failed", "reason": "sha256_mismatch", "actual": before, "expected": expected_hash.lower()}
        else:
            checks["sha256"] = {"status": "passed", "actual": before, "expected": expected_hash.lower()}
            try:
                checks.update(_inspect(path, sample["expected"]))
            except Exception as exc:
                checks["open_structure_render"] = {"status": "failed", "reason": "pdf_check_failed", "error_type": type(exc).__name__}
            if qpdf_enabled:
                try:
                    syntax = check_pdf_syntax(str(path))
                    code = syntax["returncode"]
                    outcome = "clean" if code == 0 else "warning" if code == 3 else "error"
                    checks["qpdf"] = {"status": "passed" if code == 0 else "pending" if code == 3 else "failed",
                        "outcome": outcome, "returncode": code,
                        "diagnostic_sha256": hashlib.sha256(syntax["output"].encode("utf-8")).hexdigest()}
                    if code == 3:
                        result["warnings"].append("qpdf_warnings_require_review")
                except Exception as exc:
                    checks["qpdf"] = {"status": "blocked", "reason": "qpdf_unavailable_or_timed_out", "error_type": type(exc).__name__}
            else:
                checks["qpdf"] = {"status": "pending", "reason": "not_requested"}
    except Exception as exc:
        checks["input"] = {"status": "blocked", "reason": "input_unreadable", "error_type": type(exc).__name__}
    finally:
        if before is not None:
            try:
                after = _sha256(check_unlinked_path(path))
                checks["unchanged"] = {"status": "passed" if before == after else "failed", "before": before, "after": after}
            except Exception as exc:
                checks["unchanged"] = {"status": "failed", "reason": "input_unverifiable_after_checks", "error_type": type(exc).__name__}
    result["status"] = _status(check["status"] for check in checks.values())
    result["failures"].extend(key for key, check in checks.items() if check["status"] == "failed")
    if checks.get("open", {}).get("repaired"):
        result["warnings"].append("pymupdf_repaired_input")
    return result


def run_qualification(manifest_path, output_directory, *, samples_root=None, qpdf_enabled=True):
    """Verify only declared samples. Create a new evidence directory or raise.

    The v1 service never processes PDFs or imports manual evidence. Matrix rows for
    release, real business acceptance and external readers remain pending.
    """
    manifest_path = check_unlinked_path(manifest_path)
    manifest = _load_manifest(manifest_path)
    root_setting = samples_root if samples_root is not None else manifest.get("samples_root")
    if not isinstance(root_setting, (str, Path)) or not str(root_setting):
        raise ValueError("An explicit samples_root is required")
    sample_root = Path(root_setting)
    if not sample_root.is_absolute():
        sample_root = manifest_path.parent / sample_root
    sample_root = check_unlinked_path(sample_root)
    output = check_unlinked_path(output_directory)
    if output.resolve().is_relative_to(sample_root.resolve()):
        raise ValueError("Evidence output must be outside the read-only sample root")
    output.mkdir(parents=True, exist_ok=False)
    engines = {"python": platform.python_version(), "platform": platform.system(),
        "pymupdf": fitz.VersionBind, "mupdf": fitz.VersionFitz, "qpdf": {"status": "pending"}}
    if qpdf_enabled:
        try:
            version = run_qpdf_command(["--version"], timeout=15)
            match = re.search(r"qpdf version ([0-9][0-9A-Za-z.+-]*)", version.stdout)
            engines["qpdf"] = {"status": "passed", "version": match.group(1)} if version.returncode == 0 and match else {"status": "blocked", "reason": "version_unavailable"}
        except Exception as exc:
            engines["qpdf"] = {"status": "blocked", "reason": "engine_unavailable", "error_type": type(exc).__name__}
    results = [_verify_sample(sample, sample_root, qpdf_enabled) for sample in manifest["samples"]]
    if "sample_history" in manifest:
        history = {(record["sample_id"], record["sample_version"]): record
            for record in manifest["sample_history"]["records"]}
        for sample, result in zip(manifest["samples"], results):
            record = history[(sample["id"], sample["sample_version"])]
            approval = record["approval"]["status"]
            result.update(sample_version=sample["sample_version"], expected_metrics_version=sample["expected_metrics_version"])
            result["checks"]["approval"] = {"status": {"pending": "pending", "approved": "passed", "rejected": "failed"}[approval],
                "reason": "independent_sample_change_approval", "approval_status": approval}
            result["status"] = _status([result["status"]] + [check["status"] for check in result["checks"].values()])
            if approval == "rejected":
                result["failures"].append("approval")
    by_id = {result["id"]: result for result in results}
    rows = []
    for row in manifest["matrix"]:
        item = {"id": row["id"], "evidence_level": row["evidence_level"], "sample_ids": row["sample_ids"]}
        for field in ("source_requirement", "feature_requirement", "reader"):
            if field in row:
                item[field] = row[field]
        if row["evidence_level"] == "automated_regression":
            item["status"] = _status(by_id[s]["status"] for s in row["sample_ids"])
        else:
            item.update(status="pending", reason="independent_evidence_required")
        rows.append(item)
    report = {"schema_version": REPORT_SCHEMA_VERSION, "manifest_schema_version": manifest["schema_version"],
        "matrix_version": manifest["matrix_version"], "manifest_sha256": _sha256(manifest_path),
        "sample_history_schema_version": manifest.get("sample_history", {}).get("schema_version"),
        "generated_at": datetime.now(timezone.utc).isoformat(), "engines": engines,
        "status": _status([result["status"] for result in results] + [row["status"] for row in rows] + [engines["qpdf"]["status"]]),
        "summary": {status: sum(result["status"] == status for result in results) for status in STATUSES},
        "results": results, "matrix": rows}
    with (output / "qualification-report.json").open("x", encoding="utf-8") as stream:
        json.dump(report, stream, ensure_ascii=False, indent=2)
        stream.write("\n")
    return report
