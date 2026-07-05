import csv
import json
import os
import re
import shutil
import subprocess
import sys
from pathlib import Path
from urllib.parse import unquote, urlparse

import fitz

from ratools_pdf.config.paths import get_resource_path
from ratools_pdf.pdf.font_embedding_providers import get_font_embedding_provider


def _processor_cls():
    from ratools_pdf.pdf.processor import PDFProcessor
    return PDFProcessor


def _get_qpdf_path():
    if sys.platform == "win32":
        candidates = [
            get_resource_path("plugins", "qpdf", "qpdf.exe"),
            os.environ.get("QPDF_PATH", ""),
            r"D:\Program Files\qpdf 11.9.1\bin\qpdf.exe",
            r"C:\Program Files\qpdf\bin\qpdf.exe",
        ]
        for candidate in candidates:
            if candidate and os.path.exists(candidate):
                return candidate
        return "qpdf.exe"
    return "qpdf"


def _rewrite_with_qpdf(input_pdf, output_pdf, force_version=None, linearize=False, decrypt_restrictions=False):
    qpdf_exe = _processor_cls()._get_qpdf_path()
    if sys.platform == "win32" and qpdf_exe != "qpdf.exe" and not os.path.exists(qpdf_exe):
        raise FileNotFoundError(f"未找到 qpdf 工具！\n请确保已安装或设置 QPDF_PATH: {qpdf_exe}")

    cmd = [qpdf_exe]
    if decrypt_restrictions:
        cmd.append("--decrypt")
    if linearize:
        cmd.append("--linearize")
    if force_version:
        cmd.append(f"--force-version={force_version}")
    cmd.extend([input_pdf, output_pdf])

    startupinfo = None
    if sys.platform == "win32":
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW

    result = subprocess.run(cmd, startupinfo=startupinfo, capture_output=True, text=True)
    # qpdf 退出码约定：0=成功，3=有警告但已成功生成输出，2=有错误。
    # 结构不够规范的 PDF 在线性化/版本转换时常返回 3，输出文件实际可用，不应判为失败。
    if result.returncode == 3 and os.path.exists(output_pdf):
        return

    if result.returncode != 0:
        detail = (result.stderr or "").strip() or (result.stdout or "").strip()
        if not detail:
            detail = f"qpdf 返回码 {result.returncode}，未提供详细信息"
        raise RuntimeError(f"qpdf 执行失败: {detail}")


def _format_qpdf_error(error):
    text = str(error).strip()
    lowered = text.lower()
    password_markers = ["password", "invalid password", "incorrect password", "requires a password"]
    if any(marker in lowered for marker in password_markers):
        return "该PDF需要密码，当前模式不支持输入密码解锁"
    if text.startswith("qpdf 执行失败:"):
        detail = text.split(":", 1)[1].strip()
        return f"未能移除PDF权限限制：{detail}" if detail else "未能移除PDF权限限制"
    return text

def _read_pdf_header_version(input_path):
    try:
        with open(input_path, "rb") as f:
            header = f.read(32)
    except Exception:
        return ""
    match = re.search(rb"%PDF-(\d+\.\d+)", header)
    return match.group(1).decode("ascii") if match else ""


def _is_pdf_linearized(input_path):
    try:
        with open(input_path, "rb") as f:
            header = f.read(4096)
    except Exception:
        return False
    return b"/Linearized" in header


def _qpdf_encryption_info(input_path):
    qpdf_exe = _processor_cls()._get_qpdf_path()
    if sys.platform == "win32" and qpdf_exe != "qpdf.exe" and not os.path.exists(qpdf_exe):
        return ""

    startupinfo = None
    if sys.platform == "win32":
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW

    try:
        result = subprocess.run(
            [qpdf_exe, "--show-encryption", input_path],
            startupinfo=startupinfo,
            capture_output=True,
            text=True,
            timeout=15,
        )
    except Exception:
        return ""
    return f"{result.stdout}\n{result.stderr}".strip()


def _qpdf_reports_restrictions(input_path):
    info = _processor_cls()._qpdf_encryption_info(input_path).lower()
    if not info or "file is not encrypted" in info:
        return False
    return ": not allowed" in info


def rewrite_with_qpdf(input_pdf, output_pdf, force_version=None, linearize=False, decrypt_restrictions=False):
    return _rewrite_with_qpdf(input_pdf, output_pdf, force_version, linearize, decrypt_restrictions)

def get_qpdf_path():
    return _get_qpdf_path()

def format_qpdf_error(error):
    return _format_qpdf_error(error)

