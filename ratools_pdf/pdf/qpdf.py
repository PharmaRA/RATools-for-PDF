"""qpdf 命令行工具桥接模块：命令构造、进程执行、重写与底层结构诊断。

支持 qpdf 12.x 全量核心能力（页面几何展平、对象流打包、无损 Flate 重压缩、
权限解除、线性化、合规体检与 JSON AST 探针）。
"""

import json
import os
import re
import shutil
import subprocess
import sys
from dataclasses import dataclass
from typing import Any, Dict, List, Optional

from ratools_pdf.config.paths import get_resource_path


@dataclass
class QpdfResult:
    """qpdf 命令行执行结果封装。"""
    returncode: int
    stdout: str
    stderr: str
    command: List[str]

    @property
    def has_warning(self) -> bool:
        """qpdf 退出码 3 表示成功生成输出但存在警告。"""
        return self.returncode == 3

    @property
    def is_success(self) -> bool:
        """0 为无警告成功，3 为带警告但成功。"""
        return self.returncode in (0, 3)


def get_qpdf_path() -> str:
    """获取可用 qpdf 可执行文件路径。"""
    if sys.platform == "win32":
        candidates = [
            get_resource_path("plugins", "qpdf", "qpdf.exe"),
            os.environ.get("QPDF_PATH", ""),
            shutil.which("qpdf.exe") or "",
            shutil.which("qpdf") or "",
        ]
        for candidate in candidates:
            if candidate and os.path.exists(candidate):
                return os.path.abspath(candidate)
        return "qpdf.exe"

    which_path = shutil.which("qpdf")
    if which_path:
        return which_path
    env_path = os.environ.get("QPDF_PATH", "")
    if env_path and os.path.exists(env_path):
        return env_path
    return "qpdf"


def _make_startupinfo():
    """Windows 下隐藏 CMD 控制台窗口。"""
    if sys.platform == "win32":
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        return startupinfo
    return None


def run_qpdf_command(args: List[str], timeout: Optional[int] = None) -> QpdfResult:
    """执行底层 qpdf 命令并返回 QpdfResult。"""
    qpdf_exe = get_qpdf_path()
    if sys.platform == "win32" and qpdf_exe != "qpdf.exe" and not os.path.exists(qpdf_exe):
        raise FileNotFoundError(f"未找到 qpdf 工具！\n请确保已安装或设置 QPDF_PATH: {qpdf_exe}")

    cmd = [qpdf_exe] + args
    result = subprocess.run(
        cmd,
        startupinfo=_make_startupinfo(),
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=timeout,
    )
    return QpdfResult(
        returncode=result.returncode,
        stdout=result.stdout or "",
        stderr=result.stderr or "",
        command=cmd,
    )


class QpdfCommandBuilder:
    """qpdf 命令行参数强类型构造器。"""

    def __init__(self, input_pdf: Optional[str] = None, output_pdf: Optional[str] = None):
        self.input_pdf = input_pdf
        self.output_pdf = output_pdf
        self.linearize = False
        self.force_version: Optional[str] = None
        self.min_version: Optional[str] = None
        self.decrypt_restrictions = False
        self.flatten_rotation = False
        self.flatten_annotations: Optional[str] = None
        self.object_streams: Optional[str] = None
        self.recompress_flate = False
        self.compression_level: Optional[int] = None
        self.stream_data: Optional[str] = None
        self.extra_args: List[str] = []

    def set_linearize(self, enabled: bool = True) -> "QpdfCommandBuilder":
        self.linearize = enabled
        return self

    def set_force_version(self, version: Optional[str]) -> "QpdfCommandBuilder":
        self.force_version = version
        return self

    def set_min_version(self, version: Optional[str]) -> "QpdfCommandBuilder":
        self.min_version = version
        return self

    def set_decrypt_restrictions(self, enabled: bool = True) -> "QpdfCommandBuilder":
        self.decrypt_restrictions = enabled
        return self

    def set_flatten_rotation(self, enabled: bool = True) -> "QpdfCommandBuilder":
        self.flatten_rotation = enabled
        return self

    def set_flatten_annotations(self, mode: Optional[str] = "all") -> "QpdfCommandBuilder":
        """mode 可为 'all', 'print', 'screen' 或 None。"""
        self.flatten_annotations = mode
        return self

    def set_object_streams(self, mode: Optional[str] = "generate") -> "QpdfCommandBuilder":
        """mode 可为 'generate', 'preserve', 'disable' 或 None。"""
        self.object_streams = mode
        return self

    def set_recompress_flate(self, enabled: bool = True, compression_level: int = 9) -> "QpdfCommandBuilder":
        self.recompress_flate = enabled
        self.compression_level = compression_level
        return self

    def set_stream_data(self, mode: Optional[str] = "compress") -> "QpdfCommandBuilder":
        """mode 可为 'compress', 'uncompress', 'preserve' 或 None。"""
        self.stream_data = mode
        return self

    def add_arg(self, arg: str) -> "QpdfCommandBuilder":
        self.extra_args.append(arg)
        return self

    def build_args(self) -> List[str]:
        """构建参数列表（不含可执行文件路径）。"""
        args: List[str] = []

        if self.decrypt_restrictions:
            args.append("--decrypt")

        if self.flatten_rotation:
            args.append("--flatten-rotation")

        if self.flatten_annotations:
            args.append(f"--flatten-annotations={self.flatten_annotations}")

        if self.object_streams:
            args.append(f"--object-streams={self.object_streams}")

        if self.recompress_flate:
            args.append("--recompress-flate")
            if self.compression_level is not None:
                args.append(f"--compression-level={self.compression_level}")

        if self.stream_data:
            args.append(f"--stream-data={self.stream_data}")

        if self.linearize:
            args.append("--linearize")

        if self.force_version:
            args.append(f"--force-version={self.force_version}")

        if self.min_version:
            args.append(f"--min-version={self.min_version}")

        args.extend(self.extra_args)

        if self.input_pdf:
            args.append(self.input_pdf)
        if self.output_pdf:
            args.append(self.output_pdf)

        return args


def rewrite_with_qpdf(
    input_pdf: str,
    output_pdf: str,
    force_version: Optional[str] = None,
    linearize: bool = False,
    decrypt_restrictions: bool = False,
    flatten_rotation: bool = False,
    flatten_annotations: Optional[str] = None,
    object_streams: Optional[str] = None,
    recompress_flate: bool = False,
    stream_data: Optional[str] = None,
    compression_level: int = 9,
    timeout: Optional[int] = 120,
) -> QpdfResult:
    """调用 qpdf 重写 PDF 文件。

    qpdf 退出码约定：
    - 0: 成功
    - 3: 有警告但已成功生成有效输出文件
    - 2: 致命错误
    """
    builder = QpdfCommandBuilder(input_pdf=input_pdf, output_pdf=output_pdf)
    if force_version:
        builder.set_force_version(force_version)
    if linearize:
        builder.set_linearize(True)
    if decrypt_restrictions:
        builder.set_decrypt_restrictions(True)
    if flatten_rotation:
        builder.set_flatten_rotation(True)
    if flatten_annotations:
        builder.set_flatten_annotations(flatten_annotations)
    if object_streams:
        builder.set_object_streams(object_streams)
    if recompress_flate:
        builder.set_recompress_flate(True, compression_level=compression_level)
    if stream_data:
        builder.set_stream_data(stream_data)

    cmd_args = builder.build_args()
    res = run_qpdf_command(cmd_args, timeout=timeout)

    # 退出码 3 且输出文件存在，视为成功
    if res.returncode == 3 and os.path.exists(output_pdf):
        return res

    if res.returncode != 0:
        detail = (res.stderr or "").strip() or (res.stdout or "").strip()
        if not detail:
            detail = f"qpdf 返回码 {res.returncode}，未提供详细信息"
        raise RuntimeError(f"qpdf 执行失败: {detail}")

    return res


def format_qpdf_error(error: Any) -> str:
    """格式化 qpdf 异常信息为用户友好文案。"""
    text = str(error).strip()
    lowered = text.lower()
    password_markers = ["password", "invalid password", "incorrect password", "requires a password"]
    if any(marker in lowered for marker in password_markers):
        return "该PDF需要密码，当前模式不支持输入密码解锁"
    if text.startswith("qpdf 执行失败:"):
        detail = text.split(":", 1)[1].strip()
        return f"未能移除PDF权限限制：{detail}" if detail else "未能移除PDF权限限制"
    return text


def read_pdf_header_version(input_path: str) -> str:
    """读取 PDF 头部真实声明版本（如 '1.7'）。"""
    try:
        with open(input_path, "rb") as f:
            header = f.read(32)
    except Exception:
        return ""
    match = re.search(rb"%PDF-(\d+\.\d+)", header)
    return match.group(1).decode("ascii") if match else ""


def is_pdf_linearized(input_path: str) -> bool:
    """检查 PDF 是否处于线性化状态。"""
    try:
        with open(input_path, "rb") as f:
            header = f.read(4096)
    except Exception:
        return False
    return b"/Linearized" in header


def qpdf_encryption_info(input_path: str) -> str:
    """通过 qpdf --show-encryption 查询 PDF 加密与权限详细信息。"""
    try:
        res = run_qpdf_command(["--show-encryption", input_path], timeout=15)
        return f"{res.stdout}\n{res.stderr}".strip()
    except Exception:
        return ""


def qpdf_reports_restrictions(input_path: str) -> bool:
    """检查 PDF 是否存在被 qpdf 报告的限制项（': not allowed'）。"""
    info = qpdf_encryption_info(input_path).lower()
    if not info or "file is not encrypted" in info:
        return False
    return ": not allowed" in info


def check_pdf_syntax(input_path: str, timeout: int = 30) -> Dict[str, Any]:
    """通过 qpdf --check 校验 PDF 语法结构与完整性。"""
    res = run_qpdf_command(["--check", input_path], timeout=timeout)
    return {
        "is_healthy": res.returncode == 0,
        "has_warnings": res.returncode == 3,
        "returncode": res.returncode,
        "output": f"{res.stdout}\n{res.stderr}".strip(),
    }


def inspect_pdf_json(input_path: str, timeout: int = 60) -> Optional[Dict[str, Any]]:
    """通过 qpdf --json 获取 PDF 的底层语法树与元数据 AST。"""
    res = run_qpdf_command(["--json", input_path], timeout=timeout)
    if res.is_success and res.stdout:
        try:
            return json.loads(res.stdout)
        except Exception:
            return None
    return None


@dataclass
class PageSpec:
    """页面装配单元定义。"""
    file_path: str
    page_range: str = "1-z"
    password: Optional[str] = None
    rotation: Optional[int] = None  # 0, 90, 180, 270


def parse_page_range(range_str: str, total_pages: int) -> List[int]:
    """解析 qpdf 风格页面范围字符串并返回 1-based 页码列表。

    支持单页、连续区间（1-5）、倒序（z-1）、奇偶修饰（:odd, :even）、离散列表（1,3,5-7）。
    """
    range_str = range_str.strip().lower()
    if not range_str or range_str in ("all", "1-z", "*"):
        return list(range(1, total_pages + 1))

    modifier = None
    if ":odd" in range_str:
        modifier = "odd"
        range_str = range_str.replace(":odd", "")
    elif ":even" in range_str:
        modifier = "even"
        range_str = range_str.replace(":even", "")

    def resolve_token(t: str) -> int:
        t = t.strip()
        if t == "z":
            return total_pages
        if t.startswith("r") and t[1:].isdigit():
            offset = int(t[1:])
            return max(1, total_pages - offset + 1)
        return int(t)

    pages: List[int] = []
    groups = [g.strip() for g in range_str.split(",") if g.strip()]
    for g in groups:
        try:
            if "-" in g:
                parts = g.split("-", 1)
                start = resolve_token(parts[0])
                end = resolve_token(parts[1])
                step = 1 if start <= end else -1
                pages.extend(range(start, end + step, step))
            else:
                pages.append(resolve_token(g))
        except Exception:
            continue

    pages = [p for p in pages if 1 <= p <= total_pages]
    if modifier == "odd":
        pages = pages[0::2]
    elif modifier == "even":
        pages = pages[1::2]
    return pages


def get_pdf_page_count(input_path: str) -> int:
    """获取 PDF 总页数。优先调用 qpdf，失败时回退至 pymupdf。"""
    try:
        res = run_qpdf_command(["--show-npages", input_path], timeout=10)
        if res.is_success and res.stdout.strip().isdigit():
            return int(res.stdout.strip())
    except Exception:
        pass
    try:
        import fitz
        doc = fitz.open(input_path)
        cnt = doc.page_count
        doc.close()
        return cnt
    except Exception:
        return 0


def assemble_pages(
    specs: List[PageSpec],
    output_pdf: str,
    linearize: bool = False,
    object_streams: Optional[str] = "generate",
    timeout: Optional[int] = 180,
) -> QpdfResult:
    """多文件/单文件页面装配、重排与提取。

    利用 qpdf --empty --pages 构建目标文档，并应用各区间指定的旋转角。
    """
    if not specs:
        raise ValueError("装配规格列表不能为空")

    cmd_args = ["--empty", "--pages"]
    rotations: List[tuple[int, int, int]] = []
    current_page = 1

    for spec in specs:
        total_p = get_pdf_page_count(spec.file_path)
        selected_pages = parse_page_range(spec.page_range, total_p)
        count = len(selected_pages)
        if count == 0:
            continue

        cmd_args.append(spec.file_path)
        if spec.password:
            cmd_args.append(f"--password={spec.password}")
        if spec.page_range and spec.page_range.strip() not in ("1-z", "all", "*"):
            cmd_args.append(spec.page_range.strip())

        if spec.rotation and spec.rotation % 360 != 0:
            rotations.append((spec.rotation % 360, current_page, current_page + count - 1))

        current_page += count

    cmd_args.append("--")

    for angle, start_p, end_p in rotations:
        range_str = f"{start_p}" if start_p == end_p else f"{start_p}-{end_p}"
        cmd_args.append(f"--rotate=+{angle}:{range_str}")

    if object_streams:
        cmd_args.append(f"--object-streams={object_streams}")
    if linearize:
        cmd_args.append("--linearize")

    cmd_args.append(output_pdf)

    res = run_qpdf_command(cmd_args, timeout=timeout)
    if res.returncode == 3 and os.path.exists(output_pdf):
        return res
    if res.returncode != 0:
        detail = (res.stderr or "").strip() or (res.stdout or "").strip()
        raise RuntimeError(f"页面装配失败: {detail}")
    return res


def split_pages(
    input_pdf: str,
    output_pattern: str,
    pages_per_file: Optional[int] = None,
    password: Optional[str] = None,
    timeout: Optional[int] = 180,
) -> QpdfResult:
    """拆分 PDF 页面为独立文件或多卷文件。

    output_pattern 支持含 '%d' 格式符（如 'out_%d.pdf'），或常规文件名（自动追加序号）。
    """
    cmd_args = [input_pdf]
    if password:
        cmd_args.append(f"--password={password}")

    if pages_per_file and pages_per_file > 1:
        cmd_args.append(f"--split-pages={pages_per_file}")
    else:
        cmd_args.append("--split-pages")

    cmd_args.append(output_pattern)

    res = run_qpdf_command(cmd_args, timeout=timeout)
    if res.returncode == 3:
        return res
    if res.returncode != 0:
        detail = (res.stderr or "").strip() or (res.stdout or "").strip()
        raise RuntimeError(f"页面拆分失败: {detail}")
    return res


# ================= 兼容性别名（保留原下划线函数引用） =================
_get_qpdf_path = get_qpdf_path
_rewrite_with_qpdf = rewrite_with_qpdf
_format_qpdf_error = format_qpdf_error
_read_pdf_header_version = read_pdf_header_version
_is_pdf_linearized = is_pdf_linearized
_qpdf_encryption_info = qpdf_encryption_info
_qpdf_reports_restrictions = qpdf_reports_restrictions
