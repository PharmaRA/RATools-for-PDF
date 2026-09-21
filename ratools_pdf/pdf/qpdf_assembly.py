"""qpdf 页面装配、范围提取与多卷拆分模块。

提供基于 qpdf 引擎的高性能多文档拼接、页面重排、区间旋转及分卷拆分。
"""

import os
from dataclasses import dataclass
from typing import List, Optional

from ratools_pdf.pdf.qpdf import QpdfResult, run_qpdf_command


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
    """获取 PDF 总页数。优先内存极速读取 (PyMuPDF)，失败时回退至 qpdf。"""
    try:
        import fitz

        doc = fitz.open(input_path)
        cnt = doc.page_count
        doc.close()
        return cnt
    except Exception:
        pass
    try:
        res = run_qpdf_command(["--show-npages", input_path], timeout=10)
        if res.is_success and res.stdout.strip().isdigit():
            return int(res.stdout.strip())
    except Exception:
        pass
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
