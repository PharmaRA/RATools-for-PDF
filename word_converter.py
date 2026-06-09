import os
import platform
import re
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from pathlib import Path


WORD_EXTENSIONS = {".doc", ".docx"}
WORD_ECTD_OPTIONS = {
    "convert_pdf_version",
    "fast_web_view",
    "initial_view_bookmarks_and_page",
    "page_layout_default",
    "open_page_first",
    "bookmark_inherit_zoom",
    "bookmark_open_new_window",
    "collapse_all_bookmarks",
    "link_inherit_zoom",
    "link_open_new_window",
    "cleanup_remove_annotations",
    "cleanup_remove_metadata",
    "cleanup_remove_attachments",
    "cleanup_remove_dynamic_content",
    "filename_ectd_format",
}

WD_EXPORT_FORMAT_PDF = 17
WD_EXPORT_OPTIMIZE_FOR_PRINT = 0
WD_EXPORT_ALL_DOCUMENT = 0
WD_EXPORT_DOCUMENT_CONTENT = 0
WD_EXPORT_CREATE_HEADING_BOOKMARKS = 1

XML_NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
}


class WordConversionError(Exception):
    code = "conversion_error"
    message = "Word 转换失败"

    def __init__(self, message=None, code=None):
        self.code = code or self.code
        self.message = message or self.message
        super().__init__(self.message)


class NonWindowsWordConversionError(WordConversionError):
    code = "non_windows"
    message = "当前系统不是 Windows，无法调用 Microsoft Word COM 进行转换"


class WordNotInstalledError(WordConversionError):
    code = "word_not_installed"
    message = "未检测到 Microsoft Word，请确认已安装桌面版 Word"


class WordFileNotFoundError(WordConversionError):
    code = "file_not_found"
    message = "Word 文件不存在"


class UnsupportedWordFileError(WordConversionError):
    code = "unsupported_suffix"
    message = "仅支持 .doc 和 .docx 文件"


class WordFileOccupiedError(WordConversionError):
    code = "file_occupied"
    message = "Word 文件可能正被其它程序占用，请关闭后重试"


class WordComConversionError(WordConversionError):
    code = "com_conversion_failed"
    message = "Word COM 转换失败"


class OutputPdfMissingError(WordConversionError):
    code = "output_pdf_missing"
    message = "转换完成后未生成输出 PDF"


@dataclass
class WordPrecheckFinding:
    code: str
    message: str
    severity: str = "warning"


@dataclass
class WordPrecheckReport:
    file_path: str
    available: bool = True
    needs_windows_word_check: bool = False
    findings: list = field(default_factory=list)
    error: str = ""

    def add(self, code, message, severity="warning"):
        self.findings.append(WordPrecheckFinding(code=code, message=message, severity=severity))

    @property
    def has_warnings(self):
        return bool(self.findings or self.error or self.needs_windows_word_check)

    def to_text(self):
        name = os.path.basename(self.file_path)
        lines = [f"Word 结构预检查报告：{name}"]
        if self.error:
            lines.append(f"状态：预检失败 - {self.error}")
        elif not self.findings and not self.needs_windows_word_check:
            lines.append("状态：未发现明显结构风险")
        else:
            lines.append("状态：需要复核")
        if self.needs_windows_word_check:
            lines.append("- .doc 文件需要 Windows/Word 深度检查标题、域和页面设置")
        for finding in self.findings:
            lines.append(f"- {finding.message}")
        return "\n".join(lines)

    def to_dict(self):
        return {
            "file_path": self.file_path,
            "available": self.available,
            "needs_windows_word_check": self.needs_windows_word_check,
            "error": self.error,
            "findings": [
                {"code": item.code, "message": item.message, "severity": item.severity}
                for item in self.findings
            ],
            "text": self.to_text(),
        }


def get_word_ectd_options():
    return set(WORD_ECTD_OPTIONS)


def build_word_export_config(pdfa=False):
    return {
        "ExportFormat": WD_EXPORT_FORMAT_PDF,
        "OpenAfterExport": False,
        "OptimizeFor": WD_EXPORT_OPTIMIZE_FOR_PRINT,
        "Range": WD_EXPORT_ALL_DOCUMENT,
        "From": 1,
        "To": 1,
        "Item": WD_EXPORT_DOCUMENT_CONTENT,
        "IncludeDocProps": True,
        "KeepIRM": True,
        "CreateBookmarks": WD_EXPORT_CREATE_HEADING_BOOKMARKS,
        "DocStructureTags": True,
        "BitmapMissingFonts": True,
        "UseISO19005_1": bool(pdfa),
    }


def validate_word_input(file_path, system_name=None):
    path = Path(file_path)
    if not path.exists():
        raise WordFileNotFoundError()
    if not path.is_file():
        raise WordFileNotFoundError("Word 路径不是文件")
    if path.suffix.lower() not in WORD_EXTENSIONS:
        raise UnsupportedWordFileError()
    if (system_name or platform.system()) != "Windows":
        raise NonWindowsWordConversionError()
    return str(path)


def collect_word_files(paths):
    files = []
    seen = set()
    for raw_path in paths or []:
        path = Path(raw_path)
        candidates = []
        if path.is_file():
            candidates = [path]
        elif path.is_dir():
            candidates = [
                item
                for item in path.rglob("*")
                if item.is_file() and item.suffix.lower() in WORD_EXTENSIONS and not item.name.startswith("~$")
            ]
        for candidate in candidates:
            if candidate.suffix.lower() not in WORD_EXTENSIONS or candidate.name.startswith("~$"):
                continue
            normalized = os.path.normpath(str(candidate))
            if normalized not in seen:
                seen.add(normalized)
                files.append(normalized)
    return files


def ectd_pdf_filename_for_word(word_path):
    stem = Path(word_path).stem.lower().replace(" ", "-")
    stem = re.sub(r"[^a-z0-9_-]", "", stem)
    stem = stem.strip("-_")
    if not stem:
        stem = "word-document"
    return f"{stem}.pdf"


def build_word_output_path(word_path, output_dir, common_base=""):
    word_path = os.path.abspath(word_path)
    rel_dir = ""
    if common_base:
        try:
            rel_dir = os.path.relpath(os.path.dirname(word_path), common_base)
        except ValueError:
            rel_dir = ""
        if rel_dir == ".":
            rel_dir = ""
    target_dir = os.path.join(output_dir, rel_dir) if rel_dir else output_dir
    os.makedirs(target_dir, exist_ok=True)
    return os.path.join(target_dir, ectd_pdf_filename_for_word(word_path))


def _read_docx_xml(zip_file, name):
    try:
        return zip_file.read(name)
    except KeyError:
        return b""


def _qname(local_name):
    return f"{{{XML_NS['w']}}}{local_name}"


def _attr(element, local_name):
    return element.attrib.get(_qname(local_name), "")


def _parse_xml(data):
    if not data:
        return None
    try:
        return ET.fromstring(data)
    except ET.ParseError:
        return None


def _collect_heading_style_ids(styles_root):
    heading_ids = set()
    if styles_root is None:
        return heading_ids
    for style in styles_root.findall(".//w:style", XML_NS):
        style_id = _attr(style, "styleId")
        name_node = style.find("w:name", XML_NS)
        style_name = _attr(name_node, "val") if name_node is not None else ""
        style_type = _attr(style, "type")
        if style_type != "paragraph":
            continue
        probe = f"{style_id} {style_name}".lower()
        if re.search(r"\bheading\s*[1-9]\b", probe) or "标题" in probe:
            heading_ids.add(style_id)
    heading_ids.update({f"Heading{i}" for i in range(1, 10)})
    return heading_ids


def _paragraph_text(paragraph):
    return "".join(node.text or "" for node in paragraph.findall(".//w:t", XML_NS)).strip()


def _paragraph_style_id(paragraph):
    node = paragraph.find("./w:pPr/w:pStyle", XML_NS)
    return _attr(node, "val") if node is not None else ""


def _paragraph_has_large_bold_text(paragraph):
    sizes = []
    has_bold = False
    for run in paragraph.findall("./w:r", XML_NS):
        rpr = run.find("./w:rPr", XML_NS)
        if rpr is None:
            continue
        if rpr.find("./w:b", XML_NS) is not None:
            has_bold = True
        sz_node = rpr.find("./w:sz", XML_NS)
        try:
            if sz_node is not None:
                sizes.append(int(_attr(sz_node, "val")))
        except ValueError:
            pass
    return has_bold and sizes and max(sizes) >= 28


def _paragraph_has_numbering(paragraph):
    return paragraph.find("./w:pPr/w:numPr", XML_NS) is not None


def _looks_like_manual_heading(text):
    if not text or len(text) > 90:
        return False
    return bool(re.match(r"^(\d+(\.\d+){0,5}|[A-Z][A-Z0-9 .:-]{4,}|第[一二三四五六七八九十0-9]+[章节])", text))


def _document_field_text(document_root):
    if document_root is None:
        return ""
    parts = []
    for node in document_root.findall(".//w:instrText", XML_NS):
        if node.text:
            parts.append(node.text)
    for node in document_root.findall(".//w:fldSimple", XML_NS):
        value = _attr(node, "instr")
        if value:
            parts.append(value)
    return " ".join(parts)


def _has_dirty_fields(document_root):
    if document_root is None:
        return False
    for node in document_root.findall(".//w:fldChar", XML_NS):
        if _attr(node, "dirty").lower() in {"true", "1", "on"}:
            return True
    for node in document_root.findall(".//w:fldSimple", XML_NS):
        if _attr(node, "dirty").lower() in {"true", "1", "on"}:
            return True
    return False


def _check_sections(document_root, report):
    if document_root is None:
        return
    sect_prs = document_root.findall(".//w:sectPr", XML_NS)
    for idx, sect_pr in enumerate(sect_prs, start=1):
        size = sect_pr.find("./w:pgSz", XML_NS)
        margins = sect_pr.find("./w:pgMar", XML_NS)
        if size is not None:
            try:
                width = int(_attr(size, "w"))
                height = int(_attr(size, "h"))
                normalized = sorted([width, height])
                a4 = sorted([11906, 16838])
                letter = sorted([12240, 15840])
                if normalized not in [a4, letter]:
                    report.add("page_size_hint", f"第 {idx} 节纸张尺寸不是常见 A4/Letter，需确认 eCTD 版式")
            except ValueError:
                pass
        if margins is not None:
            for side in ["top", "bottom", "left", "right"]:
                try:
                    value = int(_attr(margins, side))
                except ValueError:
                    continue
                if value < 720:
                    report.add("margin_hint", f"第 {idx} 节 {side} 页边距小于 0.5 英寸，需复核")
                    break


def precheck_word_structure(file_path):
    report = WordPrecheckReport(file_path=str(file_path))
    path = Path(file_path)
    if not path.exists():
        report.available = False
        report.error = WordFileNotFoundError.message
        return report
    if path.suffix.lower() not in WORD_EXTENSIONS:
        report.available = False
        report.error = UnsupportedWordFileError.message
        return report
    if path.suffix.lower() == ".doc":
        report.needs_windows_word_check = True
        return report

    try:
        with zipfile.ZipFile(path) as zf:
            document_root = _parse_xml(_read_docx_xml(zf, "word/document.xml"))
            styles_root = _parse_xml(_read_docx_xml(zf, "word/styles.xml"))
    except zipfile.BadZipFile:
        report.available = False
        report.error = "docx 文件结构损坏，无法读取"
        return report

    if document_root is None:
        report.available = False
        report.error = "无法读取 docx 主文档 XML"
        return report

    heading_style_ids = _collect_heading_style_ids(styles_root)
    paragraphs = document_root.findall(".//w:p", XML_NS)
    heading_paragraphs = []
    manual_heading_candidates = []
    for paragraph in paragraphs:
        text = _paragraph_text(paragraph)
        if not text:
            continue
        style_id = _paragraph_style_id(paragraph)
        if style_id in heading_style_ids:
            heading_paragraphs.append(text)
            continue
        if _looks_like_manual_heading(text) and (_paragraph_has_numbering(paragraph) or _paragraph_has_large_bold_text(paragraph)):
            manual_heading_candidates.append(text)

    field_text = _document_field_text(document_root)
    has_toc = bool(re.search(r"\bTOC\b", field_text, re.IGNORECASE))
    if has_toc and not heading_paragraphs:
        report.add("toc_without_heading_styles", "文档存在目录域，但未发现使用 Word 标题样式的段落")
    if not heading_paragraphs:
        report.add("missing_heading_styles", "未发现 Word 标题样式，转换后可能无法生成 PDF 书签")
    if manual_heading_candidates:
        sample = "、".join(manual_heading_candidates[:3])
        report.add("manual_heading_candidates", f"发现可能用加粗/编号伪装的手动标题：{sample}")
    if _has_dirty_fields(document_root) or re.search(r"\b(TOC|REF|PAGEREF)\b", field_text, re.IGNORECASE):
        report.add("fields_need_update", "文档包含 TOC/REF/PAGEREF 等域，转换前将尝试自动更新字段")

    _check_sections(document_root, report)
    return report


def ensure_windows_word_available(system_name=None):
    if (system_name or platform.system()) != "Windows":
        raise NonWindowsWordConversionError()


def _check_file_not_occupied(file_path):
    if platform.system() != "Windows":
        return
    try:
        with open(file_path, "a+b"):
            pass
    except OSError:
        raise WordFileOccupiedError()


def convert_word_to_pdf(input_path, output_pdf_path, pdfa=False, visible=False):
    validate_word_input(input_path)
    _check_file_not_occupied(input_path)
    output_pdf_path = os.path.abspath(output_pdf_path)
    os.makedirs(os.path.dirname(output_pdf_path), exist_ok=True)

    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        raise WordNotInstalledError("缺少 pywin32 或无法加载 Word COM 支持") from exc

    word = None
    document = None
    try:
        pythoncom.CoInitialize()
        try:
            word = win32com.client.DispatchEx("Word.Application")
        except Exception as exc:
            raise WordNotInstalledError() from exc
        word.Visible = bool(visible)
        word.DisplayAlerts = 0
        try:
            document = word.Documents.Open(os.path.abspath(input_path), ReadOnly=True, AddToRecentFiles=False)
        except Exception as exc:
            raise WordFileOccupiedError() from exc

        try:
            document.Fields.Update()
            for toc in document.TablesOfContents:
                toc.Update()
        except Exception:
            pass

        config = build_word_export_config(pdfa=pdfa)
        try:
            document.ExportAsFixedFormat(
                OutputFileName=output_pdf_path,
                ExportFormat=config["ExportFormat"],
                OpenAfterExport=config["OpenAfterExport"],
                OptimizeFor=config["OptimizeFor"],
                Range=config["Range"],
                From=config["From"],
                To=config["To"],
                Item=config["Item"],
                IncludeDocProps=config["IncludeDocProps"],
                KeepIRM=config["KeepIRM"],
                CreateBookmarks=config["CreateBookmarks"],
                DocStructureTags=config["DocStructureTags"],
                BitmapMissingFonts=config["BitmapMissingFonts"],
                UseISO19005_1=config["UseISO19005_1"],
            )
        except Exception as exc:
            raise WordComConversionError(f"Word COM 转换失败：{exc}") from exc
    finally:
        if document is not None:
            try:
                document.Close(False)
            except Exception:
                pass
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

    if not os.path.exists(output_pdf_path) or os.path.getsize(output_pdf_path) <= 0:
        raise OutputPdfMissingError()
    return output_pdf_path
