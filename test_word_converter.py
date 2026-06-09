import zipfile

import fitz
import pytest

from pdf_processor import PDFProcessor
from word_converter import (
    NonWindowsWordConversionError,
    UnsupportedWordFileError,
    WORD_ECTD_OPTIONS,
    build_word_export_config,
    build_word_output_path,
    collect_word_files,
    convert_word_to_pdf,
    get_word_ectd_options,
    precheck_word_structure,
    validate_word_input,
)


def _write_docx(path, document_xml, styles_xml=None):
    if styles_xml is None:
        styles_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
  </w:style>
</w:styles>"""
    with zipfile.ZipFile(path, "w") as zf:
        zf.writestr("[Content_Types].xml", "<Types/>")
        zf.writestr("word/document.xml", document_xml)
        zf.writestr("word/styles.xml", styles_xml)


def test_validate_word_input_reports_non_windows_for_existing_docx(tmp_path):
    docx = tmp_path / "source.docx"
    _write_docx(docx, '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>')

    with pytest.raises(NonWindowsWordConversionError) as exc:
        validate_word_input(docx, system_name="Linux")

    assert exc.value.code == "non_windows"
    assert "Windows" in exc.value.message


def test_validate_word_input_rejects_unsupported_suffix_before_platform(tmp_path):
    txt = tmp_path / "source.txt"
    txt.write_text("not word", encoding="utf-8")

    with pytest.raises(UnsupportedWordFileError) as exc:
        validate_word_input(txt, system_name="Linux")

    assert exc.value.code == "unsupported_suffix"


def test_convert_word_to_pdf_non_windows_error_branch(tmp_path):
    docx = tmp_path / "source.docx"
    _write_docx(docx, '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>')

    with pytest.raises(NonWindowsWordConversionError):
        convert_word_to_pdf(str(docx), str(tmp_path / "out.pdf"))


def test_collect_word_files_recurses_and_filters(tmp_path):
    nested = tmp_path / "nested"
    nested.mkdir()
    docx = nested / "a.docx"
    doc = tmp_path / "b.doc"
    temp = tmp_path / "~$locked.docx"
    other = tmp_path / "c.pdf"
    docx.write_bytes(b"x")
    doc.write_bytes(b"x")
    temp.write_bytes(b"x")
    other.write_bytes(b"x")

    files = collect_word_files([tmp_path])

    assert str(docx) in files
    assert str(doc) in files
    assert str(temp) not in files
    assert str(other) not in files


def test_docx_static_precheck_finds_toc_without_heading_and_manual_heading(tmp_path):
    docx = tmp_path / "manual-heading.docx"
    document_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:instrText> TOC \\o "1-3" \\h </w:instrText></w:r></w:p>
    <w:p>
      <w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>
      <w:r><w:rPr><w:b/><w:sz w:val="32"/></w:rPr><w:t>1. Introduction</w:t></w:r>
    </w:p>
    <w:sectPr>
      <w:pgSz w:w="10000" w:h="14000"/>
      <w:pgMar w:top="600" w:right="1440" w:bottom="1440" w:left="1440"/>
    </w:sectPr>
  </w:body>
</w:document>"""
    styles_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
</w:styles>"""
    _write_docx(docx, document_xml, styles_xml)

    report = precheck_word_structure(docx)
    codes = {finding.code for finding in report.findings}

    assert report.available is True
    assert "toc_without_heading_styles" in codes
    assert "missing_heading_styles" in codes
    assert "manual_heading_candidates" in codes
    assert "fields_need_update" in codes
    assert "page_size_hint" in codes
    assert "margin_hint" in codes


def test_docx_static_precheck_accepts_real_heading_style(tmp_path):
    docx = tmp_path / "heading.docx"
    document_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Introduction</w:t></w:r></w:p>
    <w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>
  </w:body>
</w:document>"""
    _write_docx(docx, document_xml)

    report = precheck_word_structure(docx)

    assert report.available is True
    assert report.findings == []


def test_word_ectd_option_set_matches_required_defaults():
    options = get_word_ectd_options()
    required = {
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

    assert options == required
    assert "cleanup_remove_all_links_bookmarks" not in options
    assert WORD_ECTD_OPTIONS == required


def test_pdfa_export_config_maps_to_word_fixed_format_arguments():
    normal = build_word_export_config(pdfa=False)
    pdfa = build_word_export_config(pdfa=True)

    assert normal["ExportFormat"] == 17
    assert normal["CreateBookmarks"] == 1
    assert normal["UseISO19005_1"] is False
    assert pdfa["UseISO19005_1"] is True


def test_build_word_output_path_uses_ectd_filename_and_common_base(tmp_path):
    source_dir = tmp_path / "src" / "m1"
    source_dir.mkdir(parents=True)
    word = source_dir / "My Study 文档.docx"
    word.write_bytes(b"x")
    output_root = tmp_path / "out"

    output = build_word_output_path(str(word), str(output_root), common_base=str(tmp_path / "src"))

    assert output.endswith("m1/my-study.pdf")
    assert (output_root / "m1").is_dir()


def _write_pdf(path, title=None, author=None):
    doc = fitz.open()
    page = doc.new_page(width=300, height=300)
    page.insert_text((72, 72), "RATools test PDF")
    metadata = {}
    if title is not None:
        metadata["title"] = title
    if author is not None:
        metadata["author"] = author
    if metadata:
        doc.set_metadata(metadata)
    doc.save(path)
    doc.close()


def test_filtered_pdf_precheck_only_reports_selected_options(tmp_path):
    pdf = tmp_path / "filtered.pdf"
    _write_pdf(pdf, title="Wrong title", author="Author")

    broad_report = PDFProcessor.build_precheck_report(str(pdf))
    filtered_report = PDFProcessor.build_precheck_report(str(pdf), selected_options={"title_from_filename"})

    assert "title_from_filename" in broad_report["suggestions"]
    assert "cleanup_remove_metadata" in broad_report["suggestions"]
    assert set(filtered_report["suggestions"]) == {"title_from_filename"}


def test_smart_pdf_option_resolution_filters_detectable_and_keeps_unsupported(monkeypatch, tmp_path):
    pdf = tmp_path / "resolve.pdf"
    _write_pdf(pdf)

    def fake_precheck(_path, selected_options=None):
        assert selected_options == {"title_from_filename", "cleanup_remove_metadata", "page_size_a4"}
        return {
            "available": True,
            "suggestions": {
                "title_from_filename": {"matched": True, "title": "同步文件名为标题", "reason": "title"}
            },
        }

    monkeypatch.setattr(PDFProcessor, "build_precheck_report", fake_precheck)

    resolved = PDFProcessor.resolve_processing_options(
        str(pdf),
        {"title_from_filename", "cleanup_remove_metadata", "page_size_a4"},
        processing_mode="smart",
    )

    assert resolved["options"] == {"title_from_filename", "page_size_a4"}
    assert resolved["skipped"] == ["cleanup_remove_metadata"]
    assert resolved["forced_unsupported"] == ["page_size_a4"]
    assert "智能处理" in resolved["log"]


def test_process_document_smart_skips_selected_rule_without_precheck_match(tmp_path):
    pdf = tmp_path / "clean-title.pdf"
    out = tmp_path / "out.pdf"
    _write_pdf(pdf, title="clean-title")

    success, message = PDFProcessor.process_document(
        str(pdf),
        str(out),
        {"title_from_filename"},
        processing_mode="smart",
    )

    assert success is True
    assert "智能处理" in message
    assert "已跳过未命中规则" in message
    doc = fitz.open(out)
    try:
        assert doc.metadata.get("title") == "clean-title"
    finally:
        doc.close()


def test_process_document_force_executes_all_selected_rules(tmp_path):
    pdf = tmp_path / "force.pdf"
    out = tmp_path / "out.pdf"
    _write_pdf(pdf, title="force", author="Author")

    success, message = PDFProcessor.process_document(
        str(pdf),
        str(out),
        {"cleanup_remove_metadata"},
        processing_mode="force",
    )

    assert success is True
    assert "强制执行全部勾选规则" in message
    doc = fitz.open(out)
    try:
        assert not (doc.metadata.get("author") or "").strip()
    finally:
        doc.close()
