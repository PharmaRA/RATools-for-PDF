"""文件详情文本生成（右键菜单"查看文件详情"）。纯函数，不依赖 Qt。"""

import os
from datetime import datetime

from ratools_pdf.pdf import inspect as pdf_inspect
from ratools_pdf.pdf.processor import PDFProcessor


def build_pdf_detail_text(path):
    """返回文件/文件夹/PDF 的多段详情文本；PDF 额外附结构信息与预检建议。"""
    stat = os.stat(path)
    created_time = datetime.fromtimestamp(stat.st_ctime).strftime('%Y-%m-%d %H:%M:%S')
    modified_time = datetime.fromtimestamp(stat.st_mtime).strftime('%Y-%m-%d %H:%M:%S')
    base_name = os.path.basename(path)
    size_kb = stat.st_size / 1024
    size_text = f"{size_kb / 1024:.2f} MB" if size_kb > 1024 else f"{size_kb:.2f} KB"

    details = [
        "📌 基础信息",
        f"文件名：{base_name}",
        f"路径：{path}",
        f"大小：{size_text}",
        f"创建时间：{created_time}",
        f"修改时间：{modified_time}",
    ]

    if not os.path.isfile(path):
        details.append("类型：文件夹")
        return "\n".join(details)

    if not path.lower().endswith('.pdf'):
        details.append("类型：普通文件")
        return "\n".join(details)

    doc = None
    try:
        import fitz
        doc = fitz.open(path)
        pdf_version = pdf_inspect.read_pdf_header_version(path) or "未知"
        linearized = "是" if pdf_inspect.is_pdf_linearized(path) else "否"
        restrictions = "是" if pdf_inspect.qpdf_reports_restrictions(path) else "否"

        details.extend([
            "",
            "📄 PDF信息",
            f"页数：{doc.page_count} 页",
            f"PDF版本：{pdf_version}",
            f"是否线性化：{linearized}",
            f"是否需要打开密码：{'是' if doc.needs_pass else '否'}",
            f"是否存在权限限制：{restrictions}",
        ])

        if doc.needs_pass:
            details.extend(["", "🛠️ 建议处理：该文件需要打开密码，无法展开内部结构预检"])
            return "\n".join(details)

        toc_count = len(doc.get_toc(simple=False))
        link_count = 0
        uri_link_count = 0
        annotation_count = 0
        for page in doc:
            links = page.get_links()
            link_count += len(links)
            uri_link_count += sum(1 for link in links if link.get("kind") == fitz.LINK_URI)
            annotation_count += sum(1 for _annot in (page.annots() or []))

        catalog_xref = doc.pdf_catalog()
        has_tags = (
            pdf_inspect.catalog_key_is_present(doc, catalog_xref, "StructTreeRoot")
            or pdf_inspect.catalog_key_is_present(doc, catalog_xref, "MarkInfo")
        )
        metadata_values = [
            value
            for key, value in (doc.metadata or {}).items()
            if key not in ["format", "encryption"] and str(value or "").strip()
        ]

        details.extend([
            "",
            "🧩 结构信息",
            f"书签数量：{toc_count}",
            f"页面链接数量：{link_count}",
            f"外部URI链接数量：{uri_link_count}",
            f"批注数量：{annotation_count}",
            f"内嵌附件数量：{doc.embfile_count()}",
            f"是否包含结构化标签：{'是' if has_tags else '否'}",
            f"是否包含元数据：{'是' if metadata_values else '否'}",
        ])

        report = PDFProcessor.build_precheck_report(path)
        suggestions = []
        for item in report.get("suggestions", {}).values():
            title = item.get("title", "")
            if not title:
                continue
            reason = item.get("reason", "")
            if item.get("report_only") and reason:
                suggestions.append(f"{title}：{reason}")
            else:
                suggestions.append(title)
        details.append("")
        if suggestions:
            details.append("🛠️ 建议处理：")
            details.extend(f"- {title}" for title in suggestions)
        else:
            details.append("🛠️ 建议处理：暂无明显建议项")

        return "\n".join(details)
    except Exception:
        details.extend(["", "⚠️ 提示：无法解析 PDF 内部结构，文件可能已损坏"])
        return "\n".join(details)
    finally:
        if doc is not None:
            try:
                doc.close()
            except Exception:
                pass
