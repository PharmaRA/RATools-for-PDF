"""PDF 只读检查的公开入口（供 controllers / UI 详情展示使用）。

controllers 此前直接调用 pdf 层的下划线私有函数；这里给出稳定的公开名，
pdf 层内部重构不再影响上层。全部函数只读，不修改任何 PDF。
"""

from ratools_pdf.pdf import precheck, qpdf

# 文件级信息
read_pdf_header_version = qpdf._read_pdf_header_version
is_pdf_linearized = qpdf._is_pdf_linearized
qpdf_reports_restrictions = qpdf._qpdf_reports_restrictions

# 文档结构
catalog_key_is_present = precheck._catalog_key_is_present
pdf_has_signature = precheck._pdf_has_signature

# 只读检测（文档检测模块）
collect_annotation_findings_for_path = precheck._collect_annotation_findings_for_path
collect_broken_reference_findings_for_path = precheck._collect_broken_reference_findings_for_path
