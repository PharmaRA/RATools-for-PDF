"""数字签名外观靶向烘焙固化模块。

提供在解除数字签名限制/权限脱壳前，将签名外观（手写笔迹、印章图章、签署人元数据框）
以高保真（默认 300 DPI 监管印刷级）精准烘焙压平至页面正文内容流，并安全解构签名控件。

核心保证：
- 零波及超链接 (/Link)：原文档所有内部跳转、目录索引、外部链接完全不受触碰。
- 零波及其他批注：文本高亮、划线、气泡批注等常规交互式注释保持完好。
- 零波及普通表单：仅针对 /FT /Sig（数字签名控件）进行物理固化，不强行抹平普通文本框。
- 根治阅读器外观丢失：解决 QPDF 解除限制后 Adobe Acrobat 因 AcroForm 孤立控件判定而隐藏签名的问题。
"""

import os
import shutil
from typing import Optional

import fitz


def has_digital_signatures(pdf_path: str, password: Optional[str] = None) -> bool:
    """检测 PDF 文件是否包含数字签名（包括 AcroForm SigFlags 或签名控件）。"""
    if not pdf_path or not os.path.isfile(pdf_path):
        return False

    doc = None
    try:
        doc = fitz.open(pdf_path)
        if doc.needs_pass:
            if password:
                if not doc.authenticate(password):
                    return False
            else:
                return False

        try:
            if doc.get_sigflags() > 0:
                return True
        except Exception:
            pass

        sig_type = getattr(fitz, "PDF_WIDGET_TYPE_SIGNATURE", None)
        for page in doc:
            for w in (page.widgets() or []):
                try:
                    if sig_type is not None and w.field_type == sig_type:
                        return True
                    if str(getattr(w, "field_type_string", "") or "").lower() == "signature":
                        return True
                    if "/Sig" in str(getattr(w, "field_name", "") or ""):
                        return True
                except Exception:
                    continue
        return False
    except Exception:
        return False
    finally:
        if doc is not None:
            try:
                doc.close()
            except Exception:
                pass


def bake_digital_signatures(
    input_pdf: str,
    output_pdf: str,
    password: Optional[str] = None,
    dpi: int = 300,
) -> int:
    """将 PDF 中所有的可见数字签名外观靶向烘焙压平为正文页面图像，并安全移除签名控件。

    Args:
        input_pdf: 输入 PDF 路径。
        output_pdf: 烘焙输出 PDF 路径（可与 input_pdf 相同，内部安全原子替换）。
        password: 文档打开或权限密码（如有）。
        dpi: 烘焙光栅化分辨率，默认为 300 DPI（监管及商业印刷标准）。

    Returns:
        int: 成功烘焙并移除的签名控件数量。若无签名或文档无需处理则返回 0。
    """
    if not input_pdf or not os.path.isfile(input_pdf):
        raise FileNotFoundError(f"输入文件不存在: {input_pdf}")

    doc = None
    try:
        doc = fitz.open(input_pdf)
        if doc.needs_pass:
            if password:
                if not doc.authenticate(password):
                    raise RuntimeError("密码错误，无法打开受密码保护的 PDF")
            else:
                # 需打开密码但未提供，跳过烘焙，由下游 QPDF 统一处理密码鉴权报错
                doc.close()
                doc = None
                if os.path.abspath(input_pdf) != os.path.abspath(output_pdf):
                    shutil.copy2(input_pdf, output_pdf)
                return 0

        sig_type = getattr(fitz, "PDF_WIDGET_TYPE_SIGNATURE", None)
        baked_count = 0

        for page in doc:
            # 1. 扫描当前页所有签名控件
            sig_widgets = []
            try:
                raw_widgets = list(page.widgets() or [])
            except Exception:
                raw_widgets = []

            for w in raw_widgets:
                try:
                    is_sig = False
                    if sig_type is not None and w.field_type == sig_type:
                        is_sig = True
                    elif str(getattr(w, "field_type_string", "") or "").lower() == "signature":
                        is_sig = True
                    elif "/Sig" in str(getattr(w, "field_name", "") or ""):
                        is_sig = True

                    if is_sig:
                        sig_widgets.append(w)
                except Exception:
                    continue

            if not sig_widgets:
                continue

            # 2. 逐一靶向提取外观、烘焙并安全移除
            for w in sig_widgets:
                try:
                    # 先安全备份坐标矩形（避免 delete_widget 后对象属性失效）
                    rect = fitz.Rect(w.rect)
                    if rect.is_empty or rect.width <= 0 or rect.height <= 0:
                        page.delete_widget(w)
                        baked_count += 1
                        continue

                    # 优先利用控件底层 _annot 提取独立透明外观（不受页面旋转干扰，纯净无背景干扰）
                    pix = None
                    if hasattr(w, "_annot") and w._annot:
                        try:
                            pix = w._annot.get_pixmap(dpi=dpi, alpha=True)
                        except Exception:
                            pix = None

                    # 兜底：若从 _annot 提取失败或无内容，尝试页面局部裁剪
                    if pix is None or pix.width <= 0 or pix.height <= 0:
                        try:
                            pix = page.get_pixmap(clip=rect, dpi=dpi, annots=True, alpha=True)
                        except Exception:
                            pix = None

                    # 检查是否包含实质可见像素
                    has_visible = False
                    if pix and pix.width > 0 and pix.height > 0:
                        if pix.alpha:
                            try:
                                has_visible = max(pix.samples[3::4]) > 0
                            except Exception:
                                has_visible = True
                        else:
                            has_visible = True

                    # 安全移除签名控件
                    page.delete_widget(w)
                    baked_count += 1

                    # 将外观图像物理固化写入正文流（覆盖在原矩形区域）
                    if has_visible and pix:
                        page.insert_image(rect, pixmap=pix, overlay=True, keep_proportion=False)

                except Exception:
                    # 单个控件容错处理
                    try:
                        page.delete_widget(w)
                        baked_count += 1
                    except Exception:
                        pass

        # 3. 结果保存与清理
        if baked_count > 0:
            # 清除 AcroForm 字典中的 SigFlags 标志位
            try:
                cat = doc.pdf_catalog()
                acroform_xref = doc.xref_get_key(cat, "AcroForm")
                if acroform_xref and acroform_xref[0] == "xref":
                    af_xref_num = int(acroform_xref[1].split()[0])
                    doc.xref_set_key(af_xref_num, "SigFlags", "0")
            except Exception:
                pass

            is_same = os.path.abspath(input_pdf) == os.path.abspath(output_pdf)
            save_target = output_pdf if not is_same else output_pdf + ".bake_tmp.pdf"
            doc.save(save_target, garbage=3, deflate=True)
            doc.close()
            doc = None
            if is_same:
                os.replace(save_target, output_pdf)
        else:
            doc.close()
            doc = None
            if os.path.abspath(input_pdf) != os.path.abspath(output_pdf):
                shutil.copy2(input_pdf, output_pdf)

        return baked_count
    finally:
        if doc is not None:
            try:
                doc.close()
            except Exception:
                pass
