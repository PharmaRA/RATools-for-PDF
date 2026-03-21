import fitz
import os
import sys
import subprocess
import shutil
import csv
import json
from pathlib import Path


class PDFProcessor:
    """
    专门负责 PDF 底层操作的处理引擎。
    包含 PDF 合规处理与高级数据 IO(书签CSV导出导入、超链接JSON导出导入) 机制。
    """

    @staticmethod
    def _get_gs_path():
        if getattr(sys, 'frozen', False):
            base_path = sys._MEIPASS
        else:
            base_path = os.path.abspath(".")

        if sys.platform == "win32":
            gs_exe = os.path.join(base_path, "plugins", "ghostscript", "bin", "gswin64c.exe")
        elif sys.platform == "darwin":
            gs_exe = os.path.join(base_path, "plugins", "ghostscript", "bin", "gs")
        else:
            gs_exe = "gs"
        return gs_exe

    @staticmethod
    def _embed_fonts_with_gs(input_pdf, output_pdf):
        gs_exe = PDFProcessor._get_gs_path()
        if sys.platform in ["win32", "darwin"] and not os.path.exists(gs_exe):
            raise FileNotFoundError(f"未找到 Ghostscript 引擎！\n请确保已将引擎文件放置在: {gs_exe}")

        cmd = [
            gs_exe, "-sDEVICE=pdfwrite", "-dCompatibilityLevel=1.7", "-dPDFSETTINGS=/prepress",
            "-dNOPAUSE", "-dQUIET", "-dBATCH", "-dSubsetFonts=true", "-dEmbedAllFonts=true",
            f"-sOutputFile={output_pdf}", input_pdf
        ]

        startupinfo = None
        if sys.platform == "win32":
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW

        result = subprocess.run(cmd, startupinfo=startupinfo, capture_output=True, text=True)
        if result.returncode != 0:
            raise RuntimeError(f"Ghostscript 执行失败: {result.stderr}")

    # ====================================================
    # 数据 IO：导出与导入书签 (CSV)
    # ====================================================
    @staticmethod
    def export_bookmarks(pdf_path, csv_path):
        """将书签导出为 CSV (格式: 级别, 标题, 页码)"""
        doc = fitz.open(pdf_path)
        toc = doc.get_toc(simple=False)
        # 必须使用 utf-8-sig，防止 Excel 打开中文乱码
        with open(csv_path, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f)
            writer.writerow(['Level', 'Title', 'Page'])
            for item in toc:
                lvl, title, page, _ = item
                writer.writerow([lvl, title, page])
        doc.close()

    @staticmethod
    def import_bookmarks(pdf_path, csv_path, output_path):
        """读取 CSV 结构并强制写入 PDF 书签，带有防断层与防越界保护"""
        doc = fitz.open(pdf_path)
        new_toc = []

        with open(csv_path, 'r', encoding='utf-8-sig') as f:
            reader = csv.DictReader(f)
            for row in reader:
                try:
                    lvl = int(row.get('Level', 1))
                    title = row.get('Title', '')
                    page = int(row.get('Page', 1))

                    # 防越界：强制限制在真实页码范围内
                    page = max(1, min(page, doc.page_count))
                    new_toc.append([lvl, title, page])
                except ValueError:
                    continue  # 跳过无法解析格式的脏数据行

        # 防断层算法：拉平非法的书签级别跨越 (例如: 1 -> 3 会报错，强制转为 1 -> 2)
        if new_toc:
            for i in range(len(new_toc)):
                if i == 0:
                    new_toc[i][0] = 1
                else:
                    prev_lvl = new_toc[i - 1][0]
                    if new_toc[i][0] > prev_lvl + 1:
                        new_toc[i][0] = prev_lvl + 1

        doc.set_toc(new_toc)
        doc.save(output_path, garbage=3, deflate=True)
        doc.close()

    # ====================================================
    # 数据 IO：导出与导入超链接 (JSON)
    # ====================================================
    @staticmethod
    def export_links(pdf_path, json_path):
        """将超链接的物理坐标与动作类型提取至 JSON 文件"""
        doc = fitz.open(pdf_path)
        all_links = []

        for page in doc:
            for link in page.get_links():
                rect = link['from']
                link_dict = {
                    'page_index': page.number,  # PyMuPDF 中页面索引从 0 开始
                    'rect': [rect.x0, rect.y0, rect.x1, rect.y1],
                    'kind': link.get('kind', fitz.LINK_NONE),
                    'uri': link.get('uri', ''),
                    'file': link.get('file', ''),
                    'target_page': link.get('page', 0),  # 仅 GOTO 有效
                    'zoom': link.get('zoom', 0.0)
                }
                all_links.append(link_dict)

        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(all_links, f, indent=4, ensure_ascii=False)
        doc.close()

    @staticmethod
    def import_links(pdf_path, json_path, output_path):
        """清除原有链接，根据 JSON 精准复原超链接布局"""
        doc = fitz.open(pdf_path)

        with open(json_path, 'r', encoding='utf-8') as f:
            links_data = json.load(f)

        # 先清空原有超链接，防止重复叠加
        for page in doc:
            for link in page.get_links():
                page.delete_link(link)

        for ld in links_data:
            p_idx = ld.get('page_index', 0)
            if 0 <= p_idx < doc.page_count:
                page = doc[p_idx]
                rect = fitz.Rect(ld['rect'])
                kind = ld['kind']

                # 构建 PyMuPDF 所需的动作字典
                new_link = {"kind": kind, "from": rect}
                if kind == fitz.LINK_URI:
                    new_link["uri"] = ld.get('uri', '')
                elif kind == fitz.LINK_FILE:
                    new_link["file"] = ld.get('file', '')
                elif kind in [fitz.LINK_GOTO, fitz.LINK_GOTOR]:
                    new_link["page"] = ld.get('target_page', 0)
                    new_link["zoom"] = ld.get('zoom', 0.0)
                    if kind == fitz.LINK_GOTOR:
                        new_link["file"] = ld.get('file', '')

                try:
                    page.insert_link(new_link)
                except Exception:
                    pass  # 忽略错误坐标导致无法注入的脏链接

        doc.save(output_path, garbage=3, deflate=True)
        doc.close()

    # ====================================================
    # 六大核心合规清理模块入口
    # ====================================================
    @staticmethod
    def process_document(input_path, output_path, options):
        try:
            doc = fitz.open(input_path)

            if doc.needs_pass: return False, "❌ 文件已加密"
            changed = False
            catalog_xref = doc.pdf_catalog()

            needs_gs_engine = any(opt in options for opt in [
                "一键批量嵌入所有非标准字体（中文）", "一键批量嵌入所有非标准字体（英文）", "PDF版本转换"
            ])

            if "根据文件名在PDF文档属性中自动添加文件标题" in options:
                base_name = Path(input_path).stem
                meta = doc.metadata
                if meta.get("title") != base_name:
                    meta["title"] = base_name;
                    doc.set_metadata(meta);
                    changed = True

            if "修改打开页面为第一页" in options or "修改放大率为默认" in options:
                if doc.page_count > 0:
                    page0_xref = doc[0].xref
                    action_str = f"[{page0_xref} 0 R /Fit]" if "修改放大率为默认" in options else f"[{page0_xref} 0 R /XYZ null null null]"
                    doc.xref_set_key(catalog_xref, "OpenAction", action_str);
                    changed = True

            if "修改页面布局为默认" in options:
                doc.xref_set_key(catalog_xref, "PageLayout", "/SinglePage");
                changed = True

            if "修改导览标签" in options:
                doc.xref_set_key(catalog_xref, "PageMode", "/UseOutlines");
                changed = True

            if "PDF若存在书签则收起" in options:
                toc = doc.get_toc(simple=False)
                if toc:
                    for item in toc:
                        if isinstance(item[-1], dict): item[-1]["collapse"] = True
                    doc.set_toc(toc);
                    changed = True

            if "一键批量将页面切换成A4" in options or "一键批量将页面切换成Letter" in options:
                target_rect = fitz.paper_rect("a4") if "一键批量将页面切换成A4" in options else fitz.paper_rect(
                    "letter")
                for page in doc:
                    if abs(page.rect.width - target_rect.width) > 1 or abs(page.rect.height - target_rect.height) > 1:
                        page.set_mediabox(target_rect);
                        page.set_cropbox(target_rect);
                        changed = True

            bookmark_options = ["修改书签设置为承前缩放", "修改书签的设置为在新窗口中打开", "删除书签的外部链接",
                                "删除失效的书签（即未分配任何操作的书签）",
                                "删除未知动作的书签（即GoTo, GoToR和Launch之外的书签）"]
            if any(opt in options for opt in bookmark_options):
                toc = doc.get_toc(simple=False)
                if toc:
                    new_toc = []
                    toc_modified = False
                    for item in toc:
                        lvl, title, page, dest = item
                        kind = dest.get("kind", fitz.LINK_NONE)
                        delete_it = False

                        if "删除书签的外部链接" in options and kind == fitz.LINK_URI: delete_it = True
                        if "删除失效的书签（即未分配任何操作的书签）" in options:
                            if kind == fitz.LINK_NONE or (
                                    kind == fitz.LINK_GOTO and (page < 1 or page > doc.page_count)): delete_it = True
                        if "删除未知动作的书签（即GoTo, GoToR和Launch之外的书签）" in options:
                            if kind not in [fitz.LINK_GOTO, fitz.LINK_GOTOR, fitz.LINK_LAUNCH]: delete_it = True

                        if delete_it:
                            toc_modified = True;
                            continue

                        if "修改书签设置为承前缩放" in options and kind == fitz.LINK_GOTO:
                            if dest.get("zoom") != 0.0: dest["zoom"] = 0.0; toc_modified = True
                        if "修改书签的设置为在新窗口中打开" in options and kind in [fitz.LINK_GOTOR, fitz.LINK_LAUNCH]:
                            if not dest.get("newWindow"): dest["newWindow"] = True; toc_modified = True

                        if kind == fitz.LINK_GOTO:
                            if page < 1:
                                page = 1; toc_modified = True
                            elif page > doc.page_count:
                                page = doc.page_count; toc_modified = True

                        new_toc.append([lvl, title, page, dest])

                    if toc_modified:
                        if new_toc:
                            for i in range(len(new_toc)):
                                if i == 0:
                                    new_toc[i][0] = 1
                                else:
                                    prev_lvl = new_toc[i - 1][0]
                                    if new_toc[i][0] > prev_lvl + 1: new_toc[i][0] = prev_lvl + 1
                        doc.set_toc(new_toc);
                        changed = True

            hyperlink_options = ["将外链接中的绝对路径转相对路径", "修改超链接的设置为承前缩放",
                                 "修改超链接的设置为在新窗口中打开", "修改超链接文本至蓝色字体",
                                 "修改超链接文本至黑色边框", "超链接有边框则蓝框黑字", "超链接无边框且蓝字则蓝框黑字",
                                 "删除超链接边框"]
            if any(opt in options for opt in hyperlink_options):
                for page in doc:
                    links = page.get_links()
                    for link in links:
                        link_modified = False
                        kind = link.get("kind", fitz.LINK_NONE)

                        if "将外链接中的绝对路径转相对路径" in options and kind == fitz.LINK_FILE:
                            file_path = link.get("file", "")
                            if file_path and (
                                    ":" in file_path or file_path.startswith("/") or file_path.startswith("\\")):
                                link["file"] = os.path.basename(file_path.replace("\\", "/"));
                                link_modified = True
                        if "修改超链接的设置为承前缩放" in options and kind == fitz.LINK_GOTO:
                            if link.get("zoom") != 0.0: link["zoom"] = 0.0; link_modified = True
                        if "修改超链接的设置为在新窗口中打开" in options and kind in [fitz.LINK_GOTOR,
                                                                                      fitz.LINK_LAUNCH]:
                            if not link.get("newWindow"): link["newWindow"] = True; link_modified = True

                        if link_modified: page.update_link(link); changed = True

                    for annot in page.annots():
                        if annot.type[0] == 8:  # 8 代表 LINK 注释
                            annot_changed = False
                            border = annot.border
                            has_border = border and border.get("width", 0) > 0

                            def is_text_blue(rect):
                                text_dict = page.get_text("dict", clip=rect)
                                for block in text_dict.get("blocks", []):
                                    for line in block.get("lines", []):
                                        for span in line.get("spans", []):
                                            color = span.get("color", 0)
                                            b = color & 0xFF;
                                            g = (color >> 8) & 0xFF;
                                            r = (color >> 16) & 0xFF
                                            if b > r + 40 and b > g + 40: return True
                                return False

                            if "删除超链接边框" in options:
                                if has_border: annot.set_border(width=0); annot_changed = True
                            elif "修改超链接文本至黑色边框" in options:
                                annot.set_border(width=1.0);
                                annot.set_colors(stroke=(0, 0, 0));
                                annot_changed = True
                            elif "修改超链接文本至蓝色字体" in options:
                                annot.set_border(width=1.0);
                                annot.set_colors(stroke=(0, 0, 1));
                                annot_changed = True
                            elif "超链接有边框则蓝框黑字" in options:
                                if has_border: annot.set_border(width=1.0); annot.set_colors(
                                    stroke=(0, 0, 1)); annot_changed = True
                            elif "超链接无边框且蓝字则蓝框黑字" in options:
                                if not has_border and is_text_blue(annot.rect):
                                    annot.set_border(width=1.0);
                                    annot.set_colors(stroke=(0, 0, 1));
                                    annot_changed = True

                            if annot_changed: annot.update(); changed = True

            cleanup_options = ["删除外部链接（网页、邮箱地址）", "删除外部链接（网页、邮箱地址）且将文字改成黑色",
                               "删除失效的链接（即未分配任何操作的链接）", "删除无效的超链接，且将文字改成黑色",
                               "删除未知动作的链接（即GoTo, GoToRi和Launch之外的书签之外的链接）",
                               "删除JavaScript, 3D内容或者动态内容", "删除文档附件", "删除文档标签", "删除PDF注释",
                               "删除文档说明", "删除所有链接和书签"]
            if any(opt in options for opt in cleanup_options):
                if "删除所有链接和书签" in options:
                    doc.set_toc([])
                    for page in doc:
                        for link in page.get_links(): page.delete_link(link)
                    changed = True
                else:
                    for page in doc:
                        for link in page.get_links():
                            kind = link.get("kind", fitz.LINK_NONE)
                            delete_it = False
                            if kind == fitz.LINK_URI and (
                                    "删除外部链接（网页、邮箱地址）" in options or "删除外部链接（网页、邮箱地址）且将文字改成黑色" in options): delete_it = True
                            if kind == fitz.LINK_NONE and (
                                    "删除失效的链接（即未分配任何操作的链接）" in options or "删除无效的超链接，且将文字改成黑色" in options): delete_it = True
                            if "删除未知动作的链接（即GoTo, GoToRi和Launch之外的链接）" in options and kind not in [
                                fitz.LINK_GOTO, fitz.LINK_GOTOR, fitz.LINK_LAUNCH]: delete_it = True
                            if delete_it: page.delete_link(link); changed = True

                if "删除PDF注释" in options:
                    for page in doc:
                        for annot in page.annots():
                            if annot.type[0] != 8: page.delete_annot(annot); changed = True

                if "删除JavaScript, 3D内容或者动态内容" in options:
                    doc.xref_set_key(catalog_xref, "Names", "null");
                    changed = True
                if "删除文档附件" in options:
                    if doc.embfile_count() > 0:
                        for emb in doc.embfile_names(): doc.embfile_del(emb)
                        changed = True
                if "删除文档标签" in options:
                    doc.xref_set_key(catalog_xref, "StructTreeRoot", "null")
                    doc.xref_set_key(catalog_xref, "MarkInfo", "null");
                    changed = True
                if "删除文档说明" in options:
                    doc.set_metadata({});
                    doc.xref_set_key(catalog_xref, "PieceInfo", "null");
                    changed = True

            is_linear = "修改文件为快速网页浏览" in options
            if changed or is_linear:
                if needs_gs_engine:
                    temp_pdf = str(output_path) + ".tmp.pdf"
                    doc.save(temp_pdf, garbage=3, deflate=True, linear=is_linear)
                    doc.close()
                    try:
                        PDFProcessor._embed_fonts_with_gs(temp_pdf, output_path)
                    finally:
                        if os.path.exists(temp_pdf): os.remove(temp_pdf)
                else:
                    doc.save(output_path, garbage=3, deflate=True, linear=is_linear);
                    doc.close()
            else:
                doc.close()
                if needs_gs_engine:
                    PDFProcessor._embed_fonts_with_gs(input_path, output_path)
                else:
                    shutil.copy2(input_path, output_path)

            return True, "✅ 处理成功"

        except FileNotFoundError as e:
            return False, f"⚠️ 缺少引擎组件: {str(e)}"
        except Exception as e:
            return False, f"❌ 处理失败: {str(e)}"