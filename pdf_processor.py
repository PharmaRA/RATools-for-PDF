import fitz  # 需安装 PyMuPDF 引擎: pip install PyMuPDF
import os
import sys
import subprocess
import shutil
from pathlib import Path


class PDFProcessor:
    """
    专门负责 PDF 底层操作的处理引擎。
    结合了 PyMuPDF 和 Ghostscript 双擎驱动。
    """

    @staticmethod
    def _get_gs_path():
        """
        核心机制：智能解析 Ghostscript 引擎的绝对路径。
        兼容本地开发环境与 PyInstaller 打包后的临时解压环境 (_MEIPASS)。
        """
        if getattr(sys, 'frozen', False):
            # 如果是被 PyInstaller 打包后的运行环境
            base_path = sys._MEIPASS
        else:
            # 如果是本地 Python 脚本开发环境
            base_path = os.path.abspath(".")

        # 预设的引擎存放路径 (项目根目录/plugins/ghostscript/bin/gswin64c.exe)
        if sys.platform == "win32":
            gs_exe = os.path.join(base_path, "plugins", "ghostscript", "bin", "gswin64c.exe")
        elif sys.platform == "darwin":
            gs_exe = os.path.join(base_path, "plugins", "ghostscript", "bin", "gs")  # macOS
        else:
            gs_exe = "gs"  # Linux 默认调取系统环境变量

        return gs_exe

    @staticmethod
    def _embed_fonts_with_gs(input_pdf, output_pdf):
        """
        调用本地 Ghostscript 引擎，执行 PDF 深度重构与字体强制嵌入
        """
        gs_exe = PDFProcessor._get_gs_path()

        # 严格校验引擎文件是否存在 (Windows/Mac下)
        if sys.platform in ["win32", "darwin"] and not os.path.exists(gs_exe):
            raise FileNotFoundError(
                f"未找到 Ghostscript 引擎！\n"
                f"请确保已将引擎文件放置在: {gs_exe}"
            )

        # 工业级的 GS 字体嵌入静默执行命令
        cmd = [
            gs_exe,
            "-sDEVICE=pdfwrite",  # 输出为 PDF 设备
            "-dCompatibilityLevel=1.4",  # 兼容性设置
            "-dPDFSETTINGS=/prepress",  # 使用高质量的“预印”预设，默认会强制嵌入所有字体
            "-dNOPAUSE",  # 不暂停等待用户输入
            "-dQUIET",  # 静默模式，不输出冗余日志
            "-dBATCH",  # 处理完后自动退出
            "-dSubsetFonts=true",  # 允许字体子集化（减小体积）
            "-dEmbedAllFonts=true",  # 【核心】强制嵌入所有非标准字体
            f"-sOutputFile={output_pdf}",
            input_pdf
        ]

        # 隐藏 Windows 弹出的黑色 CMD 命令行窗口，实现完美的后台静默处理
        startupinfo = None
        if sys.platform == "win32":
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW

        # 阻塞执行并捕获报错
        result = subprocess.run(cmd, startupinfo=startupinfo, capture_output=True, text=True)
        if result.returncode != 0:
            raise RuntimeError(f"Ghostscript 执行失败: {result.stderr}")

    @staticmethod
    def process_document(input_path, output_path, options):
        """
        处理：统一的文档处理入口（双擎流转）
        """
        try:
            doc = fitz.open(input_path)

            if doc.needs_pass:
                return False, "❌ 文件已加密"

            changed = False
            catalog_xref = doc.pdf_catalog()

            # 判断用户是否勾选了需要调用 GS 引擎的重构项
            needs_gs_embed = "一键批量嵌入所有非标准字体（中文）" in options or "一键批量嵌入所有非标准字体（英文）" in options

            # ====================================================
            # 模块一：初始视图与文档属性
            # ====================================================
            if "根据文件名在PDF文档属性中自动添加文件标题" in options:
                base_name = Path(input_path).stem
                meta = doc.metadata
                if meta.get("title") != base_name:
                    meta["title"] = base_name
                    doc.set_metadata(meta)
                    changed = True

            if "修改打开页面为第一页" in options or "修改放大率为默认" in options:
                if doc.page_count > 0:
                    page0_xref = doc[0].xref
                    if "修改放大率为默认" in options:
                        action_str = f"[{page0_xref} 0 R /Fit]"
                    else:
                        action_str = f"[{page0_xref} 0 R /XYZ null null null]"
                    doc.xref_set_key(catalog_xref, "OpenAction", action_str)
                    changed = True

            if "修改页面布局为默认" in options:
                doc.xref_set_key(catalog_xref, "PageLayout", "/SinglePage")
                changed = True

            if "修改导览标签" in options:
                doc.xref_set_key(catalog_xref, "PageMode", "/UseOutlines")
                changed = True

            if "PDF若存在书签则收起" in options:
                toc = doc.get_toc(simple=False)
                if toc:
                    for item in toc:
                        if isinstance(item[-1], dict):
                            item[-1]["collapse"] = True
                    doc.set_toc(toc)
                    changed = True

            # ====================================================
            # 模块二：页面与字体标准化
            # ====================================================
            if "一键批量将页面切换成A4" in options or "一键批量将页面切换成Letter" in options:
                target_rect = fitz.paper_rect("a4") if "一键批量将页面切换成A4" in options else fitz.paper_rect(
                    "letter")
                for page in doc:
                    if abs(page.rect.width - target_rect.width) > 1 or abs(page.rect.height - target_rect.height) > 1:
                        page.set_mediabox(target_rect)
                        page.set_cropbox(target_rect)
                        changed = True

            # ================= 保存与双擎流转 =================
            if changed:
                if needs_gs_embed:
                    # 如果被修改过，且需要嵌入字体：先保存为临时文件，再交接给 GS 引擎重构
                    temp_pdf = str(output_path) + ".tmp.pdf"
                    doc.save(temp_pdf, garbage=3, deflate=True)
                    doc.close()
                    try:
                        PDFProcessor._embed_fonts_with_gs(temp_pdf, output_path)
                    finally:
                        if os.path.exists(temp_pdf):
                            os.remove(temp_pdf)  # 确保临时文件被清理
                else:
                    # 仅 PyMuPDF 处理
                    doc.save(output_path, garbage=3, deflate=True)
                    doc.close()
            else:
                doc.close()
                if needs_gs_embed:
                    # 属性未变，但需要嵌入字体：直接将原文件交接给 GS 引擎
                    PDFProcessor._embed_fonts_with_gs(input_path, output_path)
                else:
                    # 什么都没改，直接复制原文件
                    shutil.copy2(input_path, output_path)

            return True, "✅ 处理成功"

        except FileNotFoundError as e:
            return False, f"⚠️ 缺少组件: {str(e)}"
        except Exception as e:
            return False, f"❌ 处理失败: {str(e)}"