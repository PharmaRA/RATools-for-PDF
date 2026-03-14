import fitz  # 需安装 PyMuPDF 引擎: pip install PyMuPDF
from pathlib import Path


class PDFProcessor:
    """
    专门负责 PDF 底层操作的处理引擎。
    """

    @staticmethod
    def process_initial_view(input_path, output_path, options):
        """
        处理：初始视图与文档属性
        :param input_path: 输入文件路径
        :param output_path: 导出文件路径
        :param options: List[str] 用户勾选的规则名称
        :return: (bool, str) 是否成功，状态信息
        """
        try:
            doc = fitz.open(input_path)

            # 针对加密文件的基本防御：如果遇到需要密码的文件，则跳过
            if doc.needs_pass:
                return False, "❌ 文件已加密"

            changed = False
            catalog_xref = doc.pdf_catalog()  # 获取 PDF Catalog 的交叉引用编号

            # ----------------------------------------------------
            # 6. 根据文件名在PDF文档属性中自动添加文件标题
            # ----------------------------------------------------
            if "根据文件名在PDF文档属性中自动添加文件标题" in options:
                base_name = Path(input_path).stem  # 获取不带后缀的文件名
                meta = doc.metadata
                # 仅当原来没有标题，或标题不一致时才修改
                if meta.get("title") != base_name:
                    meta["title"] = base_name
                    doc.set_metadata(meta)
                    changed = True

            # ----------------------------------------------------
            # 1. 修改打开页面为第一页 / 3. 修改放大率为默认
            # ----------------------------------------------------
            if "修改打开页面为第一页" in options or "修改放大率为默认" in options:
                # 确保文档至少有一页
                if doc.page_count > 0:
                    page0_xref = doc[0].xref
                    # /Fit 代表“适合页面”， /XYZ null null null 代表保持当前或默认缩放比例
                    if "修改放大率为默认" in options:
                        action_str = f"[{page0_xref} 0 R /Fit]"
                    else:
                        action_str = f"[{page0_xref} 0 R /XYZ null null null]"

                    # 直接修改底层 Catalog 设置 OpenAction
                    doc.xref_set_key(catalog_xref, "OpenAction", action_str)
                    changed = True

            # ----------------------------------------------------
            # 2. 修改页面布局为默认 (SinglePage - 单页视图)
            # ----------------------------------------------------
            if "修改页面布局为默认" in options:
                doc.xref_set_key(catalog_xref, "PageLayout", "/SinglePage")
                changed = True

            # ----------------------------------------------------
            # 4. 修改导览标签 (UseOutlines - 默认强制打开左侧书签面板)
            # ----------------------------------------------------
            if "修改导览标签" in options:
                doc.xref_set_key(catalog_xref, "PageMode", "/UseOutlines")
                changed = True

            # ----------------------------------------------------
            # 5. PDF若存在书签则收起
            # ----------------------------------------------------
            if "PDF若存在书签则收起" in options:
                # 获取带有详细字典的目录树
                toc = doc.get_toc(simple=False)
                if toc:
                    for item in toc:
                        # PyMuPDF 的 toc 结构：[级别(int), 标题(str), 页码(int), {详细属性字典}]
                        if isinstance(item[-1], dict):
                            item[-1]["collapse"] = True  # 核心属性，折叠状态设为True
                    doc.set_toc(toc)
                    changed = True

            # ================= 保存输出文件 =================
            if changed:
                # garbage=3: 清除废弃的无用对象 (极佳的瘦身效果)
                # deflate=True: 对未压缩的流进行压缩
                doc.save(output_path, garbage=3, deflate=True)
            else:
                # 如果文件本身符合要求没被修改，直接拷贝过去
                doc.save(output_path)

            doc.close()
            return True, "✅ 处理成功"

        except Exception as e:
            return False, f"❌ 处理失败: {str(e)}"