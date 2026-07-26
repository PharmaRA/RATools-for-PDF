"""手动字体嵌入子控制器：选中 PDF 后引导用户在 Acrobat 中执行印前检查。"""

import os

from PySide6.QtCore import QObject

from ratools_pdf.services import system_shell


class FontEmbeddingController(QObject):
    def __init__(self, view, parent=None):
        super().__init__(parent)
        self.view = view

    def _selected_pdf_paths(self):
        selected_items = self.view.tree.selectedItems()
        pdf_paths = []
        seen = set()
        for item in selected_items:
            path = str(item.text(1) or "").strip().strip('"')
            if not path or not os.path.isfile(path) or not path.lower().endswith(".pdf"):
                continue
            key = os.path.normcase(os.path.abspath(path))
            if key in seen:
                continue
            seen.add(key)
            pdf_paths.append(path)
        return pdf_paths

    def _open_paths_in_acrobat(self, pdf_paths):
        acrobat_path = system_shell.find_acrobat_executable()
        opened = []
        failures = []
        for pdf_path in pdf_paths:
            try:
                system_shell.open_pdf_in_acrobat_or_default(pdf_path, acrobat_path)
                opened.append(pdf_path)
            except Exception as exc:
                failures.append(f"{os.path.basename(pdf_path)}：{exc}")

        if not opened:
            return False, "无法打开选中的 PDF：\n" + "\n".join(failures)

        message = (
            f"已打开 {len(opened)} 个 PDF。\n\n"
            "请在 Acrobat 中执行：\n"
            "1. 所有工具 > 印刷制作 > 印前检查\n"
            "2. 选择“嵌入缺失的字体”\n"
            "3. 点击修复并保存\n\n"
            "处理完成后，回到 RATools 重新执行预检确认字体风险是否消失。"
        )
        if not acrobat_path:
            message += "\n\n未定位到 Acrobat.exe，已改用系统默认 PDF 程序打开。"
        if failures:
            message += "\n\n部分文件未能打开：\n" + "\n".join(failures)
        return True, message

    def open_selected_files_in_acrobat(self):
        pdf_paths = self._selected_pdf_paths()
        if not pdf_paths:
            self.view.show_warning_message("⚠️ 未选择 PDF", "请先在左侧待处理队列中选中需要嵌入缺失字体的 PDF 文件。")
            return

        self.view.show_manual_font_embedding_dialog(
            pdf_paths,
            lambda paths=tuple(pdf_paths): self._open_paths_in_acrobat(list(paths)),
        )
