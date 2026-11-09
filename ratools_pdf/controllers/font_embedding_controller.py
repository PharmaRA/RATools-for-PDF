"""手动字体嵌入子控制器：选中 PDF 后引导用户在 Acrobat 中执行印前检查。"""

import os

from PySide6.QtCore import QObject

from ratools_pdf.services.font_providers import FontWorkflow


class FontEmbeddingController(QObject):
    def __init__(self, view, parent=None):
        super().__init__(parent)
        self.view = view
        self.font_workflow = FontWorkflow()

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
        opened = []
        failures = []
        for pdf_path in pdf_paths:
            result = self.font_workflow.open_manual(pdf_path)
            if result.opened:
                opened.append(pdf_path)
            else:
                failures.append(f"{os.path.basename(pdf_path)}：{result.reason}")

        if not opened:
            return False, "无法打开选中的 PDF：\n" + "\n".join(failures)

        message = (
            f"已交接打开 {len(opened)} 个 PDF；尚未确认字体已修复。\n\n"
            "请在 Acrobat 中执行：\n"
            "1. 所有工具 > 印刷制作 > 印前检查\n"
            "2. 选择“嵌入缺失的字体”\n"
            "3. 点击修复并保存\n\n"
            "保存后点击“保存后复检字体”，检查未嵌入、替代风险与状态未知。"
        )
        if not self.font_workflow.capability().acrobat_detected:
            message += "\n\n未定位到 Acrobat.exe，已改用系统默认 PDF 程序打开。"
        if failures:
            message += "\n\n部分文件未能打开：\n" + "\n".join(failures)
        return True, message

    def _recheck_paths(self, pdf_paths):
        all_verified = bool(pdf_paths)
        lines = []
        for pdf_path in pdf_paths:
            result = self.font_workflow.recheck(pdf_path)
            all_verified = all_verified and result.verified
            name = os.path.basename(pdf_path)
            if result.unembedded_count is None:
                lines.append(f"{name}：{result.reason}")
            else:
                lines.append(
                    f"{name}：未嵌入 {result.unembedded_count}；"
                    f"替代风险 {result.substitution_count}；状态未知 {result.unknown_count}"
                )
        headline = "字体复检通过：三类风险均为 0。" if all_verified else "字体复检未通过：仍有风险或无法完成检查。"
        return all_verified, headline + "\n" + "\n".join(lines)

    def _capability_text(self):
        capability = self.font_workflow.capability()
        version = capability.external_version or "未知（未验证）"
        message = (
            f"平台：{capability.platform}；provider：{capability.provider} v{capability.version}；"
            f"Acrobat 版本：{version}\n"
            f"自动修复不可用：{capability.unavailable_reasons.get('repair', '未提供受支持后端')}"
        )
        if "acrobat" in capability.unavailable_reasons:
            message += "\n" + capability.unavailable_reasons["acrobat"]
        if "open_manual" in capability.unavailable_reasons:
            message += "\n" + capability.unavailable_reasons["open_manual"]
        return message

    def open_selected_files_in_acrobat(self):
        pdf_paths = self._selected_pdf_paths()
        if not pdf_paths:
            self.view.show_warning_message("⚠️ 未选择 PDF", "请先在左侧待处理队列中选中需要嵌入缺失字体的 PDF 文件。")
            return

        self.view.show_manual_font_embedding_dialog(
            pdf_paths,
            lambda paths=tuple(pdf_paths): self._open_paths_in_acrobat(list(paths)),
            capability_text=self._capability_text(),
            recheck_callback=lambda paths=tuple(pdf_paths): self._recheck_paths(paths),
        )
