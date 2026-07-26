"""文件树交互子控制器：右键菜单、双击打开、定位、查看详情。"""

import os

from PySide6.QtCore import QObject
from PySide6.QtWidgets import QMenu

from ratools_pdf.services import pdf_inspector, system_shell


class TreeActionsController(QObject):
    def __init__(self, host, view, parent=None):
        super().__init__(parent)
        self.host = host
        self.view = view

    def show_tree_context_menu(self, pos):
        view = self.view
        selected_items = view.tree.selectedItems()
        if not selected_items:
            return

        # 菜单样式由应用级中央 QSS（theme.py 的 QMenu 段）提供，明暗主题自动适配
        menu = QMenu(view.tree)

        action_remove = menu.addAction("🗑️ 移除选中项")

        # 只有在选中单个文件/文件夹时，才允许执行详情查看和定位
        is_single_selection = len(selected_items) == 1
        target_path = selected_items[0].text(1) if is_single_selection else ""

        action_locate = menu.addAction("🔍 定位到文件位置")
        action_locate.setEnabled(is_single_selection)

        action_details = menu.addAction("📄 查看文件详情...")
        action_details.setEnabled(is_single_selection)

        # 映射坐标并在当前鼠标位置弹出
        action = menu.exec(view.tree.viewport().mapToGlobal(pos))

        if action == action_remove:
            self.host.remove_selected_items(selected_items)
        elif action == action_locate:
            self.locate_file(target_path)
        elif action == action_details:
            self.show_file_details(target_path)

    def on_item_double_clicked(self, item, column):
        """双击列表项直接使用系统默认软件打开 PDF 文件"""
        path = item.text(1)
        if not os.path.exists(path):
            self.view.show_warning_message("⚠️ 警告", "无法打开，该文件或文件夹可能已被移动或删除！")
            return

        # 仅打开文件（如果是PDF文件），如果是文件夹则展开/收起节点（由组件默认处理）
        if os.path.isfile(path) and path.lower().endswith('.pdf'):
            try:
                system_shell.open_with_default_app(path)
            except Exception as e:
                self.view.show_error_message("❌ 打开失败", f"无法使用默认程序打开文件：\n{str(e)}")

    def locate_file(self, path):
        """定位文件或文件夹位置（在系统文件资源管理器中打开并高亮显示）"""
        if not os.path.exists(path):
            self.view.show_warning_message("⚠️ 警告", "无法定位，该文件或文件夹可能已被移动或删除！")
            return

        try:
            system_shell.reveal_in_file_manager(path)
        except Exception as e:
            self.view.show_error_message("❌ 定位失败", f"无法打开系统资源管理器：\n{str(e)}")

    def show_file_details(self, path):
        """读取并弹窗显示选中项的系统属性以及 PDF 特有元数据"""
        if not os.path.exists(path):
            self.view.show_warning_message("⚠️ 警告", "无法读取信息，该文件或文件夹可能已被移动或删除！")
            return

        try:
            info_text = pdf_inspector.build_pdf_detail_text(path)
            self.view.show_info_message("📄 文件详细信息", info_text)
        except Exception as e:
            self.view.show_error_message("❌ 读取失败", f"获取文件信息时发生异常：\n{str(e)}")
