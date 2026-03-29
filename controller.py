import os
import re
import platform
import subprocess
from pathlib import Path
from datetime import datetime
from PySide6.QtWidgets import QFileDialog, QTreeWidgetItem, QMenu
from PySide6.QtCore import QObject, QThread, Signal, Qt
from PySide6.QtGui import QColor

from pdf_processor import PDFProcessor
from view import LogDialog


class ProcessWorker(QThread):
    """
    后台处理线程：负责核心的 PDF 批量规则应用，防止 UI 卡死
    """
    progress = Signal(int, str, str)  # row_index, status_text, log_message
    finished_all = Signal(str)  # summary
    error = Signal(str)  # error_msg

    def __init__(self, files, options, output_dir, common_base="", overwrite_original=False):
        super().__init__()
        self.files = files
        self.options = options
        self.output_dir = output_dir
        self.common_base = common_base
        self.overwrite_original = overwrite_original

    def run(self):
        try:
            success_count = 0
            rename_ectd = "filename_ectd_format" in self.options

            for i, file_path in enumerate(self.files):
                base_name = os.path.basename(file_path)
                self.progress.emit(i, "正在处理...", f"[{datetime.now().strftime('%H:%M:%S')}] 开始处理: {base_name}")

                # eCTD 命名合规处理
                if rename_ectd:
                    name, ext = os.path.splitext(base_name)
                    name = name.lower().replace(" ", "-")
                    name = re.sub(r'[^a-z0-9_-]', '', name)
                    if not name: name = f"doc_{i + 1:03d}"
                    base_name = f"{name}{ext.lower()}"

                # 决定输出路径 (支持保留原有文件夹层级结构)
                if self.overwrite_original:
                    out_path = file_path + ".tmp_overwrite.pdf"
                else:
                    if self.common_base:
                        file_dir = os.path.dirname(os.path.abspath(file_path))
                        rel_dir = os.path.relpath(file_dir, self.common_base)
                        if rel_dir == '.':
                            target_dir = self.output_dir
                        else:
                            target_dir = os.path.join(self.output_dir, rel_dir)
                    else:
                        target_dir = self.output_dir

                    os.makedirs(target_dir, exist_ok=True)
                    out_path = os.path.join(target_dir, base_name)

                # 调用核心引擎静态方法
                success, msg = PDFProcessor.process_document(file_path, out_path, self.options)

                if success and self.overwrite_original:
                    try:
                        os.replace(out_path, file_path)
                        out_path = file_path
                    except Exception as e:
                        success = False
                        msg = f"覆盖原文件失败: {str(e)}"
                        if os.path.exists(out_path):
                            os.remove(out_path)

                if success:
                    success_count += 1
                    status = "处理完成"
                else:
                    status = "处理失败"

                self.progress.emit(i, status, f"   ↳ 结果: {msg}")

            summary = f"处理结束。共成功处理 {success_count} / {len(self.files)} 个文件。"
            self.finished_all.emit(summary)

        except Exception as e:
            self.error.emit(str(e))


class IOActionWorker(QThread):
    """
    高级 IO 操作后台线程：处理书签、链接等需要长时间读写的批量导入/导出操作
    """
    progress = Signal(int, str, str)
    finished_action = Signal(str)
    error_action = Signal(str)

    def __init__(self, action_type, files, target_dir, output_dir=None):
        super().__init__()
        self.action_type = action_type
        self.files = files
        self.target_dir = target_dir
        self.output_dir = output_dir

    def run(self):
        try:
            for row_idx, file_path in enumerate(self.files):
                base_name = os.path.basename(file_path)
                name_no_ext, _ = os.path.splitext(base_name)

                self.progress.emit(row_idx, "正在执行...",
                                   f"[{datetime.now().strftime('%H:%M:%S')}] 正在处理: {base_name}")
                success, msg = False, ""

                if self.action_type == 'export_bookmarks':
                    csv_path = os.path.join(self.target_dir, f"{name_no_ext}_bookmarks.csv")
                    PDFProcessor.export_bookmarks(file_path, csv_path)
                    success, msg = True, "✅ 导出书签成功"

                elif self.action_type == 'import_bookmarks':
                    csv_path = os.path.join(self.target_dir, f"{name_no_ext}_bookmarks.csv")
                    if not os.path.exists(csv_path):
                        success, msg = False, "⚠️ 未找到匹配的CSV文件"
                    else:
                        out_pdf = os.path.join(self.output_dir, base_name)
                        PDFProcessor.import_bookmarks(file_path, csv_path, out_pdf)
                        success, msg = True, "✅ 导入书签成功"

                elif self.action_type == 'export_links':
                    json_path = os.path.join(self.target_dir, f"{name_no_ext}_links.json")
                    PDFProcessor.export_links(file_path, json_path)
                    success, msg = True, "✅ 导出链接成功"

                elif self.action_type == 'import_links':
                    json_path = os.path.join(self.target_dir, f"{name_no_ext}_links.json")
                    if not os.path.exists(json_path):
                        success, msg = False, "⚠️ 未找到匹配的JSON文件"
                    else:
                        out_pdf = os.path.join(self.output_dir, base_name)
                        PDFProcessor.import_links(file_path, json_path, out_pdf)
                        success, msg = True, "✅ 导入链接成功"

                status = "操作成功" if success else "操作失败"
                if "未找到匹配" in msg:
                    status = "未匹配跳过"
                self.progress.emit(row_idx, status, f"   ↳ 结果: {msg}")

            action_name = "导出" if "export" in self.action_type else "导入"
            self.finished_action.emit(f"批量高级 '{action_name}' 任务执行完成。")

        except Exception as e:
            self.error_action.emit(str(e))


class MainController(QObject):
    def __init__(self, view):
        super().__init__()
        self.view = view
        self.loaded_files = []
        self.process_logs = ""
        self.last_output_dir = ""

        # 建立缓存字典，以便快速在文件树中更新和查找节点
        self.folder_nodes = {}
        self.file_nodes = {}

        self.setup_connections()

    def setup_connections(self):
        self.view.drop_zone.files_dropped.connect(self.add_files)
        self.view.drop_zone.mousePressEvent = self.open_file_dialog
        self.view.btn_add_files.clicked.connect(self.open_file_picker)
        self.view.add_folder_btn.clicked.connect(self.add_folder)
        self.view.btn_clear.clicked.connect(self.clear_list)
        self.view.btn_preset_china.clicked.connect(lambda: self.view.toggle_preset("china"))
        self.view.btn_preset_us.clicked.connect(lambda: self.view.toggle_preset("us"))
        self.view.btn_clear_selected_options.clicked.connect(self.view.clear_selected_options)

        self.view.btn_start.clicked.connect(self.start_processing)
        self.view.btn_log.clicked.connect(self.show_log_dialog)

        self.view.btn_export_bookmarks.clicked.connect(lambda: self.handle_io_action('export_bookmarks'))
        self.view.btn_import_bookmarks.clicked.connect(lambda: self.handle_io_action('import_bookmarks'))
        self.view.btn_export_links.clicked.connect(lambda: self.handle_io_action('export_links'))
        self.view.btn_import_links.clicked.connect(lambda: self.handle_io_action('import_links'))

        self.setup_exclusive_options()

        # 绑定树形图的右键菜单请求事件
        self.view.tree.customContextMenuRequested.connect(self.show_tree_context_menu)
        # 绑定树形图双击事件
        self.view.tree.itemDoubleClicked.connect(self.on_item_double_clicked)

    # ================= 核心：右键菜单生成与分发 =================
    def show_tree_context_menu(self, pos):
        selected_items = self.view.tree.selectedItems()
        if not selected_items:
            return

        menu = QMenu(self.view.tree)
        # 使得菜单样式与整体 UI 现代感保持一致
        menu.setStyleSheet("""
            QMenu {
                background-color: white;
                border: 1px solid #D1D5DB;
                border-radius: 6px;
                padding: 4px;
            }
            QMenu::item {
                padding: 6px 28px 6px 20px;
                border-radius: 4px;
                color: #374151;
                font-size: 13px;
            }
            QMenu::item:selected {
                background-color: #F3F4F6;
                color: #2563EB;
            }
            QMenu::separator {
                height: 1px;
                background: #E5E7EB;
                margin: 4px 8px;
            }
            QMenu::item:disabled {
                color: #9CA3AF;
            }
        """)

        action_remove = menu.addAction("🗑️ 移除选中项")

        # menu.addSeparator()

        # 只有在选中单个文件/文件夹时，才允许执行详情查看和定位
        is_single_selection = len(selected_items) == 1
        target_path = selected_items[0].text(1) if is_single_selection else ""

        action_extend_1 = menu.addAction("🔍 定位到文件位置")
        action_extend_1.setEnabled(is_single_selection)

        action_extend_2 = menu.addAction("📄 查看文件详情...")
        action_extend_2.setEnabled(is_single_selection)

        # 映射坐标并在当前鼠标位置弹出
        action = menu.exec(self.view.tree.viewport().mapToGlobal(pos))

        if action == action_remove:
            self.remove_selected_items(selected_items)
        elif action == action_extend_1:
            self.locate_file(target_path)
        elif action == action_extend_2:
            self.show_file_details(target_path)

    def on_item_double_clicked(self, item, column):
        """双击列表项直接使用系统默认软件打开 PDF 文件"""
        path = item.text(1)
        if not os.path.exists(path):
            self.view.show_warning_message("⚠️ 警告", "无法打开，该文件或文件夹可能已被移动或删除！")
            return

        # 仅打开文件（如果是PDF文件），如果是文件夹则展开/收起节点（由组件默认处理）
        if os.path.isfile(path) and path.lower().endswith('.pdf'):
            sys_plat = platform.system()
            try:
                if sys_plat == "Windows":
                    os.startfile(path)
                elif sys_plat == "Darwin":
                    subprocess.Popen(["open", path])
                else:
                    subprocess.Popen(["xdg-open", path])
            except Exception as e:
                self.view.show_error_message("❌ 打开失败", f"无法使用默认程序打开文件：\n{str(e)}")

    def locate_file(self, path):
        """定位文件或文件夹位置（在系统文件资源管理器中打开并高亮显示）"""
        if not os.path.exists(path):
            self.view.show_warning_message("⚠️ 警告", "无法定位，该文件或文件夹可能已被移动或删除！")
            return

        sys_plat = platform.system()
        try:
            if sys_plat == "Windows":
                if os.path.isfile(path):
                    # Windows 下使用 explorer /select 高亮选中指定文件
                    subprocess.Popen(['explorer', '/select,', os.path.normpath(path)])
                else:
                    os.startfile(path)
            elif sys_plat == "Darwin":
                # macOS 使用 open -R 会在 Finder 中展示并选中文件
                subprocess.Popen(["open", "-R", path])
            else:
                # Linux 一般打开其所在目录
                target_dir = os.path.dirname(path) if os.path.isfile(path) else path
                subprocess.Popen(["xdg-open", target_dir])
        except Exception as e:
            self.view.show_error_message("❌ 定位失败", f"无法打开系统资源管理器：\n{str(e)}")

    def show_file_details(self, path):
        """读取并弹窗显示选中项的系统属性以及 PDF 特有元数据"""
        if not os.path.exists(path):
            self.view.show_warning_message("⚠️ 警告", "无法读取信息，该文件或文件夹可能已被移动或删除！")
            return

        try:
            stat = os.stat(path)
            created_time = datetime.fromtimestamp(stat.st_ctime).strftime('%Y-%m-%d %H:%M:%S')
            modified_time = datetime.fromtimestamp(stat.st_mtime).strftime('%Y-%m-%d %H:%M:%S')

            details = [f"📂 路径：{path}\n"]

            if os.path.isfile(path):
                size_kb = stat.st_size / 1024
                size_mb = size_kb / 1024
                if size_mb > 1:
                    details.append(f"📏 大小：{size_mb:.2f} MB")
                else:
                    details.append(f"📏 大小：{size_kb:.2f} KB")

                # 若是 PDF 文件，则利用 PyMuPDF 深入解析其内部元数据
                if path.lower().endswith('.pdf'):
                    try:
                        import fitz  # 局部导入以避免在顶部造成非必须依赖
                        doc = fitz.open(path)
                        details.append(f"📑 页数：{doc.page_count} 页")
                        if doc.needs_pass:
                            details.append("🔒 状态：文档已加密")
                        else:
                            meta = doc.metadata
                            if meta:
                                title = meta.get("title", "")
                                author = meta.get("author", "")
                                if title: details.append(f"📌 标题：{title}")
                                if author: details.append(f"👤 作者：{author}")
                        doc.close()
                    except Exception:
                        details.append("⚠️ 提示：无法解析 PDF 内部属性，文件可能已损坏")
            else:
                details.append("📁 类型：文件夹")

            details.append("")  # 空行作为分隔
            details.append(f"🕒 创建时间：{created_time}")
            details.append(f"⏱️ 修改时间：{modified_time}")

            info_text = "\n".join(details)
            self.view.show_info_message("📄 文件详细信息", info_text)

        except Exception as e:
            self.view.show_error_message("❌ 读取失败", f"获取文件信息时发生异常：\n{str(e)}")

    def remove_selected_items(self, selected_items):
        """
        处理树节点的移除操作。
        逻辑：递归收集所有选中的文件路径 -> 更新后台数据 -> 删除 UI 节点 -> 自动清理空文件夹
        """
        paths_to_remove = set()

        # 1. 内部递归函数：若选中的是文件夹，自动把下面的文件全部圈中
        def collect_paths(item):
            path = item.text(1)
            if path in self.file_nodes:
                paths_to_remove.add(path)
            for i in range(item.childCount()):
                collect_paths(item.child(i))

        for item in selected_items:
            collect_paths(item)

        # 2. 从后台数组和字典中彻底注销这些文件
        self.loaded_files = [p for p in self.loaded_files if p not in paths_to_remove]
        for p in paths_to_remove:
            if p in self.file_nodes:
                del self.file_nodes[p]

        # 3. 移除 UI 可视节点（注意：父节点被删除时，子节点自动消亡，需防止指针悬空）
        for item in selected_items:
            if item.treeWidget() is None:
                continue  # 该节点已经被随着父节点的删除而连带删除了

            parent = item.parent()
            if parent:
                parent.removeChild(item)
            else:
                self.view.tree.takeTopLevelItem(self.view.tree.indexOfTopLevelItem(item))

        # 4. 清理残留的、由于文件被移空而变成“孤儿”的空文件夹
        self._cleanup_empty_folders()

        # 5. 更新左下角的总数统计
        self.view.update_counters_ui(len(self.loaded_files))

    def _cleanup_empty_folders(self):
        """循环扫描并删除不再包含任何文件的空文件夹节点，以及已被从UI中移除的游离节点（Ghost Nodes）"""
        changed = True
        while changed:
            changed = False
            empty_paths = []

            for path, node in self.folder_nodes.items():
                # 1. 捕获游离的幽灵节点（用户直接删除了父文件夹，导致它脱离了UI树）
                if node.treeWidget() is None:
                    empty_paths.append(path)
                # 2. 捕获空文件夹（文件夹还在UI树上，但其内部的文件被逐一删空了）
                elif node.childCount() == 0:
                    empty_paths.append(path)

            for path in empty_paths:
                node = self.folder_nodes[path]

                # 如果节点还在 UI 树上，将其可视部分移除
                if node.treeWidget() is not None:
                    parent = node.parent()
                    if parent:
                        parent.removeChild(node)
                    else:
                        self.view.tree.takeTopLevelItem(self.view.tree.indexOfTopLevelItem(node))

                # 从后台缓存字典中彻底销毁该文件夹的记录
                del self.folder_nodes[path]
                changed = True

    def setup_exclusive_options(self):
        cb_a4 = self.view.all_checkboxes.get("page_size_a4")
        cb_letter = self.view.all_checkboxes.get("page_size_letter")
        if cb_a4 and cb_letter:
            cb_a4.toggled.connect(lambda checked: cb_letter.setChecked(False) if checked else None)
            cb_letter.toggled.connect(lambda checked: cb_a4.setChecked(False) if checked else None)

    def open_file_dialog(self, event):
        if event.button() == Qt.LeftButton:
            self.open_file_picker()

    def open_file_picker(self):
        file_paths, _ = QFileDialog.getOpenFileNames(
            self.view,
            "选择 PDF 文件",
            "",
            "PDF Files (*.pdf);;All Files (*)"
        )
        if file_paths:
            self.add_files(file_paths)

    def add_files(self, paths):
        valid_pdf_paths = []
        for p in paths:
            if os.path.isfile(p) and p.lower().endswith('.pdf'):
                valid_pdf_paths.append(os.path.normpath(p))
            elif os.path.isdir(p):
                for root, _, files in os.walk(p):
                    for file in files:
                        if file.lower().endswith('.pdf'):
                            valid_pdf_paths.append(os.path.normpath(os.path.join(root, file)))

        to_add = [p for p in valid_pdf_paths if p not in self.loaded_files]
        if not to_add:
            if paths:
                self.view.show_info_message("ℹ️ 提示", "添加的文件或文件夹中没有新的 PDF 文件，或文件已存在于列表中。")
            return

        # 智能算法：获取这一次批量拖入文件的公共根路径
        dirs = [os.path.dirname(os.path.abspath(p)) for p in to_add]
        try:
            common_base = os.path.commonpath(dirs)
        except ValueError:
            common_base = ""  # 如果跨盘符（如C盘和D盘），则降级使用绝对路径树

        for path in to_add:
            self.loaded_files.append(path)
            p = Path(path)
            parent_item = self.view.tree.invisibleRootItem()

            if common_base:
                # 挂载公共根目录节点
                if common_base not in self.folder_nodes:
                    root_node = QTreeWidgetItem(parent_item)
                    root_name = os.path.basename(common_base) or common_base
                    root_node.setText(0, f"📁 {root_name}")
                    root_node.setText(1, common_base)
                    root_node.setToolTip(0, root_name)
                    root_node.setToolTip(1, common_base)
                    root_node.setExpanded(True)
                    self.folder_nodes[common_base] = root_node

                parent_item = self.folder_nodes[common_base]

                # 动态生成中间补全目录
                rel_dir = os.path.relpath(os.path.dirname(path), common_base)
                if rel_dir != '.':
                    current_path = Path(common_base)
                    for part in Path(rel_dir).parts:
                        current_path = current_path / part
                        current_path_str = str(current_path)
                        if current_path_str not in self.folder_nodes:
                            node = QTreeWidgetItem(parent_item)
                            node.setText(0, f"📁 {part}")
                            node.setText(1, current_path_str)
                            node.setToolTip(0, part)
                            node.setToolTip(1, current_path_str)
                            node.setExpanded(True)
                            self.folder_nodes[current_path_str] = node
                        parent_item = self.folder_nodes[current_path_str]
            else:
                # 跨盘符降级处理，从硬盘根目录往下建树
                current_path = Path(p.parts[0])
                root_str = str(current_path)
                if root_str not in self.folder_nodes:
                    node = QTreeWidgetItem(parent_item)
                    node.setText(0, f"💽 {root_str}")
                    node.setText(1, root_str)
                    node.setToolTip(0, root_str)
                    node.setToolTip(1, root_str)
                    node.setExpanded(True)
                    self.folder_nodes[root_str] = node
                parent_item = self.folder_nodes[root_str]

                for part in p.parts[1:-1]:
                    current_path = current_path / part
                    current_path_str = str(current_path)
                    if current_path_str not in self.folder_nodes:
                        node = QTreeWidgetItem(parent_item)
                        node.setText(0, f"📁 {part}")
                        node.setText(1, current_path_str)
                        node.setToolTip(0, part)
                        node.setToolTip(1, current_path_str)
                        node.setExpanded(True)
                        self.folder_nodes[current_path_str] = node
                    parent_item = self.folder_nodes[current_path_str]

            # 挂载最终的文件节点
            file_node = QTreeWidgetItem(parent_item)
            file_node.setText(0, f"📄 {p.name}")
            file_node.setText(1, path)
            file_node.setText(2, "等待处理")
            file_node.setToolTip(0, p.name)
            file_node.setToolTip(1, path)
            file_node.setToolTip(2, "等待处理")
            file_node.setForeground(2, Qt.darkGray)

            # 将创建的文件节点加入字典中进行状态管理
            self.file_nodes[path] = file_node

        self.view.update_counters_ui(len(self.loaded_files))

    def add_folder(self):
        folder_path = QFileDialog.getExistingDirectory(self.view, "选择包含 PDF 的文件夹")
        if folder_path:
            self.add_files([folder_path])

    def clear_list(self):
        if not self.loaded_files:
            return

        if self.view.show_confirm_message("🗑️ 确认清空", "您确定要清空待处理文件树吗？"):
            self.loaded_files.clear()
            self.folder_nodes.clear()
            self.file_nodes.clear()
            self.view.clear_tree_ui()
            self.view.update_counters_ui(0)
            self.process_logs = ""

    def start_processing(self):
        if not self.loaded_files:
            self.view.show_warning_message("⚠️ 警告", "请至少添加一个 PDF 文件！")
            return

        selected_options = self.view.get_selected_options()
        if not selected_options:
            self.view.show_warning_message("⚠️ 警告", "请至少在右侧勾选一个处理规则！")
            return

        overwrite_cb = self.view.all_checkboxes.get("覆盖原始文件 (不推荐)")
        overwrite_original = overwrite_cb.isChecked() if overwrite_cb else False

        out_dir = ""
        common_base = ""

        if overwrite_original:
            if not self.view.show_confirm_message("⚠️ 危险操作确认",
                                                  "您勾选了【覆盖原始文件】。\n此操作不可逆，强烈建议您在操作前备份文件！\n\n是否继续？"):
                return
        else:
            user_selected_dir = QFileDialog.getExistingDirectory(self.view, "选择输出文件保存的根目录")
            if not user_selected_dir:
                return

            out_dir = os.path.join(user_selected_dir, "RATools_Output")
            self.last_output_dir = out_dir

            try:
                dirs = [os.path.dirname(os.path.abspath(f)) for f in self.loaded_files]
                common_base = os.path.commonpath(dirs)
            except ValueError:
                common_base = ""

        self.view.btn_start.setEnabled(False)
        self.view.btn_start.setText("处理中...")

        self.worker = ProcessWorker(self.loaded_files, selected_options, out_dir, common_base, overwrite_original)
        self.worker.progress.connect(self.update_progress)
        self.worker.finished_all.connect(self.processing_finished)
        self.worker.error.connect(self.processing_error)
        self.worker.start()

    def handle_io_action(self, action_type):
        if not self.loaded_files:
            self.view.show_warning_message("⚠️ 警告", "请先添加目标 PDF 文件！")
            return

        is_export = "export" in action_type
        data_type = "CSV" if "bookmarks" in action_type else "JSON"
        action_name = "导出" if is_export else "导入"

        dir_path = QFileDialog.getExistingDirectory(self.view, f"请选择 {data_type} 数据{action_name}的目录")
        if not dir_path:
            return

        out_dir = None
        if not is_export:
            first_file = Path(self.loaded_files[0])
            out_dir_path = first_file.parent / f"RATools_{action_name}完成"
            out_dir_path.mkdir(exist_ok=True)
            out_dir = str(out_dir_path)

        self.io_worker = IOActionWorker(action_type, self.loaded_files, dir_path, out_dir)
        self.io_worker.progress.connect(self.update_progress)
        self.io_worker.finished_action.connect(self.on_io_action_finished)
        self.io_worker.error_action.connect(self.on_io_action_error)
        self.io_worker.start()

    def update_progress(self, row_index, status_text, log_msg):
        # 获取与该行对应的精确文件路径，用于树节点的映射更新
        file_path = self.loaded_files[row_index]

        if status_text in ["处理完成", "操作成功"]:
            color = QColor(16, 185, 129)  # 绿色
        elif status_text in ["处理失败", "操作失败"]:
            color = QColor(239, 68, 68)  # 红色
        elif status_text == "未匹配跳过":
            color = QColor(245, 158, 11)  # 橙黄色警告
        else:
            color = QColor(37, 99, 235)  # 蓝色处理中

        # 查字典，直接更新树节点UI
        if file_path in self.file_nodes:
            node = self.file_nodes[file_path]
            node.setText(2, status_text)
            node.setToolTip(2, status_text)
            node.setForeground(2, color)

        self.process_logs += f"{log_msg}\n"

    def processing_finished(self, summary):
        self.process_logs += f"\n=== 批量处理结束 ===\n{summary}\n"
        self.view.btn_start.setEnabled(True)
        self.view.btn_start.setText("▶ 开始批量处理")

        self.view.show_success_message("✅ 处理完成", "所有 PDF 文件的批量处理任务已结束！")

        auto_open_cb = self.view.all_checkboxes.get("处理完成后自动打开输出文件夹")
        if auto_open_cb and auto_open_cb.isChecked() and self.loaded_files:
            overwrite_cb = self.view.all_checkboxes.get("覆盖原始文件 (不推荐)")
            if overwrite_cb and not overwrite_cb.isChecked():
                if hasattr(self, 'last_output_dir') and self.last_output_dir and os.path.exists(self.last_output_dir):
                    self._open_directory(self.last_output_dir)

    def processing_error(self, error_msg):
        self.process_logs += f"\n[致命错误] {error_msg}\n"
        self.view.btn_start.setEnabled(True)
        self.view.btn_start.setText("▶ 开始批量处理")
        self.view.show_error_message("❌ 处理异常", f"处理过程中发生错误：\n{error_msg}")

    def on_io_action_finished(self, result_msg):
        self.process_logs += f"\n{result_msg}\n"
        self.view.show_success_message("✅ 操作成功", result_msg)

    def on_io_action_error(self, error_msg):
        self.process_logs += f"\n[IO错误] {error_msg}\n"
        self.view.show_error_message("❌ 操作失败", error_msg)

    def show_log_dialog(self):
        if not hasattr(self, 'log_dialog'):
            self.log_dialog = LogDialog(self.view)
            self.log_dialog.btn_export.clicked.connect(self.export_logs)

        self.log_dialog.text_edit.setText(self.process_logs if self.process_logs else "暂无处理日志...")
        self.log_dialog.show()
        self.log_dialog.raise_()
        self.log_dialog.activateWindow()

    def export_logs(self):
        if not self.process_logs:
            self.view.show_warning_message("⚠️ 提示", "目前暂无任何日志可供导出！")
            return

        file_path, _ = QFileDialog.getSaveFileName(self.view, "导出处理日志", "process_logs.txt",
                                                   "Text Files (*.txt);;All Files (*)")
        if file_path:
            try:
                with open(file_path, "w", encoding="utf-8") as f:
                    f.write(self.process_logs)
                self.view.show_success_message("✅ 导出成功", "处理日志已成功保存！")
            except Exception as e:
                self.view.show_error_message("❌ 导出失败", f"文件保存失败：\n{str(e)}")

    def _open_directory(self, dir_path):
        sys_plat = platform.system()
        try:
            if sys_plat == "Windows":
                os.startfile(dir_path)
            elif sys_plat == "Darwin":
                subprocess.Popen(["open", dir_path])
            else:
                subprocess.Popen(["xdg-open", dir_path])
        except Exception as e:
            self.process_logs += f"\n[警告] 自动打开文件夹失败：{str(e)}\n"
