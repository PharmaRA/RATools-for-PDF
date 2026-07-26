import os
from datetime import datetime
from pathlib import Path

from PySide6.QtCore import QCoreApplication, QObject, Qt, QTimer
from PySide6.QtGui import QColor
from PySide6.QtWidgets import QFileDialog, QTreeWidgetItem

from ratools_pdf.common.status import status_semantic
from ratools_pdf.config.features import ENABLE_UPDATE_CHECK
from ratools_pdf.controllers.io_actions import (
    _collect_ectd_rename_plan,
    common_base_dir,
)
from ratools_pdf.controllers.log_export import _structured_log_row_from_event
from ratools_pdf.controllers.detection_controller import DetectionController
from ratools_pdf.controllers.font_embedding_controller import FontEmbeddingController
from ratools_pdf.controllers.io_controller import IOController
from ratools_pdf.controllers.log_controller import LogController
from ratools_pdf.controllers.precheck_controller import PrecheckController
from ratools_pdf.controllers.tree_actions_controller import TreeActionsController
from ratools_pdf.controllers.update_controller import UpdateController
from ratools_pdf.controllers.workers import ProcessWorker
from ratools_pdf.pdf import inspect as pdf_inspect
from ratools_pdf.services import system_shell
from ratools_pdf.ui.theme import active_palette, tree_status_color


class MainController(QObject):
    def __init__(self, view):
        super().__init__()
        self.view = view
        self.loaded_files = []
        self.process_logs = ""
        self.process_log_rows = []
        self.process_log_starts = {}
        self.last_output_dir = ""
        self.processing_started_at = None
        self.processing_total = 0
        self.processing_done = 0
        self.processing_done_paths = set()
        self.processing_files = []
        self.processing_current_file = ""
        self.processing_parallel_mode = False
        self.processing_worker_count = 1
        self._last_processing_hint = ""
        self.last_failed_files = []
        self.batch_result_counts = {"success": 0, "failure": 0, "skip": 0}
        self.processing_timer = QTimer(self)
        self.processing_timer.setInterval(1000)
        self.processing_timer.timeout.connect(self._refresh_processing_hint)

        # 建立缓存字典，以便快速在文件树中更新和查找节点
        self.folder_nodes = {}
        self.file_nodes = {}

        self.worker = None
        self.detection = DetectionController(self, self.view, parent=self)
        self.io = IOController(self, self.view, parent=self)
        self.font_embedding = FontEmbeddingController(self.view, parent=self)
        self.logs = LogController(self, self.view, parent=self)
        self.precheck = PrecheckController(self, self.view, parent=self)
        self.tree_actions = TreeActionsController(self, self.view, parent=self)
        self.updates = UpdateController(self.view, parent=self)

        self.setup_connections()
        app = QCoreApplication.instance()
        if app:
            app.aboutToQuit.connect(self.updates.shutdown)

        # 树状态列颜色是逐节点 setForeground 写入的，主题切换后必须整树重刷，
        # 否则保留上一主题的前景色（暗底旧亮色对比度不足）。
        theme_manager = getattr(self.view, "theme_manager", None)
        if theme_manager is not None:
            theme_manager.changed.connect(self._recolor_tree_statuses)

    def setup_connections(self):
        self.view.drop_zone.files_dropped.connect(self.add_files)
        self.view.drop_zone.mousePressEvent = self.open_file_dialog
        self.view.btn_add_files.clicked.connect(self.open_file_picker)
        self.view.add_folder_btn.clicked.connect(self.add_folder)
        self.view.btn_clear.clicked.connect(self.clear_list)
        self.view.btn_preset_china.clicked.connect(lambda: self.view.toggle_preset("china"))
        self.view.btn_preset_us.clicked.connect(lambda: self.view.toggle_preset("us"))
        self.view.btn_preset_favorite.clicked.connect(lambda: self.view.toggle_preset("favorite"))
        self.view.btn_save_favorite_preset.clicked.connect(self.view.save_favorite_preset)
        self.view.btn_clear_selected_options.clicked.connect(self.view.clear_selected_options)
        self.view.btn_skip_current.clicked.connect(self.skip_current_file)

        self.view.btn_retry_failed.clicked.connect(self.start_retry_failed_processing)
        self.view.btn_apply_precheck.clicked.connect(self.precheck.apply_precheck_suggestions)
        self.view.btn_process_precheck_suggested.clicked.connect(self.precheck.start_precheck_suggested_processing)
        self.view.btn_precheck.clicked.connect(self.precheck.start_precheck)
        self.view.btn_start.clicked.connect(lambda _checked=False: self.start_processing())
        self.view.btn_log.clicked.connect(self.logs.show_log_dialog)
        btn_embed_missing_fonts = getattr(self.view, "btn_embed_missing_fonts", None)
        if btn_embed_missing_fonts is not None:
            btn_embed_missing_fonts.clicked.connect(self.font_embedding.open_selected_files_in_acrobat)
        if ENABLE_UPDATE_CHECK:
            self.view.btn_top_about.clicked.connect(self.updates.wire_about_dialog)

        self.view.btn_bookmark_io_wizard.clicked.connect(lambda: self.io.handle_io_wizard("bookmarks"))
        self.view.btn_link_io_wizard.clicked.connect(lambda: self.io.handle_io_wizard("links"))
        self.view.btn_detect_annotations.clicked.connect(lambda: self.detection.start_detection("annotations"))
        self.view.btn_detect_broken_refs.clicked.connect(lambda: self.detection.start_detection("broken_refs"))

        self.setup_exclusive_options()

        # 绑定树形图的右键菜单请求事件
        self.view.tree.customContextMenuRequested.connect(self.tree_actions.show_tree_context_menu)
        # 绑定树形图双击事件
        self.view.tree.itemDoubleClicked.connect(self.tree_actions.on_item_double_clicked)

    def _is_precheck_running(self):
        return self.precheck.is_running()

    @property
    def last_precheck_results(self):
        return self.precheck.last_precheck_results

    def ensure_idle(self, action_label):
        """统一忙碌互斥守卫：有任务在跑时弹提示并返回 False。

        action_label 用于拼接提示文案，如 "移除队列项" → "请等待当前批量处理结束后再移除队列项。"
        """
        if self.worker and self.worker.isRunning():
            self.view.show_warning_message("⚠️ 正在处理中", f"请等待当前批量处理结束后再{action_label}。")
            return False
        if self._is_precheck_running():
            self.view.show_warning_message("⚠️ 正在预检", f"请等待当前预检结束后再{action_label}。")
            return False
        if self._is_detection_running():
            self.view.show_warning_message("⚠️ 正在检测", f"请等待当前检测结束后再{action_label}。")
            return False
        return True

    def _mark_precheck_stale(self):
        self.precheck.mark_stale()

    def _get_processing_worker_count(self):
        settings_dialog = self.view.settings_dialog
        if not settings_dialog.cb_parallel_processing.isChecked():
            return 1
        try:
            worker_count = max(2, int(settings_dialog.spin_parallel_workers.value()))
        except Exception:
            worker_count = 2
        max_workers = getattr(settings_dialog, "parallel_max_workers", None)
        if max_workers is None and hasattr(settings_dialog.spin_parallel_workers, "maximum"):
            max_workers = settings_dialog.spin_parallel_workers.maximum()
        try:
            max_workers = max(2, int(max_workers))
        except Exception:
            max_workers = worker_count
        return min(worker_count, max_workers)

    def check_updates_on_startup(self):
        self.updates.check_updates_on_startup()

    def remove_selected_items(self, selected_items):
        """
        处理树节点的移除操作。
        逻辑：递归收集所有选中的文件路径 -> 更新后台数据 -> 删除 UI 节点 -> 自动清理空文件夹
        """
        if not self.ensure_idle("移除队列项"):
            return

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
        if paths_to_remove:
            self._mark_precheck_stale()
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

        cb_remove_external = self.view.all_checkboxes.get("cleanup_remove_external_uri")
        cb_remove_external_black = self.view.all_checkboxes.get("cleanup_remove_external_uri_and_text_black")
        if cb_remove_external and cb_remove_external_black:
            cb_remove_external.toggled.connect(
                lambda checked: cb_remove_external_black.setChecked(False) if checked else None
            )
            cb_remove_external_black.toggled.connect(
                lambda checked: cb_remove_external.setChecked(False) if checked else None
            )

        cb_remove_invalid = self.view.all_checkboxes.get("cleanup_remove_invalid_links")
        cb_remove_invalid_black = self.view.all_checkboxes.get("cleanup_remove_invalid_links_and_text_black")
        if cb_remove_invalid and cb_remove_invalid_black:
            cb_remove_invalid.toggled.connect(
                lambda checked: cb_remove_invalid_black.setChecked(False) if checked else None
            )
            cb_remove_invalid_black.toggled.connect(
                lambda checked: cb_remove_invalid.setChecked(False) if checked else None
            )

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
        # 批量处理运行中允许追加文件（历史行为），只挡预检与检测
        if self._is_precheck_running():
            self.view.show_warning_message("⚠️ 正在预检", "请等待当前预检结束后再添加文件。")
            return
        if self._is_detection_running():
            self.view.show_warning_message("⚠️ 正在检测", "请等待当前检测结束后再添加文件。")
            return

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

        # 智能算法：获取这一次批量拖入文件的公共根路径（跨盘符时降级为绝对路径树）
        common_base = common_base_dir(to_add)

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
            file_node.setForeground(2, QColor(active_palette().text_muted))

            # 将创建的文件节点加入字典中进行状态管理
            self.file_nodes[path] = file_node

        self._mark_precheck_stale()
        self.view.update_counters_ui(len(self.loaded_files))

    def add_folder(self):
        folder_path = QFileDialog.getExistingDirectory(self.view, "选择包含 PDF 的文件夹")
        if folder_path:
            self.add_files([folder_path])

    def clear_list(self):
        if not self.ensure_idle("清空待处理队列"):
            return

        if not self.loaded_files:
            return

        if self.view.show_confirm_message("🗑️ 确认清空", "您确定要清空待处理文件树吗？"):
            self.loaded_files.clear()
            self.folder_nodes.clear()
            self.file_nodes.clear()
            self.view.clear_tree_ui()
            self._mark_precheck_stale()
            self.view.update_counters_ui(0)
            self.process_logs = ""
            self.process_log_rows = []
            self.process_log_starts = {}

    def start_processing(self, processing_files=None, retry_failed=False):
        if self.worker and self.worker.isRunning():
            self.worker.request_stop()
            self.view.btn_start.setEnabled(False)
            self.view.btn_start.setText("正在停止...")
            self.view.btn_skip_current.setEnabled(False)
            return

        if self._is_precheck_running():
            self.view.show_warning_message("⚠️ 正在预检", "请等待当前预检完成后再开始处理。")
            return
        if self._is_detection_running():
            self.view.show_warning_message("⚠️ 正在检测", "请等待当前检测完成后再开始处理。")
            return

        if not self.loaded_files:
            self.view.show_warning_message("⚠️ 警告", "请至少添加一个 PDF 文件！")
            return

        if processing_files is None:
            processing_files = list(self.loaded_files)
        else:
            processing_files = list(processing_files)

        if not processing_files:
            self.view.show_warning_message("⚠️ 警告", "没有可处理的 PDF 文件！")
            return

        selected_options = self.view.get_selected_options()
        if not selected_options:
            self.view.show_warning_message("⚠️ 警告", "请至少在右侧勾选一个处理规则！")
            return

        processing_files = self._prompt_skip_signed_files(processing_files)
        if processing_files is None:
            return
        if not processing_files:
            self.view.show_warning_message("⚠️ 警告", "已跳过全部已签名文件，没有可处理的 PDF 文件！")
            return

        if "filename_ectd_format" in selected_options:
            rename_pairs, collisions = _collect_ectd_rename_plan(processing_files)

            if collisions:
                details = "\n".join([
                    f"{name}: {len(paths)} files"
                    for name, paths in sorted(collisions.items())
                ])
                self.view.show_error_message(
                    "Filename collision",
                    f"eCTD formatting would generate duplicate output names; processing was stopped:\n{details}",
                )
                return

            if rename_pairs:
                details = "\n".join([
                    f"{idx:>2}. {old}\n    -> {new}"
                    for idx, (old, new) in enumerate(rename_pairs, start=1)
                ])
                msg = (
                    "已启用【eCTD 文件名合规格式化】。\n"
                    "以下文件在输出时将被重命名：\n\n"
                    f"{details}\n\n"
                    "确认后继续处理。"
                )
                if not self.view.show_confirm_message("📝 确认文件名格式化", msg):
                    return

        overwrite_cb = self.view.all_checkboxes.get("覆盖原始文件 (不推荐)")
        overwrite_original = overwrite_cb.isChecked() if overwrite_cb else False
        processing_mode = self.view.get_processing_mode()
        max_workers = self._get_processing_worker_count()
        self.processing_parallel_mode = max_workers > 1
        self.processing_worker_count = max_workers
        out_dir = ""
        common_base = ""

        if overwrite_original:
            if not self.view.show_confirm_message("⚠️ 危险操作确认",
                                                  "您勾选了【覆盖原始文件】。\n此操作不可逆，强烈建议您在操作前备份文件！\n\n是否继续？"):
                return
        else:
            default_output_dir = self.view.settings_dialog.default_output_edit.text().strip()
            start_dir = default_output_dir if default_output_dir and os.path.isdir(default_output_dir) else os.path.expanduser("~")

            user_selected_dir = QFileDialog.getExistingDirectory(
                self.view,
                "选择输出文件保存的根目录",
                start_dir
            )
            if not user_selected_dir:
                return

            out_dir = os.path.join(user_selected_dir, "RATools_Output")
            self.last_output_dir = out_dir

            common_base = common_base_dir(processing_files)

        self.view.btn_start.setEnabled(True)
        self.view.btn_start.setText("■ 停止处理")
        self.view.btn_start.setProperty("stopMode", True)
        self.view.btn_retry_failed.setEnabled(False)
        self.view.btn_retry_failed.hide()
        self.view.btn_process_precheck_suggested.setEnabled(False)
        self.view.btn_process_precheck_suggested.hide()
        self.view.btn_precheck.setEnabled(False)
        self.view.btn_precheck.hide()
        self.view.btn_skip_current.show()
        self.view.btn_skip_current.setEnabled(True)
        if self.processing_parallel_mode:
            self.view.btn_skip_current.setText("⏹ 终止选中文件")
            self.view.btn_skip_current.setToolTip("选中队列中正在处理的 PDF 后，只终止该文件")
        else:
            self.view.btn_skip_current.setText("⏭ 跳过当前文件")
            self.view.btn_skip_current.setToolTip("跳过当前正在处理的文件")
        self.view.style().unpolish(self.view.btn_start)
        self.view.style().polish(self.view.btn_start)

        self.processing_started_at = datetime.now()
        self.processing_files = processing_files
        self.processing_total = len(processing_files)
        self.processing_done = 0
        self.processing_done_paths = set()
        self.batch_result_counts = {"success": 0, "failure": 0, "skip": 0}
        if retry_failed:
            self.process_logs += f"\n{'=' * 56}\n仅处理失败项开始：{len(processing_files)} 个文件\n处理模式：{self.view.get_processing_mode_label()}\n{'=' * 56}\n"
        elif processing_files != self.loaded_files:
            self.process_logs += f"\n{'=' * 56}\n仅处理建议文件开始：{len(processing_files)} 个文件\n处理模式：{self.view.get_processing_mode_label()}\n{'=' * 56}\n"
        else:
            self.last_failed_files = []
            self.view.btn_retry_failed.setProperty("hasFailedItems", False)
            self.view.btn_retry_failed.hide()
            self.process_logs += f"\n{'=' * 56}\n批量处理开始：{len(processing_files)} 个文件\n处理模式：{self.view.get_processing_mode_label()}\n{'=' * 56}\n"
        self.processing_current_file = ""
        self._last_processing_hint = ""
        self._refresh_processing_hint()
        self.processing_timer.start()

        self.worker = ProcessWorker(
            processing_files,
            selected_options,
            out_dir,
            common_base,
            overwrite_original,
            max_workers=max_workers,
            processing_mode=processing_mode,
        )
        self.worker.progress.connect(self.update_progress)
        self.worker.structured_progress.connect(self.record_process_log_event)
        self.worker.finished_all.connect(self.processing_finished)
        self.worker.error.connect(self.processing_error)
        self.worker.start()

    def _prompt_skip_signed_files(self, processing_files):
        """处理前检测已签名文件，并询问用户如何处理。

        返回值：
            - 文件列表：应当继续处理的文件（可能已剔除已签名文件）
            - None：用户选择取消，调用方应中止处理
        """
        signed_files = [
            path for path in processing_files
            if pdf_inspect.pdf_has_signature(path)
        ]
        if not signed_files:
            return processing_files

        action = self.view.show_signed_files_prompt(signed_files)
        if action == "cancel":
            return None
        if action == "skip":
            signed_set = set(signed_files)
            kept = [path for path in processing_files if path not in signed_set]
            self.process_logs += (
                f"\n[提示] 检测到 {len(signed_files)} 个已签名文件，已按用户选择跳过。\n"
            )
            return kept
        # process_all
        self.process_logs += (
            f"\n[提示] 检测到 {len(signed_files)} 个已签名文件，用户选择仍然处理全部，"
            "原有数字签名将失效。\n"
        )
        return processing_files

    def start_retry_failed_processing(self):
        retry_files = [path for path in self.last_failed_files if path in self.loaded_files]
        if not retry_files:
            self.view.show_warning_message("⚠️ 无失败项", "当前没有可重新处理的失败文件。")
            return
        self.start_processing(processing_files=retry_files, retry_failed=True)

    def _is_detection_running(self):
        return self.detection.is_running()

    def _recolor_tree_statuses(self, palette=None):
        """按当前主题重刷文件树"当前状态"列前景色（主题切换时调用）。"""
        palette = palette or active_palette()
        waiting_color = QColor(palette.text_muted)
        for node in self.file_nodes.values():
            status_text = node.text(2)
            if not status_text or status_text == "等待处理":
                node.setForeground(2, waiting_color)
            else:
                node.setForeground(2, QColor(tree_status_color(palette, status_semantic(status_text))))

    def update_progress(self, file_path, status_text, log_msg):
        # 信号直接携带 file_path，各类 worker 无需再依赖"当前列表 + 行索引"的隐式路由
        if not file_path:
            return

        color = QColor(tree_status_color(active_palette(), status_semantic(status_text)))

        # 查字典，直接更新树节点UI
        if file_path in self.file_nodes:
            node = self.file_nodes[file_path]
            node.setText(2, status_text)
            node.setToolTip(2, status_text)
            node.setForeground(2, color)

        if status_text in ["处理完成", "处理失败", "已跳过"] and file_path not in self.processing_done_paths:
            self.processing_done_paths.add(file_path)
            self.processing_done = len(self.processing_done_paths)
            if status_text == "处理完成":
                self.batch_result_counts["success"] += 1
            elif status_text == "处理失败":
                self.batch_result_counts["failure"] += 1
            elif status_text == "已跳过":
                self.batch_result_counts["skip"] += 1

        if status_text == "处理失败" and file_path not in self.last_failed_files:
            self.last_failed_files.append(file_path)
            self.view.btn_retry_failed.setProperty("hasFailedItems", True)
            self.view.refresh_selection_summary()
        elif status_text == "处理完成" and file_path in self.last_failed_files:
            self.last_failed_files.remove(file_path)
            self.view.btn_retry_failed.setProperty("hasFailedItems", bool(self.last_failed_files))
            self.view.refresh_selection_summary()

        if status_text == "正在处理..." and file_path:
            self.processing_current_file = os.path.basename(file_path)
        elif status_text in ["处理完成", "处理失败", "已停止", "已跳过"]:
            self.processing_current_file = ""

        if log_msg:
            self.process_logs += f"{log_msg}\n"

        self._refresh_processing_hint(status_text=status_text, file_path=file_path)

    def record_process_log_event(self, event):
        if not isinstance(event, dict):
            return
        file_path = event.get("file_path", "")
        if not file_path:
            return
        if event.get("type") == "start":
            self.process_log_starts[file_path] = dict(event)
            return
        row = _structured_log_row_from_event(event, self.process_log_starts)
        if row:
            self.process_log_rows.append(row)
            self.process_log_starts.pop(file_path, None)

    def _build_batch_result_summary(self, worker_summary):
        total = self.processing_total or sum(self.batch_result_counts.values())
        success = self.batch_result_counts.get("success", 0)
        failure = self.batch_result_counts.get("failure", 0)
        skip = self.batch_result_counts.get("skip", 0)
        elapsed = 0
        if self.processing_started_at:
            elapsed = int((datetime.now() - self.processing_started_at).total_seconds())

        lines = [
            "批次结果摘要",
            f"总数：{total}",
            f"成功：{success}",
            f"失败：{failure}",
            f"跳过：{skip}",
            f"总耗时：{elapsed}s",
        ]
        if self.last_failed_files:
            lines.append(f"失败项：{len(self.last_failed_files)} 个，可使用“重试失败项”重试")
        else:
            lines.append("失败项：0 个")
        if worker_summary:
            lines.extend(["", worker_summary])
        return "\n".join(lines)

    def _reset_processing_ui(self):
        """批处理结束/异常后的按钮与状态复位（finished 与 error 共用）。"""
        self.processing_timer.stop()
        self.view.processing_hint_label.setText("")
        self.view.btn_start.setEnabled(True)
        self.view.btn_start.setText("▶ 开始处理")
        self.view.btn_start.setProperty("stopMode", False)
        self.view.btn_precheck.show()
        self.view.btn_retry_failed.setProperty("hasFailedItems", bool(self.last_failed_files))
        self.view.btn_skip_current.setEnabled(False)
        self.view.btn_skip_current.hide()
        self.view.btn_skip_current.setText("⏭ 跳过当前文件")
        self.view.btn_skip_current.setToolTip("")
        self.view.style().unpolish(self.view.btn_start)
        self.view.style().polish(self.view.btn_start)
        self.view.refresh_selection_summary()
        self.processing_started_at = None
        self.processing_total = 0
        self.processing_done = 0
        self.processing_done_paths.clear()
        self.processing_files = []
        self.processing_current_file = ""
        self.processing_parallel_mode = False
        self.processing_worker_count = 1
        self._last_processing_hint = ""

    def processing_finished(self, summary):
        batch_summary = self._build_batch_result_summary(summary)
        self.process_logs += f"\n{'=' * 56}\n批量处理结束\n{batch_summary}\n{'=' * 56}\n"
        # 处理会改变文件状态，上一轮预检结果随之失效
        self.precheck.precheck_result_current = False
        self.view.btn_precheck.setProperty("precheckResultCurrent", False)
        self.view.btn_apply_precheck.setProperty("hasPrecheckSuggestions", False)
        self.view.btn_apply_precheck.hide()
        self.view.btn_process_precheck_suggested.setProperty("hasPrecheckSuggestedFiles", False)
        self.view.btn_process_precheck_suggested.hide()
        self._reset_processing_ui()

        if "任务已停止" in summary:
            self.view.show_info_message("⏹️ 已停止", batch_summary)
        else:
            self.view.show_success_message("✅ 处理完成", batch_summary)

        auto_open_cb = self.view.all_checkboxes.get("处理完成后自动打开输出文件夹")
        if auto_open_cb and auto_open_cb.isChecked() and self.loaded_files:
            overwrite_cb = self.view.all_checkboxes.get("覆盖原始文件 (不推荐)")
            if overwrite_cb and not overwrite_cb.isChecked():
                if hasattr(self, 'last_output_dir') and self.last_output_dir and os.path.exists(self.last_output_dir):
                    self._open_directory(self.last_output_dir)

    def processing_error(self, error_msg):
        self.process_logs += f"\n{'!' * 56}\n[致命错误] {error_msg}\n{'!' * 56}\n"
        self._reset_processing_ui()
        self.view.show_error_message("❌ 处理异常", f"处理过程中发生错误：\n{error_msg}")

    def _refresh_processing_hint(self, status_text="", file_path=""):
        if not self.processing_started_at:
            self.view.processing_hint_label.setText("")
            self._last_processing_hint = ""
            return

        elapsed = int((datetime.now() - self.processing_started_at).total_seconds())
        total = max(self.processing_total, 1)
        done = min(self.processing_done, total)
        percent = int(done * 100 / total)
        hint = f"处理中 {elapsed}s · {done}/{total} · {percent}%"

        current_name = self.processing_current_file
        if self.processing_parallel_mode:
            current_name = f"并行 {self.processing_worker_count} 个任务"
        elif status_text == "正在处理..." and file_path:
            current_name = os.path.basename(file_path)
        if current_name:
            hint += f" · {current_name}"

        if hint != self._last_processing_hint:
            self.view.processing_hint_label.setText(hint)
            self._last_processing_hint = hint

    def skip_current_file(self):
        if not (self.worker and self.worker.isRunning()):
            return

        if not self.processing_parallel_mode:
            self.worker.request_skip_current()
            return

        selected_items = self.view.tree.selectedItems()
        if not selected_items:
            self.view.show_warning_message("⚠️ 未选择文件", "请先在左侧队列选中一个正在处理的 PDF。")
            return

        requested = 0
        for item in selected_items:
            file_path = item.text(1)
            status_text = item.text(2)
            if file_path and status_text == "正在处理...":
                self.worker.request_skip_file(file_path)
                requested += 1

        if requested:
            return

        self.view.show_warning_message("⚠️ 未选择正在处理的文件", "请选择状态为“正在处理...”的 PDF 后再终止。")

    def _open_directory(self, dir_path):
        try:
            system_shell.open_directory(dir_path)
        except Exception as e:
            self.process_logs += f"\n[警告] 自动打开文件夹失败：{str(e)}\n"
