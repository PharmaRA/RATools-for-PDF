import csv
import os
import webbrowser
from datetime import date, datetime
from pathlib import Path

from PySide6.QtCore import QCoreApplication, QObject, Qt, QTimer
from PySide6.QtGui import QColor
from PySide6.QtWidgets import QFileDialog, QMenu, QTreeWidgetItem

from ratools_pdf.common.status import status_semantic
from ratools_pdf.config.features import ENABLE_UPDATE_CHECK
from ratools_pdf.controllers.io_actions import (
    _build_io_preview_rows,
    _collect_ectd_rename_plan,
    _io_action_metadata,
    _normalize_io_action_types,
)
from ratools_pdf.controllers.log_export import (
    _select_log_rows_for_export,
    _structured_log_row_from_event,
)
from ratools_pdf.controllers.workers import (
    DetectionWorker,
    IOActionWorker,
    PreCheckWorker,
    ProcessWorker,
    UpdateCheckWorker,
)
from ratools_pdf.pdf.processor import PDFProcessor
from ratools_pdf.services import system_shell
from ratools_pdf.ui.dialogs import IODataWizardDialog, LogDialog
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
        self.last_precheck_suggested_files = []
        self.batch_result_counts = {"success": 0, "failure": 0, "skip": 0}
        self.last_precheck_results = []
        self.precheck_result_current = False
        self.processing_timer = QTimer(self)
        self.processing_timer.setInterval(1000)
        self.processing_timer.timeout.connect(self._refresh_processing_hint)

        # 建立缓存字典，以便快速在文件树中更新和查找节点
        self.folder_nodes = {}
        self.file_nodes = {}

        self.setup_connections()
        self.worker = None
        self.precheck_worker = None
        self.precheck_files = []
        self.detection_worker = None
        self.detection_files = []
        self.last_detection_results = []
        self.update_worker = None
        app = QCoreApplication.instance()
        if app:
            app.aboutToQuit.connect(self.shutdown_update_worker)

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
        self.view.btn_apply_precheck.clicked.connect(self.apply_precheck_suggestions)
        self.view.btn_process_precheck_suggested.clicked.connect(self.start_precheck_suggested_processing)
        self.view.btn_precheck.clicked.connect(self.start_precheck)
        self.view.btn_start.clicked.connect(lambda _checked=False: self.start_processing())
        self.view.btn_log.clicked.connect(self.show_log_dialog)
        btn_embed_missing_fonts = getattr(self.view, "btn_embed_missing_fonts", None)
        if btn_embed_missing_fonts is not None:
            btn_embed_missing_fonts.clicked.connect(self.open_selected_files_in_acrobat_for_font_embedding)
        if ENABLE_UPDATE_CHECK:
            self.view.btn_top_about.clicked.connect(self._wire_about_dialog_updates)

        self.view.btn_bookmark_io_wizard.clicked.connect(lambda: self.handle_io_wizard("bookmarks"))
        self.view.btn_link_io_wizard.clicked.connect(lambda: self.handle_io_wizard("links"))
        self.view.btn_detect_annotations.clicked.connect(lambda: self.start_detection("annotations"))
        self.view.btn_detect_broken_refs.clicked.connect(lambda: self.start_detection("broken_refs"))

        self.setup_exclusive_options()

        # 绑定树形图的右键菜单请求事件
        self.view.tree.customContextMenuRequested.connect(self.show_tree_context_menu)
        # 绑定树形图双击事件
        self.view.tree.itemDoubleClicked.connect(self.on_item_double_clicked)

    def _is_precheck_running(self):
        precheck_worker = getattr(self, "precheck_worker", None)
        return bool(precheck_worker and precheck_worker.isRunning())

    def _mark_precheck_stale(self):
        self.precheck_result_current = False
        self.view.btn_precheck.setProperty("precheckResultCurrent", False)
        self.view.btn_precheck.show()

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

    def _record_precheck_result(self, row):
        self.last_precheck_results.append(dict(row))
        suggestion_ids = str(row.get("suggestion_ids", "") or "")
        if suggestion_ids.strip():
            self.view.btn_apply_precheck.setProperty("hasPrecheckSuggestions", True)
        if row.get("status") == "建议处理" and row.get("file_path") and suggestion_ids.strip():
            file_path = row.get("file_path")
            if file_path not in self.last_precheck_suggested_files:
                self.last_precheck_suggested_files.append(file_path)
            self.view.btn_process_precheck_suggested.setProperty("hasPrecheckSuggestedFiles", True)
        self.view.refresh_selection_summary()

    def apply_precheck_suggestions(self):
        suggested_options = []
        seen = set()
        for row in self.last_precheck_results:
            raw_ids = str(row.get("suggestion_ids", "") or "")
            for option_id in [item.strip() for item in raw_ids.split(",") if item.strip()]:
                if option_id in seen:
                    continue
                if option_id in self.view.all_checkboxes:
                    suggested_options.append(option_id)
                    seen.add(option_id)

        if not suggested_options:
            self.view.show_warning_message("⚠️ 无建议项", "最近一次预检没有可自动应用的建议处理项。")
            return

        self.view.is_applying_preset = True
        try:
            for option_id in suggested_options:
                self.view.all_checkboxes[option_id].setChecked(True)
        finally:
            self.view.is_applying_preset = False

        self.view.active_preset_key = None
        self.view._set_preset_button_state(None)
        self.view.custom_selection_before_preset = set(self.view.get_selected_options())
        self.view.persist_all_settings()
        self.view.refresh_selection_summary()
        self.view.show_success_message("✅ 已应用", f"已自动勾选 {len(suggested_options)} 条预检建议规则。")

    def _wire_about_dialog_updates(self):
        if not ENABLE_UPDATE_CHECK:
            return
        dialog = getattr(self.view, "about_dialog", None)
        if not dialog or getattr(dialog, "_update_buttons_wired", False):
            return

        dialog.btn_check_updates.clicked.connect(self.check_updates_manually)
        dialog.btn_open_release.clicked.connect(self.open_release_url)
        dialog._update_buttons_wired = True

    def check_updates_manually(self):
        if not ENABLE_UPDATE_CHECK:
            return
        self._wire_about_dialog_updates()
        dialog = getattr(self.view, "about_dialog", None)
        started = self._start_update_check(silent=False)
        if dialog and started:
            dialog.set_update_checking()
        elif dialog:
            dialog.set_update_result("已有更新检查正在进行，请稍后再试。")

    def check_updates_on_startup(self):
        if not ENABLE_UPDATE_CHECK:
            return
        settings = getattr(self.view, "app_settings", None)
        if not settings:
            return

        today = date.today().isoformat()
        if settings.value("Update/LastSilentCheckDate") == today:
            return

        if self._start_update_check(silent=True):
            settings.setValue("Update/LastSilentCheckDate", today)

    def _start_update_check(self, silent=False):
        if not ENABLE_UPDATE_CHECK:
            return False
        worker = getattr(self, "update_worker", None)
        if worker and worker.isRunning():
            return False

        worker = UpdateCheckWorker(silent=silent, parent=self)
        worker.finished_check.connect(self._handle_update_result)
        worker.finished.connect(worker.deleteLater)
        worker.finished.connect(lambda checked_worker=worker: self._clear_update_worker(checked_worker))
        self.update_worker = worker
        worker.start()
        return True

    def _clear_update_worker(self, worker):
        if getattr(self, "update_worker", None) is worker:
            self.update_worker = None

    def shutdown_update_worker(self):
        worker = getattr(self, "update_worker", None)
        if worker and worker.isRunning():
            worker.wait(9000)

    def open_release_url(self, release_url=""):
        if not release_url:
            dialog = getattr(self.view, "about_dialog", None)
            release_url = getattr(dialog, "latest_release_url", "") if dialog else ""
        if not release_url:
            return

        try:
            opened = webbrowser.open(release_url)
        except Exception as exc:
            self.view.show_error_message("打开失败", f"无法打开发布页：{exc}")
            return

        if not opened:
            self.view.show_error_message("打开失败", "无法打开发布页：浏览器拒绝打开链接")

    @staticmethod
    def _find_acrobat_executable():
        return system_shell.find_acrobat_executable()

    @staticmethod
    def _open_pdf_for_manual_font_embedding(pdf_path, acrobat_path=None):
        system_shell.open_pdf_in_acrobat_or_default(pdf_path, acrobat_path)

    def _selected_pdf_paths_for_manual_font_embedding(self):
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

    def _open_pdf_paths_in_acrobat_for_font_embedding(self, pdf_paths):
        acrobat_path = self._find_acrobat_executable()
        opened = []
        failures = []
        for pdf_path in pdf_paths:
            try:
                self._open_pdf_for_manual_font_embedding(pdf_path, acrobat_path=acrobat_path)
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

    def open_selected_files_in_acrobat_for_font_embedding(self):
        pdf_paths = self._selected_pdf_paths_for_manual_font_embedding()
        if not pdf_paths:
            self.view.show_warning_message("⚠️ 未选择 PDF", "请先在左侧待处理队列中选中需要嵌入缺失字体的 PDF 文件。")
            return

        self.view.show_manual_font_embedding_dialog(
            pdf_paths,
            lambda paths=tuple(pdf_paths): self._open_pdf_paths_in_acrobat_for_font_embedding(list(paths)),
        )

    def _handle_silent_update_result(self, result):
        if not result.ok or not result.has_update or not result.is_major or not result.latest_release:
            return

        settings = getattr(self.view, "app_settings", None)
        if not settings:
            return

        release = result.latest_release
        if settings.value("Update/IgnoredVersion", "") == release.version_text:
            return

        today = date.today().isoformat()
        prompt_marker = f"{today}:{release.version_text}"
        if settings.value("Update/LastPromptedVersion", "") == prompt_marker:
            return

        settings.setValue("Update/LastPromptedVersion", prompt_marker)
        action = self.view.show_major_update_prompt(result.current_version, release)

        if action == "open":
            self.open_release_url(release.html_url)
        elif action == "ignore":
            settings.setValue("Update/IgnoredVersion", release.version_text)

    def _handle_update_result(self, result, silent):
        if silent:
            handler = getattr(self, "_handle_silent_update_result", None)
            if handler:
                handler(result)
            return

        dialog = getattr(self.view, "about_dialog", None)
        if not dialog:
            return

        if not result.ok:
            dialog.set_update_result(f"检查更新失败：{result.error}")
            return

        if not result.has_update or not result.latest_release:
            dialog.set_update_result(f"当前已是最新版本：{result.current_version}")
            return

        release = result.latest_release
        published_at = release.published_at or "未知"
        message = (
            f"发现新版本：{release.version_text}\n"
            f"当前版本：{result.current_version}\n"
            f"发布标题：{release.title}\n"
            f"发布时间：{published_at}"
        )
        dialog.set_update_result(message, release.html_url)

    # ================= 核心：右键菜单生成与分发 =================
    def show_tree_context_menu(self, pos):
        selected_items = self.view.tree.selectedItems()
        if not selected_items:
            return

        # 菜单样式由应用级中央 QSS（theme.py 的 QMenu 段）提供，明暗主题自动适配
        menu = QMenu(self.view.tree)

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
            info_text = self._build_pdf_detail_text(path)
            self.view.show_info_message("📄 文件详细信息", info_text)

        except Exception as e:
            self.view.show_error_message("❌ 读取失败", f"获取文件信息时发生异常：\n{str(e)}")

    def _build_pdf_detail_text(self, path):
        stat = os.stat(path)
        created_time = datetime.fromtimestamp(stat.st_ctime).strftime('%Y-%m-%d %H:%M:%S')
        modified_time = datetime.fromtimestamp(stat.st_mtime).strftime('%Y-%m-%d %H:%M:%S')
        base_name = os.path.basename(path)
        size_kb = stat.st_size / 1024
        size_text = f"{size_kb / 1024:.2f} MB" if size_kb > 1024 else f"{size_kb:.2f} KB"

        details = [
            "📌 基础信息",
            f"文件名：{base_name}",
            f"路径：{path}",
            f"大小：{size_text}",
            f"创建时间：{created_time}",
            f"修改时间：{modified_time}",
        ]

        if not os.path.isfile(path):
            details.append("类型：文件夹")
            return "\n".join(details)

        if not path.lower().endswith('.pdf'):
            details.append("类型：普通文件")
            return "\n".join(details)

        doc = None
        try:
            import fitz
            doc = fitz.open(path)
            pdf_version = PDFProcessor._read_pdf_header_version(path) or "未知"
            linearized = "是" if PDFProcessor._is_pdf_linearized(path) else "否"
            restrictions = "是" if PDFProcessor._qpdf_reports_restrictions(path) else "否"

            details.extend([
                "",
                "📄 PDF信息",
                f"页数：{doc.page_count} 页",
                f"PDF版本：{pdf_version}",
                f"是否线性化：{linearized}",
                f"是否需要打开密码：{'是' if doc.needs_pass else '否'}",
                f"是否存在权限限制：{restrictions}",
            ])

            if doc.needs_pass:
                details.extend(["", "🛠️ 建议处理：该文件需要打开密码，无法展开内部结构预检"])
                return "\n".join(details)

            toc_count = len(doc.get_toc(simple=False))
            link_count = 0
            uri_link_count = 0
            annotation_count = 0
            for page in doc:
                links = page.get_links()
                link_count += len(links)
                uri_link_count += sum(1 for link in links if link.get("kind") == fitz.LINK_URI)
                annotation_count += sum(1 for _annot in (page.annots() or []))

            catalog_xref = doc.pdf_catalog()
            has_tags = (
                PDFProcessor._catalog_key_is_present(doc, catalog_xref, "StructTreeRoot")
                or PDFProcessor._catalog_key_is_present(doc, catalog_xref, "MarkInfo")
            )
            metadata_values = [
                value
                for key, value in (doc.metadata or {}).items()
                if key not in ["format", "encryption"] and str(value or "").strip()
            ]

            details.extend([
                "",
                "🧩 结构信息",
                f"书签数量：{toc_count}",
                f"页面链接数量：{link_count}",
                f"外部URI链接数量：{uri_link_count}",
                f"批注数量：{annotation_count}",
                f"内嵌附件数量：{doc.embfile_count()}",
                f"是否包含结构化标签：{'是' if has_tags else '否'}",
                f"是否包含元数据：{'是' if metadata_values else '否'}",
            ])

            report = PDFProcessor.build_precheck_report(path)
            suggestions = []
            for item in report.get("suggestions", {}).values():
                title = item.get("title", "")
                if not title:
                    continue
                reason = item.get("reason", "")
                if item.get("report_only") and reason:
                    suggestions.append(f"{title}：{reason}")
                else:
                    suggestions.append(title)
            details.append("")
            if suggestions:
                details.append("🛠️ 建议处理：")
                details.extend(f"- {title}" for title in suggestions)
            else:
                details.append("🛠️ 建议处理：暂无明显建议项")

            return "\n".join(details)
        except Exception:
            details.extend(["", "⚠️ 提示：无法解析 PDF 内部结构，文件可能已损坏"])
            return "\n".join(details)
        finally:
            if doc is not None:
                try:
                    doc.close()
                except Exception:
                    pass

    def remove_selected_items(self, selected_items):
        """
        处理树节点的移除操作。
        逻辑：递归收集所有选中的文件路径 -> 更新后台数据 -> 删除 UI 节点 -> 自动清理空文件夹
        """
        if self.worker and self.worker.isRunning():
            self.view.show_warning_message("⚠️ 正在处理中", "请等待当前批量处理结束后再移除队列项。")
            return
        if self._is_precheck_running():
            self.view.show_warning_message("⚠️ 正在预检", "请等待当前预检结束后再移除队列项。")
            return
        if self._is_detection_running():
            self.view.show_warning_message("⚠️ 正在检测", "请等待当前检测结束后再移除队列项。")
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
        if self.worker and self.worker.isRunning():
            self.view.show_warning_message("⚠️ 正在处理中", "请等待当前批量处理结束后再清空待处理队列。")
            return
        if self._is_precheck_running():
            self.view.show_warning_message("⚠️ 正在预检", "请等待当前预检结束后再清空待处理队列。")
            return
        if self._is_detection_running():
            self.view.show_warning_message("⚠️ 正在检测", "请等待当前检测结束后再清空待处理队列。")
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

            try:
                dirs = [os.path.dirname(os.path.abspath(f)) for f in processing_files]
                common_base = os.path.commonpath(dirs)
            except ValueError:
                common_base = ""

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
            if PDFProcessor._pdf_has_signature(path)
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

    def start_precheck_suggested_processing(self):
        suggested_files = [path for path in self.last_precheck_suggested_files if path in self.loaded_files]
        if not suggested_files:
            self.view.show_warning_message("⚠️ 无建议文件", "最近一次预检没有可处理的建议文件。")
            return
        self.start_processing(processing_files=suggested_files)

    def start_precheck(self):
        if self.worker and self.worker.isRunning():
            self.view.show_warning_message("⚠️ 正在处理", "批量处理进行中，无法执行预检。")
            return
        if self._is_detection_running():
            self.view.show_warning_message("⚠️ 正在检测", "检测进行中，请稍候再执行预检。")
            return
        if self._is_precheck_running():
            return
        if not self.loaded_files:
            self.view.show_warning_message("⚠️ 警告", "请至少添加一个 PDF 文件！")
            return

        self.precheck_files = list(self.loaded_files)
        self.last_precheck_results = []
        self.last_precheck_suggested_files = []
        self.precheck_result_current = False
        self.view.btn_precheck.setProperty("precheckResultCurrent", False)
        self.view.btn_apply_precheck.setProperty("hasPrecheckSuggestions", False)
        self.view.btn_apply_precheck.hide()
        self.view.btn_process_precheck_suggested.setProperty("hasPrecheckSuggestedFiles", False)
        self.view.btn_process_precheck_suggested.hide()
        self.process_logs += f"\n{'=' * 56}\n批量预检开始\n{'=' * 56}\n"
        self.view.btn_precheck.setEnabled(False)
        self.view.btn_precheck.setText("预检中...")
        self.view.btn_precheck.setProperty("precheckMode", True)
        self.view.btn_start.setEnabled(False)

        self.precheck_worker = PreCheckWorker(self.precheck_files)
        self.precheck_worker.progress.connect(self.update_progress)
        self.precheck_worker.result_ready.connect(self._record_precheck_result)
        self.precheck_worker.finished_precheck.connect(self.precheck_finished)
        self.precheck_worker.error_precheck.connect(self.precheck_error)
        self.precheck_worker.finished.connect(self.precheck_worker.deleteLater)
        self.precheck_worker.start()

    def _is_detection_running(self):
        detection_worker = getattr(self, "detection_worker", None)
        return bool(detection_worker and detection_worker.isRunning())

    DETECTION_KIND_MAP = {
        "annotations": "annotation",
        "broken_refs": "broken_reference",
    }

    DETECTION_KIND_TITLES = {
        "annotation": "批注检测",
        "broken_reference": "失效引用/链接文本检测",
    }

    def start_detection(self, ui_kind):
        detection_kind = self.DETECTION_KIND_MAP.get(ui_kind, "annotation")
        kind_title = self.DETECTION_KIND_TITLES[detection_kind]

        if self.worker and self.worker.isRunning():
            self.view.show_warning_message("⚠️ 正在处理", f"批量处理进行中，无法执行{kind_title}。")
            return
        if self._is_precheck_running():
            self.view.show_warning_message("⚠️ 正在预检", f"预检进行中，无法执行{kind_title}。")
            return
        if self._is_detection_running():
            self.view.show_warning_message("⚠️ 正在检测", "已有检测任务进行中，请稍候。")
            return
        if not self.loaded_files:
            self.view.show_warning_message("⚠️ 警告", "请至少添加一个 PDF 文件！")
            return

        self.detection_files = list(self.loaded_files)
        self.last_detection_results = []
        self.process_logs += f"\n{'=' * 56}\n{kind_title}开始\n{'=' * 56}\n"

        self.view.btn_detect_annotations.setEnabled(False)
        self.view.btn_detect_broken_refs.setEnabled(False)

        self.detection_worker = DetectionWorker(self.detection_files, detection_kind)
        self.detection_worker.progress.connect(self.update_progress)
        self.detection_worker.result_ready.connect(self._record_detection_result)
        self.detection_worker.finished_detection.connect(
            lambda summary, results, title=kind_title: self.detection_finished(summary, results, title)
        )
        self.detection_worker.error_detection.connect(
            lambda error_msg, title=kind_title: self.detection_error(error_msg, title)
        )
        self.detection_worker.finished.connect(self.detection_worker.deleteLater)
        self.detection_worker.start()

    def _record_detection_result(self, row):
        self.last_detection_results.append(dict(row))

    def _restore_detection_buttons(self):
        self.view.btn_detect_annotations.setEnabled(True)
        self.view.btn_detect_broken_refs.setEnabled(True)

    def detection_finished(self, summary, results, kind_title):
        self.process_logs += f"\n{'=' * 56}\n{kind_title}结束\n{summary}\n{'=' * 56}\n"
        self._restore_detection_buttons()
        self.detection_files = []
        self.detection_worker = None

        hit_rows = [row for row in results if row.get("status") == "发现问题"]
        failed_rows = [row for row in results if row.get("status") == "检测失败"]

        message_lines = [summary, ""]
        if hit_rows:
            message_lines.append("发现问题的文件：")
            for row in hit_rows:
                detail = row.get("summary", "") or ""
                if row.get("details"):
                    detail = f"{detail}（{row['details']}）" if detail else row["details"]
                message_lines.append(f"• {row.get('file_name', '')}：{detail}")
        else:
            message_lines.append("未在任何文件中发现相关内容。")
        if failed_rows:
            message_lines.append("")
            message_lines.append("检测失败的文件：")
            for row in failed_rows:
                message_lines.append(f"• {row.get('file_name', '')}：{row.get('error', '')}")

        self.view.show_info_message(f"🔍 {kind_title}完成", "\n".join(message_lines))

    def detection_error(self, error_msg, kind_title):
        self.process_logs += f"\n{'!' * 56}\n[{kind_title}错误] {error_msg}\n{'!' * 56}\n"
        self._restore_detection_buttons()
        self.detection_files = []
        self.detection_worker = None
        self.view.show_error_message(f"❌ {kind_title}失败", f"检测过程中发生错误：\n{error_msg}")

    def _get_loaded_files_common_base(self):
        try:
            dirs = [os.path.dirname(os.path.abspath(f)) for f in self.loaded_files]
            return os.path.commonpath(dirs)
        except ValueError:
            return ""

    def handle_io_wizard(self, default_data_kind):
        if not self.loaded_files:
            self.view.show_warning_message("⚠️ 警告", "请先添加目标 PDF 文件！")
            return

        common_base = self._get_loaded_files_common_base()

        def build_preview(action_type, dir_path):
            return _build_io_preview_rows(self.loaded_files, action_type, dir_path, common_base)

        dialog = IODataWizardDialog(
            data_kind=default_data_kind,
            file_count=len(self.loaded_files),
            preview_callback=build_preview,
            parent=self.view,
        )
        if not dialog.exec():
            return

        self.handle_io_action(
            dialog.get_action_types(),
            dir_path=dialog.get_selected_directory(),
            common_base=common_base,
            confirmed=True,
            link_scope=dialog.get_link_scope(),
            link_mode=dialog.get_link_mode(),
        )

    def handle_io_action(self, action_type, dir_path=None, common_base=None, confirmed=False,
                         link_scope="all", link_mode="overwrite"):
        if not self.loaded_files:
            self.view.show_warning_message("⚠️ 警告", "请先添加目标 PDF 文件！")
            return

        action_types = _normalize_io_action_types(action_type)
        meta = _io_action_metadata(action_types[0])
        is_export = meta["is_export"]
        data_type = "CSV/JSON" if len(action_types) > 1 else meta["data_type"]
        action_name = meta["action_name"]

        if common_base is None:
            common_base = self._get_loaded_files_common_base()

        if dir_path is None:
            dir_path = QFileDialog.getExistingDirectory(self.view, f"请选择 {data_type} 数据{action_name}的目录")
        if not dir_path:
            return

        if confirmed and not is_export:
            rows = _build_io_preview_rows(self.loaded_files, action_types, dir_path, common_base)
            if not any(row["status"] == "已匹配" for row in rows):
                self.view.show_warning_message("⚠️ 未找到数据文件", f"所选目录中没有匹配的 {data_type} 文件。")
                return

        out_dir = None
        if not is_export:
            first_file = Path(self.loaded_files[0])
            out_dir_path = first_file.parent / f"RATools_{action_name}完成"
            out_dir_path.mkdir(exist_ok=True)
            out_dir = str(out_dir_path)

        self.io_worker = IOActionWorker(
            action_types, self.loaded_files, dir_path, out_dir, common_base,
            link_scope=link_scope, link_mode=link_mode,
        )
        self.io_worker.progress.connect(self.update_progress)
        self.io_worker.finished_action.connect(self.on_io_action_finished)
        self.io_worker.error_action.connect(self.on_io_action_error)
        self.io_worker.start()

    def update_progress(self, row_index, status_text, log_msg):
        # 获取与该行对应的精确文件路径，用于树节点的映射更新
        processing_files = (
            self.processing_files
            or self.precheck_files
            or self.detection_files
            or self.loaded_files
        )
        if row_index < 0 or row_index >= len(processing_files):
            return
        file_path = processing_files[row_index]

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

    def precheck_finished(self, summary):
        self.process_logs += f"\n{'=' * 56}\n批量预检结束\n{summary}\n{'=' * 56}\n"
        self.view.btn_precheck.setText("🔎 预检")
        self.view.btn_precheck.setProperty("precheckMode", False)
        self.precheck_result_current = True
        self.view.btn_precheck.setProperty("precheckResultCurrent", True)
        self.view.btn_precheck.hide()
        self.view.btn_start.setEnabled(True)
        self.precheck_files = []
        self.precheck_worker = None
        self.view.refresh_selection_summary()
        self.view.show_info_message("🔎 预检完成", summary)

    def precheck_error(self, error_msg):
        self.process_logs += f"\n{'!' * 56}\n[预检错误] {error_msg}\n{'!' * 56}\n"
        self.view.btn_precheck.setText("🔎 预检")
        self.view.btn_precheck.setProperty("precheckMode", False)
        self.view.btn_start.setEnabled(True)
        self.precheck_files = []
        self.precheck_worker = None
        self.view.refresh_selection_summary()
        self.view.show_error_message("❌ 预检失败", f"预检过程中发生错误：\n{error_msg}")

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
            lines.append(f"失败项：{len(self.last_failed_files)} 个，可使用“仅处理失败项”重试")
        else:
            lines.append("失败项：0 个")
        if worker_summary:
            lines.extend(["", worker_summary])
        return "\n".join(lines)

    def processing_finished(self, summary):
        batch_summary = self._build_batch_result_summary(summary)
        self.process_logs += f"\n{'=' * 56}\n批量处理结束\n{batch_summary}\n{'=' * 56}\n"
        self.processing_timer.stop()
        self.view.processing_hint_label.setText("")
        self.view.btn_start.setEnabled(True)
        self.view.btn_start.setText("▶ 开始批量处理")
        self.view.btn_start.setProperty("stopMode", False)
        self.precheck_result_current = False
        self.view.btn_precheck.setProperty("precheckResultCurrent", False)
        self.view.btn_precheck.show()
        self.view.btn_apply_precheck.setProperty("hasPrecheckSuggestions", False)
        self.view.btn_apply_precheck.hide()
        self.view.btn_process_precheck_suggested.setProperty("hasPrecheckSuggestedFiles", False)
        self.view.btn_process_precheck_suggested.hide()
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
        self.processing_timer.stop()
        self.view.processing_hint_label.setText("")
        self.view.btn_start.setEnabled(True)
        self.view.btn_start.setText("▶ 开始批量处理")
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

    def on_io_action_finished(self, result_msg):
        self.process_logs += f"\n{'-' * 56}\n{result_msg}\n{'-' * 56}\n"
        self.view.show_success_message("✅ 操作成功", result_msg)

    def on_io_action_error(self, error_msg):
        self.process_logs += f"\n{'!' * 56}\n[IO错误] {error_msg}\n{'!' * 56}\n"
        self.view.show_error_message("❌ 操作失败", error_msg)

    def show_log_dialog(self):
        if not hasattr(self, 'log_dialog'):
            self.log_dialog = LogDialog(self.view)
            self.log_dialog.btn_export.clicked.connect(self.export_logs)
            self.log_dialog.btn_export_precheck.clicked.connect(self.export_precheck_results)

        self.log_dialog.set_log_data(
            self.process_logs if self.process_logs else "暂无处理日志...",
            self.process_log_rows,
        )
        self.log_dialog.btn_export_precheck.setEnabled(bool(self.last_precheck_results))
        self.log_dialog.btn_export_precheck.setToolTip(
            "导出最近一次批量预检结果" if self.last_precheck_results else "请先执行一次批量预检"
        )
        self.log_dialog.show()
        self.log_dialog.raise_()
        self.log_dialog.activateWindow()

    def export_logs(self):
        if not self.process_logs:
            self.view.show_warning_message("⚠️ 提示", "目前暂无任何日志可供导出！")
            return

        default_dir = ""
        if hasattr(self, 'last_output_dir') and self.last_output_dir and os.path.isdir(self.last_output_dir):
            default_dir = self.last_output_dir
        elif self.view.settings_dialog.default_output_edit.text().strip() and os.path.isdir(self.view.settings_dialog.default_output_edit.text().strip()):
            default_dir = self.view.settings_dialog.default_output_edit.text().strip()
        elif self.loaded_files:
            try:
                file_dirs = [os.path.dirname(os.path.abspath(f)) for f in self.loaded_files]
                default_dir = os.path.commonpath(file_dirs)
            except ValueError:
                default_dir = os.path.dirname(os.path.abspath(self.loaded_files[0]))

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        default_filename = f"RATools_process_logs_{timestamp}.csv"
        default_path = os.path.join(default_dir, default_filename) if default_dir else default_filename

        file_path, selected_filter = QFileDialog.getSaveFileName(
            self.view,
            "导出处理日志",
            default_path,
            "CSV Summary (*.csv);;Text Files (*.txt);;All Files (*)"
        )
        if file_path:
            try:
                export_csv = file_path.lower().endswith('.csv') or selected_filter.startswith("CSV")

                if export_csv and not file_path.lower().endswith('.csv'):
                    file_path += '.csv'
                if not export_csv and selected_filter.startswith("Text") and not file_path.lower().endswith('.txt'):
                    file_path += '.txt'

                if export_csv:
                    with open(file_path, "w", encoding="utf-8-sig", newline="") as f:
                        writer = csv.DictWriter(f, fieldnames=["time", "file_original", "file_output", "status", "success", "duration_sec", "changes"])
                        writer.writeheader()
                        writer.writerows(_select_log_rows_for_export(self.process_log_rows, self.process_logs))
                else:
                    with open(file_path, "w", encoding="utf-8") as f:
                        f.write(self.process_logs)
                self.view.show_success_message("✅ 导出成功", "处理日志已成功保存！")
            except Exception as e:
                self.view.show_error_message("❌ 导出失败", f"文件保存失败：\n{str(e)}")

    def export_precheck_results(self):
        if not self.last_precheck_results:
            self.view.show_warning_message("⚠️ 提示", "请先执行一次批量预检，再导出预检结果。")
            return

        default_dir = ""
        default_output_dir = self.view.settings_dialog.default_output_edit.text().strip()
        if default_output_dir and os.path.isdir(default_output_dir):
            default_dir = default_output_dir
        elif self.loaded_files:
            try:
                file_dirs = [os.path.dirname(os.path.abspath(f)) for f in self.loaded_files]
                default_dir = os.path.commonpath(file_dirs)
            except ValueError:
                default_dir = os.path.dirname(os.path.abspath(self.loaded_files[0]))

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        default_filename = f"RATools_precheck_results_{timestamp}.csv"
        default_path = os.path.join(default_dir, default_filename) if default_dir else default_filename

        file_path, _selected_filter = QFileDialog.getSaveFileName(
            self.view,
            "导出预检结果",
            default_path,
            "CSV Files (*.csv);;All Files (*)",
        )
        if not file_path:
            return
        if not file_path.lower().endswith(".csv"):
            file_path += ".csv"

        try:
            fieldnames = [
                "file_name",
                "file_path",
                "status",
                "suggestions",
                "suggestion_ids",
                "error",
                "font_summary",
                "font_details",
                "annotation_summary",
                "annotation_details",
                "broken_reference_summary",
                "broken_reference_details",
            ]
            with open(file_path, "w", encoding="utf-8-sig", newline="") as f:
                writer = csv.DictWriter(f, fieldnames=fieldnames)
                writer.writeheader()
                rows = []
                for row in self.last_precheck_results:
                    export_row = {key: row.get(key, "") for key in fieldnames}
                    rows.append(export_row)
                writer.writerows(rows)
            self.view.show_success_message("✅ 导出成功", "预检结果已成功保存！")
        except Exception as e:
            self.view.show_error_message("❌ 导出失败", f"文件保存失败：\n{str(e)}")

    def _open_directory(self, dir_path):
        try:
            system_shell.open_directory(dir_path)
        except Exception as e:
            self.process_logs += f"\n[警告] 自动打开文件夹失败：{str(e)}\n"
