import os
from copy import deepcopy
from pathlib import Path

from PySide6.QtCore import QSettings, QSize, Qt, QTimer
from PySide6.QtGui import QAction, QIcon
from PySide6.QtWidgets import (
    QApplication,
    QButtonGroup,
    QCheckBox,
    QComboBox,
    QDialog,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QMainWindow,
    QMenu,
    QProgressBar,
    QPushButton,
    QRadioButton,
    QScrollArea,
    QSizePolicy,
    QSpinBox,
    QStackedWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

from ratools_pdf.config import rules_catalog
from ratools_pdf.config.compression import (
    DEFAULT_IMAGE_DPI,
    DEFAULT_JPEG_QUALITY,
    MAX_IMAGE_DPI,
    MAX_JPEG_QUALITY,
    MIN_IMAGE_DPI,
    MIN_JPEG_QUALITY,
)
from ratools_pdf.config.paths import get_app_dir, get_resource_path
from ratools_pdf.services import profiles
from ratools_pdf.services.process_control import normalize_file_timeout
from ratools_pdf.ui import icons
from ratools_pdf.ui.dialogs import (
    AboutDialog,
    CustomMessageBox,
    ManualFontEmbeddingDialog,
    OverlayUnderlayDialog,
    PageAssemblyDialog,
    PdfInspectorDialog,
    SecurityCenterDialog,
    SettingsDialog,
)
from ratools_pdf.ui.dialogs.prompts import MajorUpdatePromptDialog, SignedFilesPromptDialog
from ratools_pdf.ui.theme import (
    ThemeManager,
    active_palette,
    apply_windows_title_bar_theme,
    build_app_qss,
)
from ratools_pdf.ui.widgets import (
    DropZoneFrame,
    ElidedLabel,
    NavModuleButton,
    QueueTreeWidget,
    RuleOptionCard,
    StatusBadgeDelegate,
)


class MainWindow(QMainWindow):
    # 规则目录与预设的唯一来源是 config.rules_catalog；
    # 类属性保留原名，兼容既有引用（controller/tests）。
    PRESET_OPTIONS = rules_catalog.PRESETS

    # 通用设置类勾选框：不参与预设/清空/规则统计（key 同时是显示文案）
    GENERAL_SETTING_OPTIONS = ("处理完成后自动打开输出文件夹", "覆盖原始文件 (不推荐)")

    # 右侧规则卡片语义徽章
    RULE_TAGS = {
        "page_size_a4": ("互斥", "info"),
        "page_size_letter": ("互斥", "info"),
        "flatten_annotations": ("不可逆", "warning"),
        "cleanup_remove_all_links_bookmarks": ("高风险", "danger"),
        "compress_images": ("有损", "warning"),
    }

    _LOSSLESS_IDS = frozenset({
        "compress_standard",
        "compress_aggressive",
        "generate_object_streams",
        "realloc_flate",
    })

    def __init__(self, settings_path=None):
        super().__init__()
        self.custom_selection_before_preset = set()
        self.active_preset_key = None
        self.is_applying_preset = False
        self.active_processing_profile = None
        self.current_module_index = 0
        self.setWindowTitle("RATools for PDF")

        # === 添加原生窗口图标 ===
        self.setWindowIcon(QIcon(get_resource_path("icon.ico")))

        self.resize(1120, 760)
        self.setMinimumSize(970, 620)

        self.all_checkboxes = {}
        self.current_file_count = 0

        # ================= 初始化 QSettings 与主题管理 =================
        ini_path = settings_path if settings_path is not None else os.path.join(get_app_dir(), "settings.ini")
        self.app_settings = QSettings(ini_path, QSettings.IniFormat)

        # 主题必须先应用到 QApplication，再让窗口级 apply_stylesheet() 做兼容回退；
        # 否则启动时会给 MainWindow 残留一份本地亮色 QSS，压过应用级暗色主题。
        app = QApplication.instance()
        theme_mode = str(self.app_settings.value("Settings/ThemeMode", "system") or "system").lower()
        if app is not None:
            self.theme_manager = ThemeManager(app, theme_mode)
            self.theme_manager.changed.connect(self.apply_native_title_bar_theme)
            self.theme_manager.changed.connect(self.refresh_icons)
            self.theme_manager.apply()
        else:
            self.theme_manager = None

        self.settings_dialog = SettingsDialog(self)
        self.all_checkboxes["处理完成后自动打开输出文件夹"] = self.settings_dialog.cb_auto_open
        self.all_checkboxes["覆盖原始文件 (不推荐)"] = self.settings_dialog.cb_overwrite
        self.settings_dialog.cb_auto_open.toggled.connect(self.on_checkbox_toggled)
        self.settings_dialog.cb_overwrite.toggled.connect(self.on_checkbox_toggled)
        self.settings_dialog.cb_overwrite.toggled.connect(lambda _checked: self.refresh_selection_summary())
        self.settings_dialog.default_output_edit.textChanged.connect(lambda _text: self.refresh_selection_summary())
        self.settings_dialog.default_output_edit.textChanged.connect(lambda _text: self.persist_default_output_dir())

        self.MODULES_DATA = rules_catalog.MODULES

        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        self._build_middle(main_layout)
        self._build_preset_bar(main_layout)
        self._build_progress_strip(main_layout)
        self._build_footer(main_layout)

        self.nav_btn_group.idClicked.connect(self.switch_settings_page)

        # 立刻触发一次以展示第一页内容
        self.switch_settings_page(0)
        self.apply_stylesheet()

        self.settings_key_map = {
            "处理完成后自动打开输出文件夹": "Settings/AutoOpenOutput",
            "覆盖原始文件 (不推荐)": "Settings/OverwriteOriginal"
        }
        for i, mod in enumerate(self.MODULES_DATA):
            for j, opt in enumerate(mod["options"]):
                self.settings_key_map[opt["id"]] = f"Modules/Mod_{i}_Opt_{j}"

        self.settings_dialog.cb_parallel_processing.toggled.connect(lambda _checked: self.persist_parallel_settings())
        self.settings_dialog.spin_parallel_workers.valueChanged.connect(lambda _value: self.persist_parallel_settings())
        self.settings_dialog.spin_file_timeout.valueChanged.connect(lambda _value: self.persist_parallel_settings())
        self.radio_smart_processing.toggled.connect(lambda _checked: self.persist_processing_mode())
        self.radio_force_processing.toggled.connect(lambda _checked: self.persist_processing_mode())

        self.settings_dialog.set_theme_mode(theme_mode)
        for mode, btn in self.settings_dialog.theme_buttons.items():
            btn.clicked.connect(lambda _checked=False, m=mode: self.on_theme_mode_changed(m))

        self.load_all_settings()
        self.refresh_selection_summary()
        self.refresh_icons(active_palette())

    def show_info_message(self, title, message):
        CustomMessageBox(title, message, msg_type="info", parent=self).exec()

    def show_success_message(self, title, message):
        CustomMessageBox(title, message, msg_type="success", parent=self).exec()

    def show_warning_message(self, title, message):
        CustomMessageBox(title, message, msg_type="warning", parent=self).exec()

    def show_error_message(self, title, message):
        CustomMessageBox(title, message, msg_type="error", parent=self).exec()

    def show_manual_font_embedding_dialog(self, pdf_paths, open_callback, *, capability_text="", recheck_callback=None):
        dlg = ManualFontEmbeddingDialog(
            pdf_paths, open_callback, parent=self,
            capability_text=capability_text, recheck_callback=recheck_callback,
        )
        dlg.exec()

    def show_confirm_message(self, title, message):
        dlg = CustomMessageBox(title, message, msg_type="question", show_cancel=True, parent=self)
        return dlg.exec() == QDialog.Accepted

    def show_major_update_prompt(self, current_version, release):
        dlg = MajorUpdatePromptDialog(current_version, release, parent=self)
        dlg.exec()
        return dlg.chosen_action

    def show_signed_files_prompt(self, signed_files):
        """检测到已签名文件时的三选一提示：跳过 / 全部处理 / 取消。

        返回 "skip"、"process_all" 或 "cancel"。
        """
        dlg = SignedFilesPromptDialog(signed_files, parent=self)
        dlg.exec()
        return dlg.chosen_action

    def show_page_assembly_dialog(self, initial_files=None):
        """弹出页面装配与拆分工作台对话框。"""
        if initial_files is None:
            initial_files = []
            if hasattr(self, "tree") and self.tree:
                root = self.tree.invisibleRootItem()
                for i in range(root.childCount()):
                    item = root.child(i)
                    path = item.text(1)
                    if path.lower().endswith(".pdf") and os.path.exists(path):
                        initial_files.append(path)
        dlg = PageAssemblyDialog(initial_files=initial_files, parent=self)
        dlg.exec()

    def show_security_center_dialog(self, initial_file=None):
        """弹出 PDF 安全与加密中心对话框。"""
        if initial_file is None:
            if hasattr(self, "tree") and self.tree:
                selected = self.tree.selectedItems()
                if selected and selected[0].text(1).lower().endswith(".pdf"):
                    initial_file = selected[0].text(1)
        dlg = SecurityCenterDialog(initial_file=initial_file, parent=self)
        dlg.exec()

    def show_overlay_dialog(self, initial_file=None):
        """弹出图层套印与印章管理器对话框。"""
        if initial_file is None:
            if hasattr(self, "tree") and self.tree:
                selected = self.tree.selectedItems()
                if selected and selected[0].text(1).lower().endswith(".pdf"):
                    initial_file = selected[0].text(1)
        dlg = OverlayUnderlayDialog(initial_file=initial_file, parent=self)
        dlg.exec()

    def show_pdf_inspector_dialog(self, initial_file=None):
        """弹出 PDF 深度诊断与受损修复面板。"""
        if initial_file is None:
            if hasattr(self, "tree") and self.tree:
                selected = self.tree.selectedItems()
                if selected and selected[0].text(1).lower().endswith(".pdf"):
                    initial_file = selected[0].text(1)
                else:
                    root = self.tree.invisibleRootItem()
                    for i in range(root.childCount()):
                        p = root.child(i).text(1)
                        if p.lower().endswith(".pdf") and os.path.exists(p):
                            initial_file = p
                            break
        dlg = PdfInspectorDialog(initial_file=initial_file, parent=self)
        dlg.exec()

    def _build_middle(self, main_layout):
        """中段：左侧导航 + 中间单层队列工作区（动态空态 / 文件树）+ 右侧规则设置区。"""
        middle_container = QFrame()
        middle_container.setObjectName("middleContainer")
        middle_layout = QHBoxLayout(middle_container)
        middle_layout.setContentsMargins(0, 0, 0, 0)
        middle_layout.setSpacing(0)

        # ================= 左侧导航栏 =================
        left_sidebar = QFrame()
        left_sidebar.setObjectName("leftSidebar")
        left_sidebar.setFixedWidth(216)
        left_layout = QVBoxLayout(left_sidebar)
        left_layout.setContentsMargins(12, 14, 12, 14)
        left_layout.setSpacing(6)

        nav_title = QLabel("功能模块")
        nav_title.setObjectName("navTitle")
        left_layout.addWidget(nav_title)

        self.nav_buttons = []
        self.nav_btn_group = QButtonGroup(self)
        self.nav_btn_group.setExclusive(True)

        for idx, mod in enumerate(self.MODULES_DATA):
            icon_name = icons.MODULE_ICON_MAP.get(mod["title"], "file_text")
            btn = NavModuleButton(mod["title"], icon_name=icon_name)
            self.nav_buttons.append(btn)
            self.nav_btn_group.addButton(btn, idx)
            left_layout.addWidget(btn)

        self.nav_buttons[0].setChecked(True)
        left_layout.addStretch()

        sidebar_sep = QFrame()
        sidebar_sep.setObjectName("sidebarSep")
        sidebar_sep.setFixedHeight(1)
        left_layout.addWidget(sidebar_sep)

        self.btn_top_security = QPushButton("安全中心")
        self.btn_top_security.setObjectName("sidebarUtilBtn")
        self.btn_top_security.setCursor(Qt.PointingHandCursor)
        self.btn_top_security.setIconSize(QSize(15, 15))
        self.btn_top_security.clicked.connect(lambda: self.show_security_center_dialog())
        left_layout.addWidget(self.btn_top_security)

        self.btn_top_settings = QPushButton("全局设置")
        self.btn_top_settings.setObjectName("sidebarUtilBtn")
        self.btn_top_settings.setCursor(Qt.PointingHandCursor)
        self.btn_top_settings.setIconSize(QSize(15, 15))
        self.btn_top_settings.clicked.connect(self.settings_dialog.show)
        left_layout.addWidget(self.btn_top_settings)

        # 暂时隐藏的入口按钮（代码保留供后续按需开放）
        self.btn_top_assembly = QPushButton("页面工作台", self)
        self.btn_top_assembly.setObjectName("sidebarUtilBtn")
        self.btn_top_assembly.clicked.connect(lambda: self.show_page_assembly_dialog())
        self.btn_top_assembly.setVisible(False)

        self.btn_top_overlay = QPushButton("图层套印", self)
        self.btn_top_overlay.setObjectName("sidebarUtilBtn")
        self.btn_top_overlay.clicked.connect(lambda: self.show_overlay_dialog())
        self.btn_top_overlay.setVisible(False)

        self.btn_top_inspector = QPushButton("深度诊断", self)
        self.btn_top_inspector.setObjectName("sidebarUtilBtn")
        self.btn_top_inspector.clicked.connect(lambda: self.show_pdf_inspector_dialog())
        self.btn_top_inspector.setVisible(False)

        sidebar_footer_row = QHBoxLayout()
        sidebar_footer_row.setContentsMargins(0, 2, 0, 0)
        sidebar_footer_row.setSpacing(6)

        self.btn_top_about = QPushButton("关于")
        self.btn_top_about.setObjectName("sidebarFooterBtn")
        self.btn_top_about.setCursor(Qt.PointingHandCursor)
        self.btn_top_about.setIconSize(QSize(14, 14))
        self.btn_top_about.clicked.connect(self.show_about_dialog)

        self.btn_theme_toggle = QPushButton("深色")
        self.btn_theme_toggle.setObjectName("themeQuickBtn")
        self.btn_theme_toggle.setCursor(Qt.PointingHandCursor)
        self.btn_theme_toggle.setIconSize(QSize(14, 14))
        self.btn_theme_toggle.setToolTip("快速切换浅色 / 深色外观主题")
        self.btn_theme_toggle.clicked.connect(self._toggle_quick_theme)

        sidebar_footer_row.addWidget(self.btn_top_about, 1)
        sidebar_footer_row.addWidget(self.btn_theme_toggle, 1)
        left_layout.addLayout(sidebar_footer_row)

        middle_layout.addWidget(left_sidebar)

        # ================= 中间队列工作区（单层悬浮卡片） =================
        main_view = QFrame()
        main_view.setObjectName("mainView")
        main_view_layout = QVBoxLayout(main_view)
        main_view_layout.setContentsMargins(14, 14, 14, 14)
        main_view_layout.setSpacing(0)

        queue_card = QFrame()
        queue_card.setObjectName("queueCard")
        queue_layout = QVBoxLayout(queue_card)
        queue_layout.setContentsMargins(1, 1, 1, 1)
        queue_layout.setSpacing(0)

        list_header = QFrame()
        list_header.setObjectName("listHeader")
        list_header_layout = QHBoxLayout(list_header)
        list_header_layout.setContentsMargins(14, 9, 12, 9)
        list_header_layout.setSpacing(6)

        self.list_title = QLabel("待处理队列 (0)")
        self.list_title.setObjectName("listTitle")
        self.list_title.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        self.queue_meta_label = ElidedLabel("当前队列为空")
        self.queue_meta_label.setObjectName("mutedLabel")

        self.btn_add_files_bar = QPushButton("文件")
        self.btn_add_files_bar.setObjectName("compactSecondaryBtn")
        self.btn_add_files_bar.setToolTip("添加 PDF 文件到待处理队列")
        self.btn_add_files_bar.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        self.btn_add_files_bar.setCursor(Qt.PointingHandCursor)
        self.btn_add_files_bar.setIconSize(QSize(14, 14))
        self.btn_add_files_bar.hide()

        self.btn_add_folder_bar = QPushButton("文件夹")
        self.btn_add_folder_bar.setObjectName("compactSecondaryBtn")
        self.btn_add_folder_bar.setToolTip("递归添加文件夹中的所有 PDF")
        self.btn_add_folder_bar.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        self.btn_add_folder_bar.setCursor(Qt.PointingHandCursor)
        self.btn_add_folder_bar.setIconSize(QSize(14, 14))
        self.btn_add_folder_bar.hide()

        self.btn_clear = QPushButton("清空")
        self.btn_clear.setObjectName("ghostBtn")
        self.btn_clear.setToolTip("清空当前待处理队列")
        self.btn_clear.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        self.btn_clear.setCursor(Qt.PointingHandCursor)
        self.btn_clear.setIconSize(QSize(14, 14))

        list_header_layout.addWidget(self.list_title)
        list_header_layout.addWidget(self.queue_meta_label, 1)
        list_header_layout.addWidget(self.btn_add_files_bar)
        list_header_layout.addWidget(self.btn_add_folder_bar)
        list_header_layout.addWidget(self.btn_clear)
        queue_layout.addWidget(list_header)

        # 空态引导 vs 文件树堆栈切换
        self.queue_stack = QStackedWidget()

        empty_page = QWidget()
        empty_page.setObjectName("emptyDropContainer")
        empty_page_layout = QVBoxLayout(empty_page)
        empty_page_layout.setContentsMargins(20, 20, 20, 20)

        self.drop_zone = DropZoneFrame()
        self.drop_zone.setCursor(Qt.PointingHandCursor)
        drop_layout = QVBoxLayout(self.drop_zone)
        drop_layout.setContentsMargins(36, 28, 36, 28)
        drop_layout.setSpacing(10)
        drop_layout.addStretch()

        self.empty_drop_icon = QLabel()
        self.empty_drop_icon.setObjectName("emptyDropIconBadge")
        self.empty_drop_icon.setFixedSize(48, 48)
        self.empty_drop_icon.setAlignment(Qt.AlignCenter)
        drop_layout.addWidget(self.empty_drop_icon, 0, Qt.AlignCenter)
        drop_layout.addSpacing(4)

        empty_title = QLabel("拖拽 PDF 文件或整个文件夹到这里")
        empty_title.setObjectName("emptyDropTitle")
        empty_title.setAlignment(Qt.AlignCenter)
        empty_title.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Minimum)
        drop_layout.addWidget(empty_title)

        empty_desc = QLabel("自动递归收集子目录中的 PDF 并保留相对目录层级，处理时不会改动原文件结构。")
        empty_desc.setObjectName("emptyDropDesc")
        empty_desc.setWordWrap(True)
        empty_desc.setAlignment(Qt.AlignCenter)
        empty_desc.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Minimum)
        drop_layout.addWidget(empty_desc)
        drop_layout.addSpacing(8)

        quick_actions = QHBoxLayout()
        quick_actions.setSpacing(10)
        quick_actions.addStretch()
        self.btn_add_files = QPushButton("选择 PDF 文件")
        self.btn_add_files.setObjectName("primaryBtn")
        self.btn_add_files.setCursor(Qt.PointingHandCursor)
        self.btn_add_files.setIconSize(QSize(14, 14))
        self.add_folder_btn = QPushButton("选择文件夹")
        self.add_folder_btn.setObjectName("secondaryBtn")
        self.add_folder_btn.setCursor(Qt.PointingHandCursor)
        self.add_folder_btn.setIconSize(QSize(14, 14))
        quick_actions.addWidget(self.btn_add_files)
        quick_actions.addWidget(self.add_folder_btn)
        quick_actions.addStretch()
        drop_layout.addLayout(quick_actions)
        drop_layout.addStretch()

        self.btn_add_files_bar.clicked.connect(self.btn_add_files.click)
        self.btn_add_folder_bar.clicked.connect(self.add_folder_btn.click)

        empty_page_layout.addWidget(self.drop_zone)
        self.queue_stack.addWidget(empty_page)

        # ================= 文件树视图 =================
        self.tree = QueueTreeWidget()
        self.tree.setHeaderLabels(["文件 / 文件夹", "绝对路径", "当前状态"])
        self.tree.header().setSectionResizeMode(0, QHeaderView.Interactive)
        self.tree.header().setSectionResizeMode(1, QHeaderView.Stretch)
        self.tree.header().setSectionResizeMode(2, QHeaderView.Fixed)
        self.tree.header().setStretchLastSection(False)
        self.tree.setColumnWidth(0, 270)
        self.tree.setColumnWidth(2, 120)

        self.tree.setSelectionBehavior(QueueTreeWidget.SelectRows)
        self.tree.setSelectionMode(QueueTreeWidget.ExtendedSelection)
        self.tree.setContextMenuPolicy(Qt.CustomContextMenu)
        self.tree.setIndentation(14)
        self.tree.setUniformRowHeights(True)
        self.tree.setAllColumnsShowFocus(True)
        self.tree.header().setMinimumSectionSize(90)
        self.tree.header().resizeSection(1, 360)
        self.tree.setAlternatingRowColors(True)
        self.tree.setRootIsDecorated(True)
        self.status_badge_delegate = StatusBadgeDelegate(self.tree)
        self.tree.setItemDelegateForColumn(2, self.status_badge_delegate)

        self.queue_stack.addWidget(self.tree)
        self.queue_stack.setCurrentIndex(0)
        queue_layout.addWidget(self.queue_stack, 1)

        main_view_layout.addWidget(queue_card, 1)
        middle_layout.addWidget(main_view, 1)

        # ================= 右侧设置区 =================
        right_sidebar = QFrame()
        right_sidebar.setObjectName("rightSidebar")
        right_sidebar.setFixedWidth(304)
        right_layout = QVBoxLayout(right_sidebar)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(0)

        right_header = QFrame()
        right_header.setObjectName("rightHeader")
        rh_layout = QVBoxLayout(right_header)
        rh_layout.setContentsMargins(16, 14, 16, 12)
        rh_layout.setSpacing(8)

        rh_top_row = QHBoxLayout()
        rh_top_row.setContentsMargins(0, 0, 0, 0)
        rh_top_row.setSpacing(8)
        self.rh_title = ElidedLabel("处理规则选项")
        self.rh_title.setObjectName("rightPanelTitle")
        self.selection_summary_label = QLabel("尚未选择任何处理规则")
        self.selection_summary_label.setObjectName("selectionSummary")
        self.selection_summary_label.setSizePolicy(QSizePolicy.Maximum, QSizePolicy.Fixed)
        rh_top_row.addWidget(self.rh_title, 1)
        rh_top_row.addWidget(self.selection_summary_label, 0, Qt.AlignVCenter)
        rh_layout.addLayout(rh_top_row)

        rh_sub_row = QHBoxLayout()
        rh_sub_row.setContentsMargins(0, 0, 0, 0)
        rh_sub_row.setSpacing(6)
        self.module_selection_label = QLabel("本模块已选 0 / 0")
        self.module_selection_label.setObjectName("moduleStatsLabel")
        self.btn_select_module = QPushButton("全选本页")
        self.btn_select_module.setObjectName("moduleActionLink")
        self.btn_select_module.setCursor(Qt.PointingHandCursor)
        self.btn_select_module.clicked.connect(self._select_current_module_options)
        self.btn_clear_module = QPushButton("清空")
        self.btn_clear_module.setObjectName("moduleActionLink")
        self.btn_clear_module.setCursor(Qt.PointingHandCursor)
        self.btn_clear_module.clicked.connect(self._clear_current_module_options)
        rh_sub_row.addWidget(self.module_selection_label)
        rh_sub_row.addStretch()
        rh_sub_row.addWidget(self.btn_select_module)
        rh_sub_row.addWidget(self.btn_clear_module)
        rh_layout.addLayout(rh_sub_row)

        right_layout.addWidget(right_header)

        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setFrameShape(QFrame.NoFrame)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        scroll_area.setObjectName("settingsScroll")

        self.settings_container = QWidget()
        self.settings_layout = QVBoxLayout(self.settings_container)
        self.settings_layout.setContentsMargins(0, 0, 0, 0)
        self.settings_layout.setSpacing(0)

        self.settings_pages = []

        self.btn_bookmark_io_wizard = QPushButton("书签数据导入/导出...")
        self.btn_link_io_wizard = QPushButton("链接数据导入/导出...")
        self.btn_embed_missing_fonts = QPushButton("嵌入缺失字体")
        self.btn_embed_missing_fonts.setObjectName("secondaryBtn")
        self.btn_embed_missing_fonts.setCursor(Qt.PointingHandCursor)
        self.btn_embed_missing_fonts.setIconSize(QSize(14, 14))
        self.btn_embed_missing_fonts.setToolTip("打开选中的 PDF，并在 Acrobat 中手动执行印前检查的“嵌入缺失的字体”。")

        self.btn_detect_annotations = QPushButton("检测批注 (便签/高亮等)")
        self.btn_detect_annotations.setObjectName("secondaryBtn")
        self.btn_detect_annotations.setCursor(Qt.PointingHandCursor)
        self.btn_detect_annotations.setIconSize(QSize(14, 14))
        self.btn_detect_annotations.setToolTip("扫描队列中所有 PDF，报告便签、高亮等批注情况，仅检测不修改文件。")

        self.btn_detect_broken_refs = QPushButton("检测失效引用/链接文本")
        self.btn_detect_broken_refs.setObjectName("secondaryBtn")
        self.btn_detect_broken_refs.setCursor(Qt.PointingHandCursor)
        self.btn_detect_broken_refs.setIconSize(QSize(14, 14))
        self.btn_detect_broken_refs.setToolTip("扫描队列中所有 PDF，检测 Word 转 PDF 后残留的“错误！未找到引用源”等失效占位文本，仅检测不修改文件。")

        self.btn_open_inspector_studio = QPushButton("打开 PDF 深度诊断与受损修复工作台...")
        self.btn_open_inspector_studio.setObjectName("secondaryBtn")
        self.btn_open_inspector_studio.setCursor(Qt.PointingHandCursor)
        self.btn_open_inspector_studio.setIconSize(QSize(14, 14))
        self.btn_open_inspector_studio.setToolTip("使用 qpdf 对 PDF 执行语法合规体检(--check)、查看加密透视、分析线性化报告及强制重建受损 XRef。")
        self.btn_open_inspector_studio.clicked.connect(lambda: self.show_pdf_inspector_dialog())

        for btn in [self.btn_bookmark_io_wizard, self.btn_link_io_wizard]:
            btn.setObjectName("secondaryBtn")
            btn.setCursor(Qt.PointingHandCursor)
            btn.setIconSize(QSize(14, 14))

        for mod in self.MODULES_DATA:
            page = QWidget()
            page_layout = QVBoxLayout(page)
            page_layout.setContentsMargins(14, 14, 14, 16)
            page_layout.setSpacing(8)

            if mod["title"] == "文件级优化与输出":
                for opt in mod["options"]:
                    if opt["id"] in ("convert_pdf_version", "fast_web_view"):
                        page_layout.addWidget(self._create_checkbox(opt["id"], opt["title"], opt["desc"], False))

                comp_card = self._build_compression_card()
                page_layout.addWidget(comp_card)

                for opt in mod["options"]:
                    if opt["id"] == "filename_ectd_format":
                        page_layout.addWidget(self._create_checkbox(opt["id"], opt["title"], opt["desc"], False))
            else:
                for opt in mod["options"]:
                    page_layout.addWidget(self._create_checkbox(opt["id"], opt["title"], opt["desc"], False))

            if mod["title"] == "页面与字体标准化":
                page_layout.addSpacing(8)
                page_layout.addWidget(self._create_section_label("字体手动处理"))
                font_action_hint = QLabel("如预检提示存在未嵌入字体，请先选中左侧 PDF，再点击下方按钮跳转 Acrobat 处理。")
                font_action_hint.setWordWrap(True)
                font_action_hint.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Minimum)
                font_action_hint.setObjectName("mutedSmall")
                page_layout.addWidget(font_action_hint)
                page_layout.addWidget(self.btn_embed_missing_fonts)

            elif mod["title"] == "书签管理":
                page_layout.addSpacing(8)
                page_layout.addWidget(self._create_section_label("导出 / 导入书签"))
                page_layout.addWidget(self.btn_bookmark_io_wizard)

            elif mod["title"] == "超链接处理":
                page_layout.addSpacing(8)
                page_layout.addWidget(self._create_section_label("导出 / 导入链接"))
                page_layout.addWidget(self.btn_link_io_wizard)

            elif mod["title"] == "文档检测":
                page_layout.addWidget(self._create_section_label("只读检测（不修改文件）"))
                detect_hint = QLabel("对整个待处理队列执行只读扫描，结果会弹窗汇总并写入日志，不会改动任何 PDF。")
                detect_hint.setWordWrap(True)
                detect_hint.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Minimum)
                detect_hint.setObjectName("mutedSmall")
                page_layout.addWidget(detect_hint)
                page_layout.addSpacing(4)
                page_layout.addWidget(self.btn_detect_annotations)
                page_layout.addWidget(self.btn_detect_broken_refs)

            page_layout.addStretch()

            self.settings_layout.addWidget(page)
            page.hide()
            self.settings_pages.append(page)

        scroll_area.setWidget(self.settings_container)
        right_layout.addWidget(scroll_area)

        middle_layout.addWidget(right_sidebar)
        main_layout.addWidget(middle_container, 1)

    def _build_preset_bar(self, main_layout):
        """快速预设栏：中国/美国 eCTD、我的常用与收敛的配置管理菜单。"""
        preset_bar = QFrame()
        preset_bar.setObjectName("presetBar")
        preset_layout = QHBoxLayout(preset_bar)
        preset_layout.setContentsMargins(20, 9, 20, 9)
        preset_layout.setSpacing(8)

        preset_label = QLabel("快速预设")
        preset_label.setObjectName("presetLabel")
        preset_label.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        preset_layout.addWidget(preset_label)

        self.btn_preset_china = QPushButton("中国eCTD")
        self.btn_preset_china.setObjectName("presetBtn")
        self.btn_preset_china.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        self.btn_preset_china.setCheckable(True)
        self.btn_preset_china.setCursor(Qt.PointingHandCursor)
        self.btn_preset_china.setFocusPolicy(Qt.NoFocus)
        preset_layout.addWidget(self.btn_preset_china)

        self.btn_preset_us = QPushButton("美国eCTD")
        self.btn_preset_us.setObjectName("presetBtn")
        self.btn_preset_us.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        self.btn_preset_us.setCheckable(True)
        self.btn_preset_us.setCursor(Qt.PointingHandCursor)
        self.btn_preset_us.setFocusPolicy(Qt.NoFocus)
        preset_layout.addWidget(self.btn_preset_us)

        self.btn_preset_favorite = QPushButton("我的常用")
        self.btn_preset_favorite.setObjectName("presetBtn")
        self.btn_preset_favorite.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        self.btn_preset_favorite.setCheckable(True)
        self.btn_preset_favorite.setCursor(Qt.PointingHandCursor)
        self.btn_preset_favorite.setFocusPolicy(Qt.NoFocus)
        preset_layout.addWidget(self.btn_preset_favorite)

        self.preset_btn_group = QButtonGroup(self)
        self.preset_btn_group.setExclusive(True)
        self.preset_btn_group.addButton(self.btn_preset_china)
        self.preset_btn_group.addButton(self.btn_preset_us)
        self.preset_btn_group.addButton(self.btn_preset_favorite)

        self.preset_summary_label = ElidedLabel("默认载入中国 eCTD 预设，可按需微调。")
        self.preset_summary_label.setObjectName("presetSummary")
        preset_layout.addSpacing(4)
        preset_layout.addWidget(self.preset_summary_label, 1)

        self.btn_clear_selected_options = QPushButton("全部取消")
        self.btn_clear_selected_options.setObjectName("ghostBtn")
        self.btn_clear_selected_options.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        self.btn_clear_selected_options.setCursor(Qt.PointingHandCursor)
        preset_layout.addWidget(self.btn_clear_selected_options)

        # 保留真实按钮实例与对外属性，同时用收敛的“配置管理”菜单提供干净的视觉呈现
        self.btn_save_favorite_preset = QPushButton("保存为常用")
        self.btn_import_profile = QPushButton("导入配置")
        self.btn_export_profile = QPushButton("导出配置")
        for button in (self.btn_save_favorite_preset, self.btn_import_profile, self.btn_export_profile):
            button.setObjectName("secondaryBtn")
            button.hide()
            preset_layout.addWidget(button)
        self.btn_import_profile.clicked.connect(self.import_processing_profile)
        self.btn_export_profile.clicked.connect(self.export_processing_profile)

        self.btn_config_menu = QPushButton("配置管理")
        self.btn_config_menu.setObjectName("configMenuBtn")
        self.btn_config_menu.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        self.btn_config_menu.setCursor(Qt.PointingHandCursor)
        self.btn_config_menu.setIconSize(QSize(14, 14))
        config_menu = QMenu(self.btn_config_menu)
        act_save_fav = QAction("保存为我的常用", self)
        act_import = QAction("导入配置 (JSON)...", self)
        act_export = QAction("导出配置 (JSON)...", self)
        act_save_fav.triggered.connect(self.btn_save_favorite_preset.click)
        act_import.triggered.connect(self.btn_import_profile.click)
        act_export.triggered.connect(self.btn_export_profile.click)
        config_menu.addAction(act_save_fav)
        config_menu.addSeparator()
        config_menu.addAction(act_import)
        config_menu.addAction(act_export)
        self.btn_config_menu.setMenu(config_menu)
        preset_layout.addWidget(self.btn_config_menu)

        main_layout.addWidget(preset_bar)

    def _build_progress_strip(self, main_layout):
        """底部操作栏顶部的 3px 细进度条（空闲时隐藏，处理/预检时显现）。"""
        self.progress_bar = QProgressBar()
        self.progress_bar.setObjectName("thinProgressBar")
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.hide()
        main_layout.addWidget(self.progress_bar)

    def _build_footer(self, main_layout):
        """底部操作栏：日志、状态摘要、模式分段切换、动作按钮组、开始按钮。"""
        footer = QFrame()
        footer.setObjectName("footer")
        footer.setFixedHeight(60)
        footer_layout = QHBoxLayout(footer)
        footer_layout.setContentsMargins(20, 10, 20, 10)
        footer_layout.setSpacing(10)

        self.btn_log = QPushButton("日志")
        self.btn_log.setObjectName("actionBtn")
        self.btn_log.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        self.btn_log.setCursor(Qt.PointingHandCursor)
        self.btn_log.setIconSize(QSize(14, 14))
        self.btn_log.setToolTip("查看或导出处理日志")
        footer_layout.addWidget(self.btn_log)

        self.info_label = ElidedLabel("0 个文件 · 0 条规则 · 自定义选择")
        self.info_label.setObjectName("footerSummary")
        self.processing_hint_label = ElidedLabel("")
        self.processing_hint_label.setObjectName("processingHint")
        self.processing_hint_label.hide()
        self.risk_hint_label = ElidedLabel("")
        self.risk_hint_label.setObjectName("footerHint")
        self.risk_hint_label.hide()
        footer_layout.addWidget(self.info_label)
        footer_layout.addWidget(self.processing_hint_label)
        footer_layout.addWidget(self.risk_hint_label)
        footer_layout.addStretch()

        mode_seg = QFrame()
        mode_seg.setObjectName("modeSegContainer")
        mode_seg.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        mode_seg_layout = QHBoxLayout(mode_seg)
        mode_seg_layout.setContentsMargins(3, 3, 3, 3)
        mode_seg_layout.setSpacing(2)

        self.processing_mode_group = QButtonGroup(self)
        self.radio_smart_processing = QRadioButton("智能处理")
        self.radio_smart_processing.setObjectName("modeRadio")
        self.radio_smart_processing.setCursor(Qt.PointingHandCursor)
        self.radio_smart_processing.setFocusPolicy(Qt.NoFocus)
        self.radio_smart_processing.setToolTip(
            "处理前先按已勾选规则进行预检，仅处理预检发现需要修改的规则，速度更快。无法可靠预检的勾选规则仍会执行。"
        )
        self.radio_force_processing = QRadioButton("全部处理")
        self.radio_force_processing.setObjectName("modeRadio")
        self.radio_force_processing.setCursor(Qt.PointingHandCursor)
        self.radio_force_processing.setFocusPolicy(Qt.NoFocus)
        self.radio_force_processing.setToolTip("强制执行全部已勾选规则，不做预检筛选，处理更彻底但耗时更长。")
        self.radio_smart_processing.setChecked(True)
        self.processing_mode_group.addButton(self.radio_smart_processing)
        self.processing_mode_group.addButton(self.radio_force_processing)
        mode_seg_layout.addWidget(self.radio_smart_processing)
        mode_seg_layout.addWidget(self.radio_force_processing)

        self.btn_skip_current = QPushButton("跳过当前文件")
        self.btn_skip_current.setObjectName("actionBtn")
        self.btn_skip_current.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        self.btn_skip_current.setCursor(Qt.PointingHandCursor)
        self.btn_skip_current.setIconSize(QSize(14, 14))
        self.btn_skip_current.setEnabled(False)
        self.btn_skip_current.hide()

        self.btn_retry_failed = QPushButton("重试失败项")
        self.btn_retry_failed.setObjectName("actionBtn")
        self.btn_retry_failed.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        self.btn_retry_failed.setCursor(Qt.PointingHandCursor)
        self.btn_retry_failed.setIconSize(QSize(14, 14))
        self.btn_retry_failed.setEnabled(False)
        self.btn_retry_failed.hide()

        self.btn_apply_precheck = QPushButton("应用建议")
        self.btn_apply_precheck.setObjectName("actionBtn")
        self.btn_apply_precheck.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        self.btn_apply_precheck.setCursor(Qt.PointingHandCursor)
        self.btn_apply_precheck.setIconSize(QSize(14, 14))
        self.btn_apply_precheck.setEnabled(False)
        self.btn_apply_precheck.hide()

        self.btn_process_precheck_suggested = QPushButton("处理建议文件")
        self.btn_process_precheck_suggested.setObjectName("actionBtn")
        self.btn_process_precheck_suggested.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        self.btn_process_precheck_suggested.setCursor(Qt.PointingHandCursor)
        self.btn_process_precheck_suggested.setIconSize(QSize(14, 14))
        self.btn_process_precheck_suggested.setEnabled(False)
        self.btn_process_precheck_suggested.hide()

        self.btn_precheck = QPushButton("预检")
        self.btn_precheck.setObjectName("actionBtn")
        self.btn_precheck.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        self.btn_precheck.setCursor(Qt.PointingHandCursor)
        self.btn_precheck.setIconSize(QSize(14, 14))

        self.btn_start = QPushButton("开始处理")
        self.btn_start.setObjectName("startBtn")
        self.btn_start.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        self.btn_start.setCursor(Qt.PointingHandCursor)
        self.btn_start.setIconSize(QSize(14, 14))

        footer_layout.addWidget(mode_seg, 0, Qt.AlignVCenter)
        footer_layout.addWidget(self.btn_skip_current)
        footer_layout.addWidget(self.btn_retry_failed)
        footer_layout.addWidget(self.btn_apply_precheck)
        footer_layout.addWidget(self.btn_process_precheck_suggested)
        footer_layout.addWidget(self.btn_precheck)
        footer_layout.addWidget(self.btn_start)
        main_layout.addWidget(footer)

    def set_progress_state(self, visible: bool, value: int = 0, maximum: int = 100):
        """控制底部细进度条可见性与进度值。"""
        bar = getattr(self, "progress_bar", None)
        if bar is None:
            return
        max_val = max(0, int(maximum))
        bar.setRange(0, max_val)
        if max_val > 0:
            bar.setValue(max(0, min(int(value), max_val)))
        bar.setVisible(bool(visible))

    def load_all_settings(self):
        self.is_applying_preset = True
        for opt_id, cb in self.all_checkboxes.items():
            key = self.settings_key_map.get(opt_id)
            if key:
                val = self.app_settings.value(key)
                if val is not None:
                    is_checked = str(val).lower() == 'true'
                    cb.setChecked(is_checked)

        default_output_dir = str(self.app_settings.value("Settings/DefaultOutputDir", "") or "")
        self.settings_dialog.default_output_edit.setText(default_output_dir)

        parallel_enabled = str(self.app_settings.value("Settings/ParallelProcessingEnabled", "false")).lower() == 'true'
        try:
            parallel_workers = int(self.app_settings.value("Settings/ParallelWorkerCount", 2))
        except Exception:
            parallel_workers = 2
        self.settings_dialog.spin_parallel_workers.setValue(parallel_workers)
        self.settings_dialog.cb_parallel_processing.setChecked(parallel_enabled)
        self.settings_dialog.set_parallel_worker_controls_enabled(parallel_enabled)
        try:
            file_timeout = normalize_file_timeout(self.app_settings.value("Settings/FileTimeoutSeconds", 0))
        except ValueError:
            file_timeout = 0
        self.settings_dialog.spin_file_timeout.setValue(int(file_timeout))

        processing_mode = str(self.app_settings.value("Settings/ProcessingMode", "smart") or "smart").lower()
        self.radio_force_processing.setChecked(processing_mode == "force")
        self.radio_smart_processing.setChecked(processing_mode != "force")

        self.is_applying_preset = False

        # 默认恢复上次会话的勾选状态（自定义），不自动套用预设
        self.custom_selection_before_preset = set(self.get_selected_options())
        self.active_preset_key = None
        self._set_preset_button_state(None)

        # 同步压缩卡片单选态与 DPI 设置
        if hasattr(self, "rb_compress_aggressive"):
            if self.cb_compress_aggressive.isChecked():
                self.rb_compress_aggressive.setChecked(True)
                self.lbl_compress_hint.setText("最大化清理对象与结构碎片 (garbage=4 + clean)，适用于超限文件。")
            elif self.cb_compress_standard.isChecked() or self.cb_generate_object_streams.isChecked():
                self.rb_compress_standard.setChecked(True)
                self.lbl_compress_hint.setText("融合对象流打包、Flate 最优重压缩与垃圾回收，100% 保真。")
            else:
                self.rb_compress_none.setChecked(True)
                self.lbl_compress_hint.setText("保持原始文档字节流与对象组织不变。")

        if hasattr(self, "combo_image_dpi"):
            try:
                saved_dpi = int(self.app_settings.value("Compression/ImageDPI", DEFAULT_IMAGE_DPI))
            except Exception:
                saved_dpi = DEFAULT_IMAGE_DPI
            if saved_dpi in (150, 200, 300):
                idx = self.combo_image_dpi.findData(saved_dpi)
                if idx >= 0:
                    self.combo_image_dpi.setCurrentIndex(idx)
                if hasattr(self, "row_custom_widget"):
                    self.row_custom_widget.setVisible(False)
            else:
                idx = self.combo_image_dpi.findData(-1)
                if idx >= 0:
                    self.combo_image_dpi.setCurrentIndex(idx)
                self.spin_image_dpi.setValue(saved_dpi)
                if hasattr(self, "row_custom_widget"):
                    self.row_custom_widget.setVisible(True)
        if hasattr(self, "spin_image_quality"):
            try:
                quality = int(self.app_settings.value("Compression/ImageQuality", DEFAULT_JPEG_QUALITY))
            except (TypeError, ValueError):
                quality = DEFAULT_JPEG_QUALITY
            self.spin_image_quality.setValue(quality)

    def closeEvent(self, event):
        close_guard = getattr(self, "close_guard", None)
        if close_guard is not None and not close_guard():
            self.processing_hint_label.setText("后台任务仍在运行，请等待完成后关闭。")
            if hasattr(self.processing_hint_label, "setVisible"):
                self.processing_hint_label.setVisible(True)
            event.ignore()
            return
        self.persist_all_settings()
        super().closeEvent(event)

    def persist_all_settings(self):
        if not hasattr(self, "app_settings"):
            return
        for opt_id, cb in self.all_checkboxes.items():
            key = self.settings_key_map.get(opt_id)
            if key:
                self.app_settings.setValue(key, cb.isChecked())
        self.app_settings.setValue("Settings/DefaultOutputDir", self.settings_dialog.default_output_edit.text().strip())
        self.app_settings.setValue("Settings/ParallelProcessingEnabled", self.settings_dialog.cb_parallel_processing.isChecked())
        self.app_settings.setValue("Settings/ParallelWorkerCount", self.settings_dialog.spin_parallel_workers.value())
        self.app_settings.setValue("Settings/ProcessingMode", self.get_processing_mode())
        self.app_settings.setValue("Settings/FileTimeoutSeconds", self.settings_dialog.spin_file_timeout.value())

    def persist_parallel_settings(self):
        if not hasattr(self, "app_settings") or self.is_applying_preset:
            return
        self.app_settings.setValue("Settings/ParallelProcessingEnabled", self.settings_dialog.cb_parallel_processing.isChecked())
        self.app_settings.setValue("Settings/ParallelWorkerCount", self.settings_dialog.spin_parallel_workers.value())
        self.app_settings.setValue("Settings/FileTimeoutSeconds", self.settings_dialog.spin_file_timeout.value())

    def persist_processing_mode(self):
        if not hasattr(self, "app_settings"):
            return
        self.app_settings.setValue("Settings/ProcessingMode", self.get_processing_mode())

    def _toggle_quick_theme(self):
        next_mode = "light" if active_palette().is_dark else "dark"
        self.on_theme_mode_changed(next_mode)

    def on_theme_mode_changed(self, mode):
        """切换外观主题：更新分段按钮选中态、应用配色并持久化。"""
        self.settings_dialog.set_theme_mode(mode)
        if self.theme_manager is not None:
            self.theme_manager.set_mode(mode)
        if hasattr(self, "app_settings"):
            self.app_settings.setValue("Settings/ThemeMode", mode)
        self.refresh_icons(active_palette())
        QTimer.singleShot(0, self.apply_native_title_bar_theme)

    def apply_native_title_bar_theme(self, palette=None):
        """同步 Windows 原生标题栏颜色；非 Windows 平台会静默跳过。"""
        apply_windows_title_bar_theme(self, palette or active_palette())

    def refresh_icons(self, palette=None):
        """依据当前 Palette 刷新全窗口线性 SVG 图标。"""
        pal = palette or active_palette()

        for attr, icon_name, icon_size in (
            ("btn_top_security", "lock", 15),
            ("btn_top_settings", "settings", 15),
            ("btn_top_about", "info", 14),
            ("btn_top_assembly", "file_text", 15),
            ("btn_top_overlay", "package", 15),
            ("btn_top_inspector", "search", 15),
        ):
            btn = getattr(self, attr, None)
            if btn is not None:
                btn.setIcon(
                    icons.svg_icon(
                        icon_name,
                        pal.text_body,
                        size=icon_size,
                        disabled_color=pal.text_faint,
                    )
                )

        theme_btn = getattr(self, "btn_theme_toggle", None)
        if theme_btn is not None:
            if pal.is_dark:
                theme_btn.setIcon(icons.svg_icon("sun", pal.text_body, size=14))
                theme_btn.setText("浅色")
            else:
                theme_btn.setIcon(icons.svg_icon("moon", pal.text_body, size=14))
                theme_btn.setText("深色")

        for nav_btn in getattr(self, "nav_buttons", []):
            if hasattr(nav_btn, "set_theme_icon"):
                nav_btn.set_theme_icon(pal)

        empty_icon = getattr(self, "empty_drop_icon", None)
        if empty_icon is not None:
            empty_icon.setPixmap(icons.render_svg_pixmap("upload", pal.primary_text, size=24, stroke_width=1.85))

        add_files_btn = getattr(self, "btn_add_files", None)
        if add_files_btn is not None:
            add_files_btn.setIcon(
                icons.svg_icon("plus", pal.text_on_primary, size=14, disabled_color=pal.surface)
            )

        for attr, icon_name in (
            ("add_folder_btn", "folder"),
            ("btn_add_files_bar", "plus"),
            ("btn_add_folder_bar", "folder"),
            ("btn_clear", "trash"),
            ("btn_embed_missing_fonts", "wrench"),
            ("btn_bookmark_io_wizard", "arrow_up_down"),
            ("btn_link_io_wizard", "arrow_up_down"),
            ("btn_detect_annotations", "scan"),
            ("btn_detect_broken_refs", "link"),
            ("btn_open_inspector_studio", "search"),
            ("btn_config_menu", "sliders"),
            ("btn_log", "list_log"),
            ("btn_precheck", "search"),
            ("btn_skip_current", "skip"),
            ("btn_retry_failed", "refresh"),
            ("btn_apply_precheck", "check"),
            ("btn_process_precheck_suggested", "play"),
        ):
            btn = getattr(self, attr, None)
            if btn is not None:
                btn.setIcon(icons.svg_icon(icon_name, pal.text_body, size=14, disabled_color=pal.text_faint))

        start_btn = getattr(self, "btn_start", None)
        if start_btn is not None:
            if start_btn.property("stopMode") is True:
                start_btn.setIcon(
                    icons.svg_icon(
                        "square", pal.danger_text, size=13,
                        active_color=pal.text_on_primary, disabled_color=pal.surface,
                    )
                )
            else:
                start_btn.setIcon(
                    icons.svg_icon("play", pal.text_on_primary, size=13, disabled_color=pal.surface)
                )

    def persist_current_checkbox(self, checkbox):
        if not hasattr(self, "app_settings"):
            return
        for opt_id, cb in self.all_checkboxes.items():
            if cb is checkbox:
                key = self.settings_key_map.get(opt_id)
                if key:
                    self.app_settings.setValue(key, checkbox.isChecked())
                    return

    def persist_default_output_dir(self):
        if not hasattr(self, "app_settings"):
            return
        self.app_settings.setValue("Settings/DefaultOutputDir", self.settings_dialog.default_output_edit.text().strip())

    def get_selected_preset(self):
        return self.active_preset_key

    def get_processing_mode(self):
        return "force" if self.radio_force_processing.isChecked() else "smart"

    def get_processing_mode_label(self):
        if self.get_processing_mode() == "force":
            return "全部处理"
        return "智能处理"

    def _set_preset_button_state(self, preset_key):
        self.preset_btn_group.setExclusive(False)
        self.btn_preset_china.setChecked(preset_key == "china")
        self.btn_preset_us.setChecked(preset_key == "us")
        self.btn_preset_favorite.setChecked(preset_key == "favorite")
        self.preset_btn_group.setExclusive(True)

    def get_processing_options(self):
        return [
            opt_id
            for mod in self.MODULES_DATA
            for opt in mod["options"]
            for opt_id in [opt["id"]]
            if self.all_checkboxes.get(opt_id) and self.all_checkboxes[opt_id].isChecked()
        ]

    def get_favorite_preset_options(self):
        saved = self.app_settings.value("Presets/FavoriteProfile", "")
        if saved:
            return profiles.loads_profile(saved)["enabled_rules"]
        value = self.app_settings.value("Presets/FavoriteOptions", [])
        if isinstance(value, str):
            return [item for item in value.split(",") if item]
        return list(value or [])

    def save_favorite_preset(self):
        try:
            profile = self.get_processing_profile()
        except ValueError as exc:
            self.show_warning_message("配置未保存", str(exc))
            return
        favorite_options = self.get_processing_options()
        self.app_settings.setValue("Presets/FavoriteProfile", profiles.dumps_profile(profile))
        self.app_settings.setValue("Presets/FavoriteOptions", favorite_options)
        self.active_preset_key = "favorite"
        self._set_preset_button_state("favorite")
        self.custom_selection_before_preset = set(favorite_options)
        self.refresh_selection_summary()
        self.show_success_message("已保存", "当前处理规则已保存为我的常用。")

    def _rule_checkbox_groups(self):
        """规则勾选框按控件聚合（同一控件可能挂多个 option id 别名）。"""
        groups = {}
        for opt_id, cb in self.all_checkboxes.items():
            if opt_id in self.GENERAL_SETTING_OPTIONS:
                continue
            groups.setdefault(id(cb), {"checkbox": cb, "keys": set()})["keys"].add(opt_id)
        return groups

    def _set_rule_selection(self, target_options):
        """按目标 id 集合刷勾选状态（预设应用/恢复自定义/清空共用）。"""
        self.is_applying_preset = True
        try:
            for group in self._rule_checkbox_groups().values():
                should_check = any(key in target_options for key in group["keys"])
                group["checkbox"].setChecked(should_check)

            # 同步更新压缩卡片分段按钮态
            if hasattr(self, "rb_compress_aggressive"):
                if "compress_aggressive" in target_options:
                    self.rb_compress_aggressive.setChecked(True)
                    self.lbl_compress_hint.setText("最大化清理对象与结构碎片 (garbage=4 + clean)，适用于超限文件。")
                elif "compress_standard" in target_options or "generate_object_streams" in target_options:
                    self.rb_compress_standard.setChecked(True)
                    self.lbl_compress_hint.setText("融合对象流打包、Flate 最优重压缩与垃圾回收，100% 保真。")
                else:
                    self.rb_compress_none.setChecked(True)
                    self.lbl_compress_hint.setText("保持原始文档字节流与对象组织不变。")

            if hasattr(self, "cb_compress_images"):
                self.cb_compress_images.setChecked("compress_images" in target_options)
        finally:
            self.is_applying_preset = False

    def restore_custom_selection(self):
        self._set_rule_selection(self.custom_selection_before_preset)
        self.active_processing_profile = None
        self.active_preset_key = None
        self._set_preset_button_state(None)
        self.persist_all_settings()
        self.refresh_selection_summary()

    def toggle_preset(self, preset_key):
        if preset_key == "favorite" and not self.get_favorite_preset_options():
            self.show_warning_message("尚未保存", "请先在配置管理中保存当前规则组合为我的常用。")
            self._set_preset_button_state(self.active_preset_key)
            return

        if self.active_preset_key == preset_key:
            self.restore_custom_selection()
            return

        if self.active_preset_key is None:
            self.custom_selection_before_preset = set(self.get_selected_options())

        self.apply_preset(preset_key)

    def clear_selected_options(self):
        self._set_rule_selection(set())
        self.active_processing_profile = None
        self.active_preset_key = None
        self._set_preset_button_state(None)
        self.custom_selection_before_preset = set()
        self.persist_all_settings()
        self.refresh_selection_summary()

    def _module_option_stats(self, mod):
        """返回模块在 UI 上呈现的可选项统计：(已选数量, 总数量)。"""
        if mod["title"] == "文件级优化与输出":
            visible_ids = ("convert_pdf_version", "fast_web_view", "compress_images", "filename_ectd_format")
            checked = sum(
                1 for opt_id in visible_ids
                if self.all_checkboxes.get(opt_id) and self.all_checkboxes[opt_id].isChecked()
            )
            has_lossless = any(
                self.all_checkboxes.get(opt_id) and self.all_checkboxes[opt_id].isChecked()
                for opt_id in self._LOSSLESS_IDS
            )
            return checked + (1 if has_lossless else 0), len(visible_ids) + 1

        options = mod.get("options", [])
        checked = sum(
            1 for opt in options
            if self.all_checkboxes.get(opt["id"]) and self.all_checkboxes[opt["id"]].isChecked()
        )
        return checked, len(options)

    def _refresh_module_indicators(self):
        nav_buttons = getattr(self, "nav_buttons", None)
        modules = getattr(self, "MODULES_DATA", None)
        if not nav_buttons or not modules:
            return

        for idx, mod in enumerate(modules):
            if idx < len(nav_buttons) and hasattr(nav_buttons[idx], "set_badge_count"):
                selected, _total = self._module_option_stats(mod)
                nav_buttons[idx].set_badge_count(selected)

        mod_idx = getattr(self, "current_module_index", 0)
        if 0 <= mod_idx < len(modules):
            current_mod = modules[mod_idx]
            selected, total = self._module_option_stats(current_mod)
            stats_lbl = getattr(self, "module_selection_label", None)
            btn_select = getattr(self, "btn_select_module", None)
            btn_clear = getattr(self, "btn_clear_module", None)
            if stats_lbl is not None:
                stats_lbl.setText(f"本模块已选 {selected} / {total}" if total > 0 else "只读检测工具")
            if btn_select is not None:
                btn_select.setVisible(total > 0)
                btn_select.setEnabled(total > 0 and selected < total)
            if btn_clear is not None:
                btn_clear.setVisible(total > 0)
                btn_clear.setEnabled(selected > 0)

    def _select_current_module_options(self):
        """一键勾选当前模块的所有非冲突、非高风险规则。"""
        mod_idx = getattr(self, "current_module_index", 0)
        if not (0 <= mod_idx < len(self.MODULES_DATA)):
            return
        mod = self.MODULES_DATA[mod_idx]
        target = set(self.get_processing_options())

        skip_ids = {"cleanup_remove_all_links_bookmarks"}
        conflict_pairs = (
            ("page_size_a4", "page_size_letter"),
            ("cleanup_remove_external_uri", "cleanup_remove_external_uri_and_text_black"),
            ("cleanup_remove_invalid_links", "cleanup_remove_invalid_links_and_text_black"),
            ("flatten_annotations", "cleanup_remove_annotations"),
            ("cleanup_remove_metadata", "title_from_filename"),
        )
        link_border_group = {
            "link_remove_border",
            "link_black_border",
            "link_bordered_to_blue_border",
            "link_unbordered_blue_to_blue_border",
        }

        for opt in mod.get("options", []):
            opt_id = opt["id"]
            if opt_id in skip_ids or opt_id in self._LOSSLESS_IDS:
                continue
            if opt_id in target:
                continue
            blocked = False
            for a, b in conflict_pairs:
                if opt_id == a and b in target:
                    blocked = True
                elif opt_id == b and a in target:
                    blocked = True
            if opt_id in link_border_group:
                active_lb = target & link_border_group
                if active_lb and not (
                    opt_id in {"link_bordered_to_blue_border", "link_unbordered_blue_to_blue_border"}
                    and active_lb <= {"link_bordered_to_blue_border", "link_unbordered_blue_to_blue_border"}
                ):
                    blocked = True
            if not blocked:
                target.add(opt_id)

        if mod["title"] == "文件级优化与输出" and not (target & self._LOSSLESS_IDS):
            target.update({"compress_standard", "generate_object_streams", "realloc_flate"})

        self._set_rule_selection(target)
        self.active_processing_profile = None
        self.active_preset_key = None
        self._set_preset_button_state(None)
        self.custom_selection_before_preset = set(self.get_selected_options())
        self.persist_all_settings()
        self.refresh_selection_summary()

    def _clear_current_module_options(self):
        """清空当前模块的已选规则。"""
        mod_idx = getattr(self, "current_module_index", 0)
        if not (0 <= mod_idx < len(self.MODULES_DATA)):
            return
        mod = self.MODULES_DATA[mod_idx]
        mod_ids = {opt["id"] for opt in mod.get("options", [])}
        target = set(self.get_processing_options()) - mod_ids
        self._set_rule_selection(target)
        self.active_processing_profile = None
        self.active_preset_key = None
        self._set_preset_button_state(None)
        self.custom_selection_before_preset = set(self.get_selected_options())
        self.persist_all_settings()
        self.refresh_selection_summary()

    def apply_preset(self, preset_key, persist=True):
        if preset_key == "favorite":
            try:
                saved = self.app_settings.value("Presets/FavoriteProfile", "")
                profile = profiles.loads_profile(saved) if saved else profiles.migrate_favorite(
                    self.app_settings.value("Presets/FavoriteOptions", []), self.get_compression_settings())
                self.apply_processing_profile(profile, persist=persist)
            except ValueError as exc:
                self.show_warning_message("常用配置无效", str(exc))
                return
            self.active_preset_key = preset_key
            self._set_preset_button_state(preset_key)
            self.refresh_selection_summary()
            return
        else:
            preset = self.PRESET_OPTIONS.get(preset_key)
            if not preset:
                return
            target_options = set(preset["options"])

        if not target_options:
            return

        self._set_rule_selection(target_options)
        self.active_processing_profile = profiles.builtin_profile(preset_key)
        self.active_preset_key = preset_key
        self._set_preset_button_state(preset_key)
        self.persist_all_settings()
        self.refresh_selection_summary()

    def get_processing_profile(self):
        rules = sorted(self.get_processing_options())
        current = self.active_processing_profile
        parameters = {"compress_images": self.get_compression_settings()} if "compress_images" in rules else {}
        if current and current["enabled_rules"] == rules and current["rule_parameters"] == parameters:
            return deepcopy(current)
        self.active_processing_profile = None
        return profiles.custom_profile(rules, self.get_compression_settings())

    def _validate_profile_selection(self, profile):
        """Reject aliases and radio combinations the desktop cannot represent exactly."""
        target = set(profile["enabled_rules"])
        expanded = set()
        for group in self._rule_checkbox_groups().values():
            if target & group["keys"]:
                expanded.update(group["keys"])
        if expanded != target:
            raise ValueError("此配置包含界面无法精确表示的规则或别名组合")
        lossless = target & {"compress_standard", "compress_aggressive", "generate_object_streams", "realloc_flate"}
        if lossless not in (set(), {"compress_standard", "generate_object_streams", "realloc_flate"},
                            {"compress_aggressive", "generate_object_streams", "realloc_flate"}):
            raise ValueError("压缩规则必须与界面的完整无损压缩级别一致，未自动增减规则")

    def apply_processing_profile(self, payload, persist=True):
        profile = profiles.validate_profile(payload)
        self._validate_profile_selection(profile)
        # All validation precedes widget changes; handlers suppress persistence during application.
        previous = self.get_processing_options()
        previous_parameters = self.get_compression_settings()
        previous_identity = self.active_processing_profile
        try:
            self._set_rule_selection(set(profile["enabled_rules"]))
            self.is_applying_preset = True
            params = profile["rule_parameters"].get("compress_images", previous_parameters)
            self._set_profile_image_parameters(params)
            actual = sorted(self.get_processing_options())
            actual_parameters = {"compress_images": self.get_compression_settings()} if "compress_images" in actual else {}
            if actual != profile["enabled_rules"] or actual_parameters != profile["rule_parameters"]:
                raise ValueError("界面实际配置与导入配置不一致，已取消应用")
        except Exception:
            self._set_rule_selection(set(previous))
            self.is_applying_preset = True
            self._set_profile_image_parameters(previous_parameters)
            self.active_processing_profile = previous_identity
            raise
        finally:
            self.is_applying_preset = False
        self.active_processing_profile = deepcopy(profile)
        self.active_preset_key = None
        self._set_preset_button_state(None)
        if persist:
            self.persist_all_settings()
            self.app_settings.setValue("Compression/ImageDPI", params["dpi"])
            self.app_settings.setValue("Compression/ImageQuality", params["quality"])
        self.refresh_selection_summary()
        return True

    def _set_profile_image_parameters(self, params):
        index = self.combo_image_dpi.findData(params["dpi"])
        self.combo_image_dpi.setCurrentIndex(index if index >= 0 else self.combo_image_dpi.findData(-1))
        self.spin_image_dpi.setValue(params["dpi"])
        self.spin_image_quality.setValue(params["quality"])

    def import_processing_profile(self):
        path, _filter = QFileDialog.getOpenFileName(self, "导入处理配置", "", "JSON 配置 (*.json)")
        if not path:
            return
        try:
            profile = profiles.loads_profile(Path(path).read_text(encoding="utf-8-sig"))
            self._validate_profile_selection(profile)
            diff = profiles.diff_profiles(self.get_processing_profile(), profile)
            added = ", ".join(diff["added_rules"]) or "无"
            removed = ", ".join(diff["removed_rules"]) or "无"
            changes = "\n".join(f"{key}: {values['before']} -> {values['after']}"
                                for key, values in diff["parameter_changes"].items()) or "无"
            message = (f"配置：{profile['profile_id']} / {profile['profile_version']}\n"
                       f"机构/地区：{profile['agency_or_region']}（不代表官方认证）\n"
                       f"增加规则：{added}\n移除规则：{removed}\n参数变化：\n{changes}\n"
                       f"验证级别：{profile['validation_level']}\n确认应用？")
            if self.show_confirm_message("配置差异预览", message):
                self.apply_processing_profile(profile)
        except (OSError, ValueError, UnicodeError):
            self.show_warning_message("导入失败", "配置文件无效、包含不支持的规则组合或无法读取；未应用配置。")

    def export_processing_profile(self):
        try:
            text = profiles.dumps_profile(self.get_processing_profile())
            path, _filter = QFileDialog.getSaveFileName(self, "导出处理配置", "processing-profile.json", "JSON 配置 (*.json)")
            if not path:
                return
            # Exclusive creation preserves any existing file even when the native dialog permits overwrite.
            with open(path, "x", encoding="utf-8") as stream:
                stream.write(text)
            self.show_success_message("配置已导出", "仅导出处理规则、参数与版本，不包含输出目录或个人设置。")
        except (OSError, ValueError):
            self.show_warning_message("导出失败", "当前配置无效或目标已存在/不可写，请选择新文件名。")

    def on_checkbox_toggled(self, _checked):
        if self.is_applying_preset:
            return

        sender_cb = self.sender()
        if sender_cb not in [self.all_checkboxes.get(key) for key in self.GENERAL_SETTING_OPTIONS]:
            self.active_processing_profile = None
        if isinstance(sender_cb, QCheckBox):
            self.persist_current_checkbox(sender_cb)

        if self.active_preset_key is not None:
            self.active_preset_key = None
            self._set_preset_button_state(None)

        self.custom_selection_before_preset = set(self.get_selected_options())
        self.refresh_selection_summary()

    def show_about_dialog(self):
        if not hasattr(self, 'about_dialog'):
            self.about_dialog = AboutDialog(self)
        self.about_dialog.show()
        self.about_dialog.raise_()
        self.about_dialog.activateWindow()

    def switch_settings_page(self, index):
        # ⚠️使用原生的显示/隐藏逻辑：隐藏的 QWidget 在 QVBoxLayout 中绝对不占任何空间参与高度计算
        self.current_module_index = index
        for i, page in enumerate(self.settings_pages):
            if i == index:
                page.show()
            else:
                page.hide()
        if hasattr(self, "rh_title") and 0 <= index < len(self.MODULES_DATA):
            self.rh_title.setText(self.MODULES_DATA[index]["title"])
        self._refresh_module_indicators()

    def get_selected_options(self):
        selected = []
        for opt_id, cb in self.all_checkboxes.items():
            if cb.isChecked():
                selected.append(opt_id)
        return selected

    def get_compression_settings(self):
        """图像压缩参数（dpi/quality）。

        pdf 层不读取 UI 配置：由控制器在启动处理时取走本方法的返回值，
        经 worker 传给 process_document。非法值由 config.compression 的
        normalize 归一化，这里原样透出即可。
        """
        dpi = DEFAULT_IMAGE_DPI
        if hasattr(self, "combo_image_dpi"):
            data = self.combo_image_dpi.currentData()
            if data == -1 and hasattr(self, "spin_image_dpi"):
                dpi = self.spin_image_dpi.value()
            elif data and data > 0:
                dpi = data
        elif hasattr(self, "app_settings"):
            dpi = self.app_settings.value("Compression/ImageDPI", DEFAULT_IMAGE_DPI)

        return {
            "dpi": dpi,
            "quality": self.spin_image_quality.value() if hasattr(self, "spin_image_quality") else DEFAULT_JPEG_QUALITY,
        }

    def clear_tree_ui(self):
        self.tree.clear()

    def update_counters_ui(self, count):
        self.current_file_count = count
        self.list_title.setText(f"待处理队列 ({count})")
        if count == 0:
            self.queue_meta_label.setText("当前队列为空")
        else:
            self.queue_meta_label.setText("可继续拖入追加")
        queue_stack = getattr(self, "queue_stack", None)
        if queue_stack is not None:
            queue_stack.setCurrentIndex(1 if count > 0 else 0)
        for attr in ("btn_add_files_bar", "btn_add_folder_bar"):
            btn = getattr(self, attr, None)
            if btn is not None:
                btn.setVisible(count > 0)
        if hasattr(self, "btn_precheck") and self.btn_precheck.property("precheckResultCurrent") is True:
            self.btn_precheck.setProperty("precheckResultCurrent", False)
            self.btn_precheck.show()
        self.refresh_selection_summary()

    def set_task_busy(self, task_kind, busy):
        tasks = set(getattr(self, "_busy_tasks", set()))
        if busy:
            tasks.add(task_kind)
        else:
            tasks.discard(task_kind)
        self._busy_tasks = tasks
        self.refresh_selection_summary()

    def refresh_selection_summary(self):
        selected_count = len([
            opt for opt in self.get_selected_options()
            if opt not in self.GENERAL_SETTING_OPTIONS
        ])
        preset_titles = {key: value["title"] for key, value in self.PRESET_OPTIONS.items()}
        preset_titles["favorite"] = "我的常用"
        preset_key = self.active_preset_key if isinstance(self.active_preset_key, str) else ""
        preset_text = preset_titles.get(preset_key, "自定义选择")
        total_files = self.current_file_count

        if selected_count == 0:
            self.selection_summary_label.setText("未选规则")
        else:
            self.selection_summary_label.setText(f"已选 {selected_count} 条")

        if self.active_preset_key:
            self.preset_summary_label.setText(f"当前已应用 {preset_text} 预设，可继续手动微调。")
        else:
            self.preset_summary_label.setText("当前为自定义规则组合，可随时切换 eCTD 预设。")

        self.info_label.setText(f"{total_files} 个文件 · {selected_count} 条规则 · {preset_text}")

        is_prechecking = hasattr(self, "btn_precheck") and self.btn_precheck.property("precheckMode") is True
        busy_tasks = getattr(self, "_busy_tasks", set())
        is_processing = self.btn_start.property("stopMode") is True
        other_busy = bool(busy_tasks - {"processing"})
        # The result signal may precede thread completion. Keep the controls
        # locked through that final interval as well.
        finishing_processing = "processing" in busy_tasks and not is_processing

        if is_prechecking or other_busy or finishing_processing:
            busy_hint = "任务进行中，请稍候"
            self.btn_start.setEnabled(False)
            self.btn_start.setToolTip(busy_hint)
            self.btn_precheck.setEnabled(False)
            self.btn_precheck.setToolTip(busy_hint)
            if hasattr(self, "btn_retry_failed"):
                self.btn_retry_failed.setEnabled(False)
                self.btn_retry_failed.hide()
                self.btn_retry_failed.setToolTip(busy_hint)
            if hasattr(self, "btn_apply_precheck"):
                self.btn_apply_precheck.setEnabled(False)
                self.btn_apply_precheck.hide()
                self.btn_apply_precheck.setToolTip(busy_hint)
            if hasattr(self, "btn_process_precheck_suggested"):
                self.btn_process_precheck_suggested.setEnabled(False)
                self.btn_process_precheck_suggested.hide()
                self.btn_process_precheck_suggested.setToolTip(busy_hint)
        elif self.btn_start.property("stopMode") is not True:
            can_start = (total_files > 0 and selected_count > 0)
            self.btn_start.setEnabled(can_start)
            if total_files == 0:
                self.btn_start.setToolTip("请先添加至少一个 PDF 文件")
            elif selected_count == 0:
                self.btn_start.setToolTip("请至少勾选一条处理规则")
            else:
                self.btn_start.setToolTip("")

            if hasattr(self, "btn_precheck"):
                self.btn_precheck.setEnabled(total_files > 0)
                if total_files == 0:
                    self.btn_precheck.setToolTip("请先添加至少一个 PDF 文件")
                else:
                    self.btn_precheck.setToolTip("扫描队列中文档状态，并提示建议勾选的处理项目")
            if hasattr(self, "btn_retry_failed"):
                has_failed = self.btn_retry_failed.property("hasFailedItems") is True
                self.btn_retry_failed.setVisible(has_failed)
                self.btn_retry_failed.setEnabled(total_files > 0 and selected_count > 0 and has_failed)
                if not has_failed:
                    self.btn_retry_failed.setToolTip("当前没有可重试的失败项")
                elif selected_count == 0:
                    self.btn_retry_failed.setToolTip("请至少勾选一条处理规则")
                else:
                    self.btn_retry_failed.setToolTip("仅重新处理上一轮失败的文件")
            if hasattr(self, "btn_apply_precheck"):
                has_precheck_suggestions = self.btn_apply_precheck.property("hasPrecheckSuggestions") is True
                self.btn_apply_precheck.setVisible(has_precheck_suggestions)
                self.btn_apply_precheck.setEnabled(has_precheck_suggestions)
                if has_precheck_suggestions:
                    self.btn_apply_precheck.setToolTip("自动勾选最近一次预检建议的处理规则")
                else:
                    self.btn_apply_precheck.setToolTip("请先执行一次包含建议项的预检")
            if hasattr(self, "btn_process_precheck_suggested"):
                has_precheck_suggested_files = self.btn_process_precheck_suggested.property("hasPrecheckSuggestedFiles") is True
                self.btn_process_precheck_suggested.setVisible(has_precheck_suggested_files)
                self.btn_process_precheck_suggested.setEnabled(total_files > 0 and selected_count > 0 and has_precheck_suggested_files)
                if not has_precheck_suggested_files:
                    self.btn_process_precheck_suggested.setToolTip("请先执行一次包含建议项的预检")
                elif selected_count == 0:
                    self.btn_process_precheck_suggested.setToolTip("请至少勾选一条处理规则")
                else:
                    self.btn_process_precheck_suggested.setToolTip("仅处理最近一次预检中存在建议项的文件")
        elif hasattr(self, "btn_precheck"):
            self.btn_precheck.setEnabled(False)
            self.btn_precheck.setToolTip("处理中无法执行预检")
            if hasattr(self, "btn_retry_failed"):
                self.btn_retry_failed.setEnabled(False)
                self.btn_retry_failed.hide()
                self.btn_retry_failed.setToolTip("处理中无法重试失败项")
            if hasattr(self, "btn_apply_precheck"):
                self.btn_apply_precheck.setEnabled(False)
                self.btn_apply_precheck.hide()
                self.btn_apply_precheck.setToolTip("处理中无法应用预检建议")
            if hasattr(self, "btn_process_precheck_suggested"):
                self.btn_process_precheck_suggested.setEnabled(False)
                self.btn_process_precheck_suggested.hide()
                self.btn_process_precheck_suggested.setToolTip("处理中无法仅处理建议文件")

        busy = bool(busy_tasks) or is_prechecking or is_processing
        for name in ("btn_bookmark_io_wizard", "btn_link_io_wizard",
                     "btn_detect_annotations", "btn_detect_broken_refs", "btn_clear"):
            btn = getattr(self, name, None)
            if btn is not None:
                btn.setEnabled(total_files > 0 and not busy)
        add_blocked = is_prechecking or other_busy or finishing_processing
        for name in ("btn_add_files", "add_folder_btn", "drop_zone",
                     "btn_add_files_bar", "btn_add_folder_bar"):
            btn = getattr(self, name, None)
            if btn is not None:
                btn.setEnabled(not add_blocked)

        overwrite_cb = self.all_checkboxes.get("覆盖原始文件 (不推荐)")
        if overwrite_cb and overwrite_cb.isChecked():
            self.risk_hint_label.setText("已启用覆盖原文件，执行前请确认已有备份。")
            self.risk_hint_label.setProperty("danger", True)
            if hasattr(self.risk_hint_label, "setVisible"):
                self.risk_hint_label.setVisible(True)
        else:
            self.risk_hint_label.setText("")
            self.risk_hint_label.setProperty("danger", False)
            if hasattr(self.risk_hint_label, "setVisible"):
                self.risk_hint_label.setVisible(False)

        self.style().unpolish(self.risk_hint_label)
        self.style().polish(self.risk_hint_label)
        if hasattr(self, "_refresh_module_indicators"):
            self._refresh_module_indicators()

    def _create_section_label(self, text):
        lbl = QLabel(text)
        lbl.setObjectName("sectionCaption")
        return lbl

    def _create_checkbox(self, opt_id, title, desc, checked):
        display_title = title.replace(" ⚠️", "") if opt_id == "compress_images" else title
        tag = self.RULE_TAGS.get(opt_id)
        card = RuleOptionCard(display_title, desc, checked, tag=tag)
        cb = card.checkbox
        cb.toggled.connect(self.on_checkbox_toggled)
        self.all_checkboxes[opt_id] = cb
        return card

    def _build_compression_card(self):
        """构建'文件体积瘦身与压缩'专属卡片（无损分级分段控件 + 图像内联参数）。"""
        card = QFrame()
        card.setObjectName("compressionCard")
        v_card = QVBoxLayout(card)
        v_card.setContentsMargins(12, 10, 12, 10)
        v_card.setSpacing(8)

        title_lbl = QLabel("文件体积瘦身与压缩")
        title_lbl.setObjectName("compressionCardTitle")
        v_card.addWidget(title_lbl)

        lbl_lossless = QLabel("无损压缩级别")
        lbl_lossless.setObjectName("compressionCardSection")
        v_card.addWidget(lbl_lossless)

        seg_frame = QFrame()
        seg_frame.setObjectName("segmentedContainer")
        seg_frame.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        seg_layout = QHBoxLayout(seg_frame)
        seg_layout.setContentsMargins(3, 3, 3, 3)
        seg_layout.setSpacing(3)

        self.compress_level_group = QButtonGroup(self)
        self.rb_compress_none = QPushButton("不压缩")
        self.rb_compress_standard = QPushButton("智能无损")
        self.rb_compress_aggressive = QPushButton("极限无损")
        for btn in (self.rb_compress_none, self.rb_compress_standard, self.rb_compress_aggressive):
            btn.setCheckable(True)
            btn.setObjectName("segmentedBtn")
            btn.setCursor(Qt.PointingHandCursor)
            btn.setFocusPolicy(Qt.NoFocus)
            seg_layout.addWidget(btn, 1)

        self.rb_compress_none.setChecked(True)
        self.compress_level_group.addButton(self.rb_compress_none, 0)
        self.compress_level_group.addButton(self.rb_compress_standard, 1)
        self.compress_level_group.addButton(self.rb_compress_aggressive, 2)
        v_card.addWidget(seg_frame)

        self.lbl_compress_hint = QLabel("保持原始文档字节流与对象组织不变。")
        self.lbl_compress_hint.setObjectName("compressionCardHint")
        self.lbl_compress_hint.setWordWrap(True)
        self.lbl_compress_hint.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Minimum)
        v_card.addWidget(self.lbl_compress_hint)

        # 内部虚拟复选框（注册进 self.all_checkboxes，与底层处理管线和预检逻辑无缝映射）
        self.cb_compress_standard = QCheckBox()
        self.cb_compress_aggressive = QCheckBox()
        self.cb_generate_object_streams = QCheckBox()
        self.cb_realloc_flate = QCheckBox()

        self.all_checkboxes["compress_standard"] = self.cb_compress_standard
        self.all_checkboxes["compress_aggressive"] = self.cb_compress_aggressive
        self.all_checkboxes["generate_object_streams"] = self.cb_generate_object_streams
        self.all_checkboxes["realloc_flate"] = self.cb_realloc_flate

        for cb in [self.cb_compress_standard, self.cb_compress_aggressive,
                   self.cb_generate_object_streams, self.cb_realloc_flate]:
            cb.toggled.connect(self.on_checkbox_toggled)

        self.compress_level_group.idToggled.connect(self._on_compress_level_changed)

        # 分割线
        sep = QFrame()
        sep.setObjectName("compressionCardSep")
        sep.setFixedHeight(1)
        v_card.addWidget(sep)

        # 图像压缩部分
        img_header = QHBoxLayout()
        img_header.setContentsMargins(0, 0, 0, 0)
        img_header.setSpacing(6)
        self.cb_compress_images = QCheckBox("压缩内嵌图像 (扫描件降采样)")
        self.all_checkboxes["compress_images"] = self.cb_compress_images
        self.cb_compress_images.toggled.connect(self.on_checkbox_toggled)
        self.cb_compress_images.toggled.connect(lambda _c: self.refresh_selection_summary())
        img_tag = QLabel("有损")
        img_tag.setObjectName("ruleTag")
        img_tag.setProperty("tone", "warning")
        img_header.addWidget(self.cb_compress_images, 1)
        img_header.addWidget(img_tag, 0)
        v_card.addLayout(img_header)

        lbl_img_desc = QLabel("将图像降采样至目标 DPI，显著减小扫描类 PDF 体积。")
        lbl_img_desc.setObjectName("compressionCardHint")
        lbl_img_desc.setWordWrap(True)
        lbl_img_desc.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Minimum)
        v_card.addWidget(lbl_img_desc)

        self.dpi_container = QWidget()
        v_dpi = QVBoxLayout(self.dpi_container)
        v_dpi.setContentsMargins(0, 0, 0, 0)
        v_dpi.setSpacing(6)

        row_dpi = QHBoxLayout()
        row_dpi.setSpacing(6)
        lbl_target_dpi = QLabel("目标清晰度：")
        lbl_target_dpi.setObjectName("compressionCardSection")
        row_dpi.addWidget(lbl_target_dpi)

        self.combo_image_dpi = QComboBox()
        self.combo_image_dpi.setFixedHeight(28)
        self.combo_image_dpi.addItem("150 DPI (普通扫描件/黑白)", 150)
        self.combo_image_dpi.addItem("200 DPI (图文平衡)", 200)
        self.combo_image_dpi.addItem("300 DPI (推荐：高清图表)", 300)
        self.combo_image_dpi.addItem("自定义 DPI...", -1)
        row_dpi.addWidget(self.combo_image_dpi, stretch=1)
        v_dpi.addLayout(row_dpi)

        row_custom = QHBoxLayout()
        row_custom.setSpacing(6)
        lbl_custom = QLabel("自定义数值：")
        lbl_custom.setObjectName("compressionCardHint")
        self.spin_image_dpi = QSpinBox()
        self.spin_image_dpi.setFixedHeight(28)
        self.spin_image_dpi.setRange(MIN_IMAGE_DPI, MAX_IMAGE_DPI)
        self.spin_image_dpi.setValue(DEFAULT_IMAGE_DPI)
        self.spin_image_dpi.setSuffix(" DPI")
        row_custom.addWidget(lbl_custom)
        row_custom.addWidget(self.spin_image_dpi, stretch=1)
        self.row_custom_widget = QWidget()
        self.row_custom_widget.setLayout(row_custom)
        self.row_custom_widget.setVisible(False)
        v_dpi.addWidget(self.row_custom_widget)

        quality_row = QHBoxLayout()
        quality_lbl = QLabel("JPEG 质量：")
        quality_lbl.setObjectName("compressionCardSection")
        quality_row.addWidget(quality_lbl)
        self.spin_image_quality = QSpinBox()
        self.spin_image_quality.setRange(MIN_JPEG_QUALITY, MAX_JPEG_QUALITY)
        self.spin_image_quality.setValue(DEFAULT_JPEG_QUALITY)
        quality_row.addWidget(self.spin_image_quality, stretch=1)
        v_dpi.addLayout(quality_row)
        self.spin_image_quality.valueChanged.connect(self._on_image_quality_changed)

        self.dpi_container.setVisible(False)
        v_card.addWidget(self.dpi_container)

        self.cb_compress_images.toggled.connect(self.dpi_container.setVisible)
        self.combo_image_dpi.currentIndexChanged.connect(self._on_image_dpi_combo_changed)
        self.spin_image_dpi.valueChanged.connect(self._on_image_dpi_spin_changed)

        return card

    def _on_compress_level_changed(self, button_id, checked):
        if not checked:
            return
        if button_id == 0:  # 不压缩
            self.lbl_compress_hint.setText("保持原始文档字节流与对象组织不变。")
            self.cb_compress_standard.setChecked(False)
            self.cb_compress_aggressive.setChecked(False)
            self.cb_generate_object_streams.setChecked(False)
            self.cb_realloc_flate.setChecked(False)
        elif button_id == 1:  # 智能无损瘦身
            self.lbl_compress_hint.setText("融合对象流打包、Flate 最优重压缩与垃圾回收，100% 保真。")
            self.cb_compress_standard.setChecked(True)
            self.cb_compress_aggressive.setChecked(False)
            self.cb_generate_object_streams.setChecked(True)
            self.cb_realloc_flate.setChecked(True)
        elif button_id == 2:  # 极限无损瘦身
            self.lbl_compress_hint.setText("最大化清理对象与结构碎片 (garbage=4 + clean)，适用于超限文件。")
            self.cb_compress_standard.setChecked(False)
            self.cb_compress_aggressive.setChecked(True)
            self.cb_generate_object_streams.setChecked(True)
            self.cb_realloc_flate.setChecked(True)

        if not self.is_applying_preset:
            self.active_processing_profile = None
            self.active_preset_key = None
            self._set_preset_button_state(None)
            self.persist_all_settings()
            self.refresh_selection_summary()

    def _on_image_dpi_combo_changed(self, _index):
        val = self.combo_image_dpi.currentData()
        if val == -1:
            if hasattr(self, "row_custom_widget"):
                self.row_custom_widget.setVisible(True)
            dpi = self.spin_image_dpi.value()
        else:
            if hasattr(self, "row_custom_widget"):
                self.row_custom_widget.setVisible(False)
            dpi = val
        if hasattr(self, "app_settings") and not self.is_applying_preset:
            self.active_processing_profile = None
            self.app_settings.setValue("Compression/ImageDPI", dpi)

    def _on_image_dpi_spin_changed(self, value):
        if self.combo_image_dpi.currentData() == -1:
            if hasattr(self, "app_settings") and not self.is_applying_preset:
                self.active_processing_profile = None
                self.app_settings.setValue("Compression/ImageDPI", value)

    def _on_image_quality_changed(self, value):
        if hasattr(self, "app_settings") and not self.is_applying_preset:
            self.active_processing_profile = None
            self.app_settings.setValue("Compression/ImageQuality", value)

    def _on_compress_images_toggled(self, checked):
        pass

    def apply_stylesheet(self):
        # 视觉样式全部集中在 theme.py 的应用级 QSS 中，由 ThemeManager 应用到
        # QApplication 并级联到本窗口。此处保留方法名以兼容旧调用点：主题未初始化
        # (如单元测试直接实例化 MainWindow) 时，回退到用当前生效配色本地渲染一次，
        # 确保窗口不至于完全无样式。
        app = QApplication.instance()
        if app is not None and app.styleSheet():
            return
        self.setStyleSheet(build_app_qss(active_palette()))
