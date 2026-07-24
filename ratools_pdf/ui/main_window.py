import os

from PySide6.QtCore import QSettings, Qt, QTimer
from PySide6.QtGui import QIcon
from PySide6.QtWidgets import (
    QAbstractSpinBox, QApplication, QButtonGroup, QCheckBox, QFileDialog, QFrame, QHBoxLayout,
    QHeaderView, QLabel, QMainWindow, QPushButton, QRadioButton, QScrollArea, QSizePolicy, QTreeWidget,
    QTreeWidgetItem, QVBoxLayout, QWidget,
)

from ratools_pdf.config.features import ENABLE_UPDATE_CHECK
from ratools_pdf.config.paths import get_app_dir, get_resource_path
from ratools_pdf.ui.dialogs import (
    AboutDialog,
    CustomMessageBox,
    FramelessDraggableDialog,
    IODataWizardDialog,
    LogDialog,
    ManualFontEmbeddingDialog,
    SettingsDialog,
)
from ratools_pdf.ui.theme import (
    active_palette,
    apply_windows_title_bar_theme,
    build_app_qss,
    ThemeManager,
)
from ratools_pdf.ui.widgets import DropZoneLabel


class MainWindow(QMainWindow):
    PRESET_OPTIONS = {
        "china": {
            "title": "中国 eCTD",
            "options": {
                "convert_pdf_version",
                "fast_web_view",
                "initial_view_bookmarks_and_page",
                "page_layout_default",
                "open_page_first",
                "bookmark_inherit_zoom",
                "cleanup_remove_external_uri",
                "cleanup_remove_annotations",
                "cleanup_remove_metadata",
                "cleanup_remove_attachments",
                "cleanup_remove_dynamic_content",
                "link_inherit_zoom",
                "link_open_new_window",
                "bookmark_open_new_window",
                "collapse_all_bookmarks",
                "filename_ectd_format",
            },
        },
        "us": {
            "title": "美国 eCTD",
            "options": {
                "convert_pdf_version",
                "fast_web_view",
                "initial_view_bookmarks_and_page",
                "page_layout_default",
                "open_page_first",
                "bookmark_inherit_zoom",
                "cleanup_remove_annotations",
                "cleanup_remove_metadata",
                "cleanup_remove_attachments",
                "cleanup_remove_dynamic_content",
                "link_inherit_zoom",
                "link_open_new_window",
                "bookmark_open_new_window",
            },
        },
    }

    def __init__(self):
        super().__init__()
        self.custom_selection_before_preset = set()
        self.active_preset_key = None
        self.is_applying_preset = False
        self.setWindowTitle("RATools for PDF")

        # === 添加原生窗口图标 ===
        self.setWindowIcon(QIcon(get_resource_path("icon.ico")))

        self.resize(1100, 750)
        self.setMinimumSize(900, 600)

        self.all_checkboxes = {}
        self.current_file_count = 0

        # ================= 初始化 QSettings 与主题管理 =================
        current_dir = get_app_dir()
        ini_path = os.path.join(current_dir, "settings.ini")
        self.app_settings = QSettings(ini_path, QSettings.IniFormat)

        # 主题必须先应用到 QApplication，再让窗口级 apply_stylesheet() 做兼容回退；
        # 否则启动时会给 MainWindow 残留一份本地亮色 QSS，压过应用级暗色主题。
        app = QApplication.instance()
        theme_mode = str(self.app_settings.value("Settings/ThemeMode", "system") or "system").lower()
        if app is not None:
            self.theme_manager = ThemeManager(app, theme_mode)
            self.theme_manager.changed.connect(self.apply_native_title_bar_theme)
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

        self.MODULES_DATA = [
            {
                "icon": "👀",
                "title": "初始视图与文档属性",
                "options": [
                    {"id": "open_page_first", "title": "设为首页打开", "desc": "强制文档打开时默认显示第一页"},
                    {"id": "page_layout_default", "title": "重置页面布局", "desc": "将页面布局恢复为默认"},
                    {"id": "zoom_default", "title": "重置缩放比例", "desc": "将打开时的缩放比例设置为默认"},
                    {"id": "initial_view_bookmarks_and_page", "title": "设置导览标签", "desc": "包含书签的文档，导览标签设置为书签面板和页面；不包含书签的文档，导览标签设置为页面。"},
                    {"id": "collapse_all_bookmarks", "title": "折叠所有书签", "desc": "将书签树默认设置为折叠状态，保持界面整洁"},
                    {"id": "title_from_filename", "title": "同步文件名为标题", "desc": "自动将当前PDF的文件名写入文档属性的“标题”元数据中"}
                ]
            },
            {
                "icon": "📄",
                "title": "页面与字体标准化",
                "options": [
                    {"id": "page_size_a4", "title": "适配到A4尺寸", "desc": "无法通过预检精确判断是否需要处理；智能模式下若勾选仍会执行"},
                    {"id": "page_size_letter", "title": "适配到Letter尺寸", "desc": "按原页面方向等比缩放并居中留白，适配到Letter (信纸) 尺寸，尽量保留全部内容"}
                ]
            },
            {
                "icon": "🔖",
                "title": "书签管理",
                "options": [
                    {"id": "bookmark_inherit_zoom", "title": "书签设为承前缩放", "desc": "点击书签跳转时，保持当前页面的缩放比例不变 (Inherit Zoom)"},
                    {"id": "bookmark_open_new_window", "title": "书签动作：新窗口打开", "desc": "配置书签的链接跳转默认在新的PDF浏览器窗口中打开"},
                    {"id": "bookmark_remove_external_links", "title": "清理书签外部链接", "desc": "移除书签中指向网页或外部文件的URI动作"},
                    {"id": "bookmark_remove_invalid", "title": "清理失效书签", "desc": "自动检测并删除未指向任何有效页面或动作的空书签"},
                    {"id": "bookmark_remove_unknown_actions", "title": "清理非标准动作书签", "desc": "仅保留内部跳转、外部文档和调用命令，删除其它未知动作"}
                ]
            },
            {
                "icon": "🔗",
                "title": "超链接处理",
                "options": [
                    {"id": "link_abs_to_rel_path", "title": "绝对路径转相对路径", "desc": "将外部文件链接的绝对路径自动转换为相对路径"},
                    {"id": "link_inherit_zoom", "title": "超链接设为承前缩放", "desc": "点击链接跳转时，保持当前屏幕的视图缩放比例 (Inherit Zoom)"},
                    {"id": "link_open_new_window", "title": "链接动作：新窗口打开", "desc": "强制外部文档或网页链接在独立的新窗口中打开"},
                    {"id": "link_text_blue", "title": "链接文本设为蓝色", "desc": "自动识别超链接区域并将其文本颜色变更为标准蓝色"},
                    {"id": "link_black_border", "title": "链接区域加黑框", "desc": "为所有的有效超链接区域添加1px的黑色实线边框"},
                    {"id": "link_bordered_to_blue_border", "title": "标准化有框链接", "desc": "若超链接已存在边框，则统一转为蓝框黑字样式"},
                    {"id": "link_unbordered_blue_to_blue_border", "title": "标准化无框蓝字链接", "desc": "若超链接无边框且文字为蓝色，则统一转为蓝框黑字样式"},
                    {"id": "link_remove_border", "title": "清除所有链接边框", "desc": "移除文档内所有超链接的可见边框，保持页面排版干净"}
                ]
            },
            {
                "icon": "🛡️",
                "title": "内容合规与安全性",
                "options": [
                    {"id": "cleanup_remove_external_uri", "title": "删除外部URI链接", "desc": "清理指向外部网站、邮箱等所有URI类型的超链接"},
                    {"id": "cleanup_remove_external_uri_and_text_black", "title": "删除外部URI链接并去色", "desc": "清理URI链接的同时，将该链接对应的文本颜色重置为黑色"},
                    {"id": "cleanup_remove_invalid_links", "title": "清理失效超链接", "desc": "自动扫描并移除所有未分配有效动作 (Action) 的空链接"},
                    {"id": "cleanup_remove_invalid_links_and_text_black", "title": "清理失效链接并去色", "desc": "移除空链接，并将该区域相关的文本颜色恢复为普通黑色"},
                    {"id": "cleanup_remove_unknown_action_links", "title": "清理非标准动作链接", "desc": "仅保留内部/外部跳转和执行动作，移除其它所有的特殊行为"},
                    {"id": "cleanup_remove_dynamic_content", "title": "彻底清除动态内容 (JS/3D)", "desc": "删除文档内所有的JavaScript脚本、3D模型等交互元素以满足安全合规"},
                    {"id": "cleanup_remove_attachments", "title": "移除所有内嵌附件", "desc": "清理PDF内部打包的所有附加文件 (.zip, .xml 等)"},
                    {"id": "cleanup_remove_tags", "title": "移除结构化标签", "desc": "删除PDF结构树 (StructTreeRoot) 和标记信息 (MarkInfo)"},
                    {"id": "cleanup_remove_annotations", "title": "清理所有高亮/批注", "desc": "删除文本框、高亮、画笔等所有非链接类型的交互式注释"},
                    {"id": "cleanup_remove_metadata", "title": "清空文档元数据", "desc": "移除所有标题、作者、创建时间等PieceInfo和Metadata"},
                    {"id": "cleanup_remove_all_links_bookmarks", "title": "移除全部链接和书签", "desc": "仅删除页面链接与书签，不删除普通批注"}
                ]
            },
            {
                "icon": "📦",
                "title": "文件级优化与输出",
                "options": [
                    {"id": "convert_pdf_version", "title": "PDF版本转换", "desc": "将PDF版本修改为1.7版本"},
                    {"id": "remove_pdf_restrictions", "title": "PDF解除权限限制", "desc": "尝试移除禁止复制、打印、编辑等权限限制，不处理需要打开密码的加密文档"},
                    {"id": "fast_web_view", "title": "启用线性化 (快速网页浏览)", "desc": "优化文档结构以支持Web环境下的流式加载和边下边看"},
                    {"id": "filename_ectd_format", "title": "eCTD文件名合规格式化", "desc": "自动将输出文件名转为小写、去除空格并替换非法字符"}
                ]
            }
        ]

        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        # ================= 顶部 Header =================
        header = QFrame()
        header.setObjectName("header")
        header.setFixedHeight(56)
        header_layout = QHBoxLayout(header)
        header_layout.setContentsMargins(24, 0, 24, 0)

        self.btn_top_settings = QPushButton("⚙️ 全局设置")
        self.btn_top_settings.setObjectName("topBtn")
        self.btn_top_settings.clicked.connect(self.settings_dialog.show)

        self.btn_top_about = QPushButton("ℹ️ 关于")
        self.btn_top_about.setObjectName("topBtn")
        self.btn_top_about.clicked.connect(self.show_about_dialog)

        header_layout.addWidget(self.btn_top_settings)
        header_layout.addWidget(self.btn_top_about)
        header_layout.addStretch()
        main_layout.addWidget(header)

        middle_container = QFrame()
        middle_layout = QHBoxLayout(middle_container)
        middle_layout.setContentsMargins(0, 0, 0, 0)
        middle_layout.setSpacing(0)

        # ================= 左侧导航栏 =================
        left_sidebar = QFrame()
        left_sidebar.setObjectName("leftSidebar")
        left_sidebar.setFixedWidth(256)
        left_layout = QVBoxLayout(left_sidebar)
        left_layout.setContentsMargins(16, 16, 16, 16)
        left_layout.setSpacing(8)
        nav_title = QLabel("功能模块")
        nav_title.setObjectName("navTitle")
        left_layout.addWidget(nav_title)

        self.nav_buttons = []
        self.nav_btn_group = QButtonGroup(self)
        self.nav_btn_group.setExclusive(True)

        for idx, mod in enumerate(self.MODULES_DATA):
            btn = QPushButton(f"{mod['icon']}  {mod['title']}")
            btn.setCheckable(True)
            btn.setObjectName("navBtn")
            self.nav_buttons.append(btn)
            self.nav_btn_group.addButton(btn, idx)
            left_layout.addWidget(btn)

        self.nav_buttons[0].setChecked(True)
        left_layout.addStretch()
        middle_layout.addWidget(left_sidebar)

        # ================= 中间主要视图 =================
        main_view = QFrame()
        main_view.setObjectName("mainView")
        main_view_layout = QVBoxLayout(main_view)
        main_view_layout.setContentsMargins(20, 20, 20, 20)
        main_view_layout.setSpacing(12)

        import_card = QFrame()
        import_card.setObjectName("importCard")
        import_layout = QVBoxLayout(import_card)
        import_layout.setContentsMargins(18, 18, 18, 14)
        import_layout.setSpacing(12)

        import_header = QHBoxLayout()
        import_header.setContentsMargins(0, 0, 0, 0)
        import_header.setSpacing(10)
        import_title = QLabel("导入待处理文件")
        import_title.setObjectName("sectionTitle")
        import_hint = QLabel("支持拖入 PDF 或整个文件夹")
        import_hint.setObjectName("sectionHint")
        import_header.addWidget(import_title)
        import_header.addStretch()
        import_header.addWidget(import_hint)
        import_layout.addLayout(import_header)

        self.drop_zone = DropZoneLabel("拖拽 PDF 到这里\n或点击下方按钮快速添加")
        self.drop_zone.setObjectName("dropZone")
        self.drop_zone.setAlignment(Qt.AlignCenter)
        self.drop_zone.setFixedHeight(96)
        import_layout.addWidget(self.drop_zone)

        quick_actions = QHBoxLayout()
        quick_actions.setContentsMargins(0, 0, 0, 0)
        quick_actions.setSpacing(10)
        self.btn_add_files = QPushButton("选择 PDF 文件")
        self.btn_add_files.setObjectName("secondaryBtn")
        self.add_folder_btn = QPushButton("选择文件夹")
        self.add_folder_btn.setObjectName("secondaryBtn")
        self.queue_meta_label = QLabel("当前队列为空")
        self.queue_meta_label.setObjectName("mutedLabel")
        quick_actions.addWidget(self.btn_add_files)
        quick_actions.addWidget(self.add_folder_btn)
        quick_actions.addSpacing(8)
        quick_actions.addWidget(self.queue_meta_label, 1)
        import_layout.addLayout(quick_actions)

        list_container = QFrame()
        list_container.setObjectName("listContainer")
        list_layout = QVBoxLayout(list_container)
        list_layout.setContentsMargins(0, 0, 0, 0)
        list_layout.setSpacing(0)

        list_header = QFrame()
        list_header.setObjectName("listHeader")
        list_header_layout = QHBoxLayout(list_header)
        list_header_layout.setContentsMargins(16, 8, 16, 8)
        self.list_title = QLabel("待处理队列 (0)")
        self.list_title.setObjectName("listTitle")
        self.btn_clear = QPushButton("清空列表")
        self.btn_clear.setObjectName("actionBtn")
        list_header_layout.addWidget(self.list_title)
        list_header_layout.addSpacing(12)
        list_header_layout.addStretch()
        list_header_layout.addWidget(self.btn_clear)
        list_layout.addWidget(list_header)

        # ================= 文件树视图 =================
        self.tree = QTreeWidget()
        self.tree.setHeaderLabels(["文件 / 文件夹", "绝对路径", "当前状态"])
        self.tree.header().setSectionResizeMode(0, QHeaderView.Interactive)
        self.tree.header().setSectionResizeMode(1, QHeaderView.Stretch)
        self.tree.header().setSectionResizeMode(2, QHeaderView.Fixed)
        self.tree.header().setStretchLastSection(False)
        self.tree.setColumnWidth(0, 280)
        self.tree.setColumnWidth(2, 110)

        self.tree.setSelectionBehavior(QTreeWidget.SelectRows)
        self.tree.setSelectionMode(QTreeWidget.ExtendedSelection)
        self.tree.setContextMenuPolicy(Qt.CustomContextMenu)
        self.tree.setIndentation(14)
        self.tree.setUniformRowHeights(True)
        self.tree.setAllColumnsShowFocus(True)
        self.tree.header().setMinimumSectionSize(90)
        self.tree.header().resizeSection(1, 380)

        self.tree.setAlternatingRowColors(True)
        self.tree.setRootIsDecorated(True)
        list_layout.addWidget(self.tree)
        import_layout.addWidget(list_container, 1)

        main_view_layout.addWidget(import_card, 1)

        middle_layout.addWidget(main_view)

        # ================= 右侧设置区 =================
        right_sidebar = QFrame()
        right_sidebar.setObjectName("rightSidebar")
        right_sidebar.setFixedWidth(320)
        right_layout = QVBoxLayout(right_sidebar)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(0)

        right_header = QFrame()
        right_header.setObjectName("rightHeader")
        rh_layout = QVBoxLayout(right_header)
        rh_layout.setContentsMargins(16, 18, 16, 14)
        rh_layout.setSpacing(18)

        self.rh_title = QLabel("处理规则选项")
        self.rh_title.setObjectName("rightPanelTitle")

        self.selection_summary_label = QLabel("尚未选择任何处理规则")
        self.selection_summary_label.setObjectName("selectionSummary")
        rh_layout.addWidget(self.selection_summary_label)
        rh_layout.addWidget(self.rh_title)
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

        # 存储每一个页面的包装容器
        self.settings_pages = []

        self.btn_bookmark_io_wizard = QPushButton("书签数据导入/导出...")
        self.btn_link_io_wizard = QPushButton("链接数据导入/导出...")
        self.btn_embed_missing_fonts = QPushButton("🛠 嵌入缺失字体")
        self.btn_embed_missing_fonts.setObjectName("secondaryBtn")
        self.btn_embed_missing_fonts.setCursor(Qt.PointingHandCursor)
        self.btn_embed_missing_fonts.setToolTip("打开选中的 PDF，并在 Acrobat 中手动执行印前检查的“嵌入缺失的字体”。")

        for btn in [self.btn_bookmark_io_wizard, self.btn_link_io_wizard]:
            btn.setObjectName("secondaryBtn")
            btn.setCursor(Qt.PointingHandCursor)

        for mod in self.MODULES_DATA:
            page = QWidget()
            page_layout = QVBoxLayout(page)
            page_layout.setContentsMargins(20, 18, 20, 20)
            page_layout.setSpacing(14)

            for opt in mod["options"]:
                page_layout.addWidget(self._create_checkbox(opt["id"], opt["title"], opt["desc"], False))

            if mod["title"] == "页面与字体标准化":
                page_layout.addSpacing(12)
                page_layout.addWidget(self._create_section_label("字体手动处理"))
                font_action_hint = QLabel("如预检提示存在未嵌入字体，请先选中左侧 PDF，再点击下面的按钮跳转 Acrobat 进行处理。")
                font_action_hint.setWordWrap(True)
                font_action_hint.setObjectName("mutedSmall")
                page_layout.addWidget(font_action_hint)
                page_layout.addWidget(self.btn_embed_missing_fonts)

            elif mod["title"] == "书签管理":
                page_layout.addSpacing(12)
                page_layout.addWidget(self._create_section_label("导出/导入书签"))
                btn_layout = QVBoxLayout()
                btn_layout.setSpacing(8)
                btn_layout.addWidget(self.btn_bookmark_io_wizard)
                page_layout.addLayout(btn_layout)

            elif mod["title"] == "超链接处理":
                page_layout.addSpacing(12)
                page_layout.addWidget(self._create_section_label("导出/导入链接"))
                btn_layout = QVBoxLayout()
                btn_layout.setSpacing(8)
                btn_layout.addWidget(self.btn_link_io_wizard)
                page_layout.addLayout(btn_layout)

            page_layout.addStretch()

            # 将每个页面按顺序加入核心布局并暂时隐藏
            self.settings_layout.addWidget(page)
            page.hide()
            self.settings_pages.append(page)

        scroll_area.setWidget(self.settings_container)
        right_layout.addWidget(scroll_area)

        middle_layout.addWidget(right_sidebar)
        main_layout.addWidget(middle_container)

        preset_bar = QFrame()
        preset_bar.setObjectName("presetBar")
        preset_layout = QHBoxLayout(preset_bar)
        preset_layout.setContentsMargins(24, 12, 24, 12)
        preset_layout.setSpacing(10)

        preset_label = QLabel("快速预设")
        preset_label.setObjectName("presetLabel")
        preset_layout.addWidget(preset_label)

        self.btn_preset_china = QPushButton("中国eCTD")
        self.btn_preset_china.setObjectName("presetBtn")
        self.btn_preset_china.setCheckable(True)
        self.btn_preset_china.setFocusPolicy(Qt.NoFocus)
        preset_layout.addWidget(self.btn_preset_china)

        self.btn_preset_us = QPushButton("美国eCTD")
        self.btn_preset_us.setObjectName("presetBtn")
        self.btn_preset_us.setCheckable(True)
        self.btn_preset_us.setFocusPolicy(Qt.NoFocus)
        preset_layout.addWidget(self.btn_preset_us)

        self.btn_preset_favorite = QPushButton("我的常用")
        self.btn_preset_favorite.setObjectName("presetBtn")
        self.btn_preset_favorite.setCheckable(True)
        self.btn_preset_favorite.setFocusPolicy(Qt.NoFocus)
        preset_layout.addWidget(self.btn_preset_favorite)

        self.btn_clear_selected_options = QPushButton("全部取消")
        self.btn_clear_selected_options.setObjectName("actionBtn")
        self.btn_save_favorite_preset = QPushButton("保存为常用")
        self.btn_save_favorite_preset.setObjectName("actionBtn")

        self.preset_summary_label = QLabel("默认载入中国 eCTD 预设，可按需微调。")
        self.preset_summary_label.setObjectName("presetSummary")

        self.preset_btn_group = QButtonGroup(self)
        self.preset_btn_group.setExclusive(True)
        self.preset_btn_group.addButton(self.btn_preset_china)
        self.preset_btn_group.addButton(self.btn_preset_us)
        self.preset_btn_group.addButton(self.btn_preset_favorite)

        preset_layout.addSpacing(8)
        preset_layout.addWidget(self.preset_summary_label)
        preset_layout.addStretch()
        preset_layout.addWidget(self.btn_save_favorite_preset)
        preset_layout.addWidget(self.btn_clear_selected_options)
        main_layout.addWidget(preset_bar)

        # ================= 底部操作栏 =================
        footer = QFrame()
        footer.setObjectName("footer")
        footer.setFixedHeight(64)
        footer_layout = QHBoxLayout(footer)
        footer_layout.setContentsMargins(24, 0, 24, 0)

        self.btn_log = QPushButton("📋 查看/导出日志")
        self.btn_log.setObjectName("actionBtn")
        footer_layout.addWidget(self.btn_log)
        footer_layout.addStretch()

        self.info_label = QLabel("0 个文件 · 0 条规则 · 中国 eCTD 预设")
        self.info_label.setObjectName("footerSummary")
        self.processing_hint_label = QLabel("")
        self.processing_hint_label.setObjectName("processingHint")
        self.risk_hint_label = QLabel("")
        self.risk_hint_label.setObjectName("footerHint")
        self.processing_mode_group = QButtonGroup(self)
        self.radio_smart_processing = QRadioButton("智能处理")
        self.radio_smart_processing.setObjectName("modeRadio")
        self.radio_smart_processing.setToolTip("处理前先按已勾选规则进行预检，仅处理预检发现需要修改的规则，速度更快。无法可靠预检的勾选规则仍会执行。")
        self.radio_force_processing = QRadioButton("全部处理")
        self.radio_force_processing.setObjectName("modeRadio")
        self.radio_force_processing.setToolTip("强制执行全部已勾选规则，不做预检筛选，处理更彻底但耗时更长。")
        self.radio_smart_processing.setChecked(True)
        self.processing_mode_group.addButton(self.radio_smart_processing)
        self.processing_mode_group.addButton(self.radio_force_processing)
        self.btn_skip_current = QPushButton("⏭ 跳过当前文件")
        self.btn_skip_current.setObjectName("actionBtn")
        self.btn_skip_current.setEnabled(False)
        self.btn_skip_current.hide()
        self.btn_retry_failed = QPushButton("↻ 仅处理失败项")
        self.btn_retry_failed.setObjectName("actionBtn")
        self.btn_retry_failed.setEnabled(False)
        self.btn_retry_failed.hide()
        self.btn_apply_precheck = QPushButton("✓ 应用预检建议")
        self.btn_apply_precheck.setObjectName("actionBtn")
        self.btn_apply_precheck.setEnabled(False)
        self.btn_apply_precheck.hide()
        self.btn_process_precheck_suggested = QPushButton("▶ 仅处理建议文件")
        self.btn_process_precheck_suggested.setObjectName("actionBtn")
        self.btn_process_precheck_suggested.setEnabled(False)
        self.btn_process_precheck_suggested.hide()
        self.btn_precheck = QPushButton("🔎 预检")
        self.btn_precheck.setObjectName("actionBtn")
        self.btn_start = QPushButton("▶ 开始处理")
        self.btn_start.setObjectName("startBtn")
        footer_layout.addWidget(self.info_label)
        footer_layout.addSpacing(16)
        footer_layout.addWidget(self.processing_hint_label)
        footer_layout.addSpacing(16)
        footer_layout.addWidget(self.risk_hint_label)
        footer_layout.addSpacing(16)
        footer_layout.addWidget(self.radio_smart_processing)
        footer_layout.addSpacing(8)
        footer_layout.addWidget(self.radio_force_processing)
        footer_layout.addSpacing(16)
        footer_layout.addWidget(self.btn_skip_current)
        footer_layout.addSpacing(10)
        footer_layout.addWidget(self.btn_retry_failed)
        footer_layout.addSpacing(10)
        footer_layout.addWidget(self.btn_apply_precheck)
        footer_layout.addSpacing(10)
        footer_layout.addWidget(self.btn_process_precheck_suggested)
        footer_layout.addSpacing(10)
        footer_layout.addWidget(self.btn_precheck)
        footer_layout.addSpacing(10)
        footer_layout.addWidget(self.btn_start)
        main_layout.addWidget(footer)

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
        self.radio_smart_processing.toggled.connect(lambda _checked: self.persist_processing_mode())
        self.radio_force_processing.toggled.connect(lambda _checked: self.persist_processing_mode())

        self.settings_dialog.set_theme_mode(theme_mode)
        for mode, btn in self.settings_dialog.theme_buttons.items():
            btn.clicked.connect(lambda _checked=False, m=mode: self.on_theme_mode_changed(m))

        self.load_all_settings()
        self.refresh_selection_summary()

    def show_info_message(self, title, message):
        CustomMessageBox(title, message, msg_type="info", parent=self).exec()

    def show_success_message(self, title, message):
        CustomMessageBox(title, message, msg_type="success", parent=self).exec()

    def show_warning_message(self, title, message):
        CustomMessageBox(title, message, msg_type="warning", parent=self).exec()

    def show_error_message(self, title, message):
        CustomMessageBox(title, message, msg_type="error", parent=self).exec()

    def show_manual_font_embedding_dialog(self, pdf_paths, open_callback):
        dlg = ManualFontEmbeddingDialog(pdf_paths, open_callback, parent=self)
        dlg.exec()

    def show_confirm_message(self, title, message):
        dlg = CustomMessageBox(title, message, msg_type="question", show_cancel=True, parent=self)
        return dlg.exec() == QDialog.Accepted

    def show_major_update_prompt(self, current_version, release):
        dlg = FramelessDraggableDialog("重要更新可用", self)
        dlg.resize(460, 300)
        dlg.update_action = "later"
        dlg.content_layout.setSpacing(14)

        message = QLabel(
            f"发现重要更新：{release.version_text}\n"
            f"当前版本：{current_version}\n"
            f"发布标题：{release.title}\n"
            f"发布时间：{release.published_at or '未知'}"
        )
        message.setWordWrap(True)
        message.setObjectName("aboutText")
        dlg.content_layout.addWidget(message)
        dlg.content_layout.addStretch()

        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(10)
        btn_layout.addStretch()

        btn_ignore = QPushButton("忽略此版本")
        btn_ignore.setObjectName("dialogSecondaryBtn")
        btn_later = QPushButton("稍后提醒")
        btn_later.setObjectName("dialogSecondaryBtn")
        btn_open = QPushButton("查看更新")
        btn_open.setObjectName("dialogPrimaryBtn")

        def finish(action):
            dlg.update_action = action
            dlg.accept()

        btn_open.clicked.connect(lambda: finish("open"))
        btn_later.clicked.connect(lambda: finish("later"))
        btn_ignore.clicked.connect(lambda: finish("ignore"))

        btn_layout.addWidget(btn_ignore)
        btn_layout.addWidget(btn_later)
        btn_layout.addWidget(btn_open)
        dlg.content_layout.addLayout(btn_layout)

        dlg.exec()
        return dlg.update_action

    def show_signed_files_prompt(self, signed_files):
        """检测到已签名文件时的三选一提示：跳过 / 全部处理 / 取消。

        返回 "skip"、"process_all" 或 "cancel"。
        """
        count = len(signed_files)

        dlg = FramelessDraggableDialog("检测到已签名文件", self)
        dlg.signed_action = "cancel"
        dlg.content_layout.setSpacing(16)

        # ---- 顶部警告卡片：图标 + 标题 + 说明 ----
        warn_card = QFrame()
        warn_card.setObjectName("signedWarnCard")
        warn_layout = QHBoxLayout(warn_card)
        warn_layout.setContentsMargins(14, 14, 14, 14)
        warn_layout.setSpacing(12)

        warn_icon = QLabel("⚠️")
        warn_icon.setObjectName("signedWarnIcon")
        warn_icon.setAlignment(Qt.AlignTop)
        warn_layout.addWidget(warn_icon, 0, Qt.AlignTop)

        warn_text_layout = QVBoxLayout()
        warn_text_layout.setSpacing(4)
        warn_title = QLabel(f"待处理列表中有 {count} 个文件已包含数字签名")
        warn_title.setWordWrap(True)
        warn_title.setObjectName("signedWarnTitle")
        warn_desc = QLabel("继续处理会重写这些文件，使原有数字签名失效。此操作不可逆。")
        warn_desc.setWordWrap(True)
        warn_desc.setObjectName("signedWarnDesc")
        warn_text_layout.addWidget(warn_title)
        warn_text_layout.addWidget(warn_desc)
        warn_layout.addLayout(warn_text_layout, 1)

        dlg.content_layout.addWidget(warn_card)

        # ---- 文件列表标题 ----
        caption = QLabel("已签名文件：")
        caption.setObjectName("signedListCaption")
        dlg.content_layout.addWidget(caption)

        # ---- 可滚动文件列表（带项目符号）----
        file_list = QLabel("\n".join(f"•  {os.path.basename(p)}" for p in signed_files))
        file_list.setWordWrap(True)
        file_list.setTextInteractionFlags(Qt.TextSelectableByMouse)
        file_list.setAlignment(Qt.AlignTop | Qt.AlignLeft)
        file_list.setObjectName("signedFileList")
        # 悬停显示完整路径，方便定位同名文件
        file_list.setToolTip("\n".join(signed_files))

        scroll = QScrollArea()
        scroll.setObjectName("signedListScroll")
        scroll.setWidgetResizable(True)
        scroll.setWidget(file_list)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        # 单文件时不需要滚动条占位，多文件时限制高度触发滚动
        scroll.setMinimumHeight(56)
        scroll.setMaximumHeight(180 if count > 1 else 72)
        dlg.content_layout.addWidget(scroll)

        # ---- 底部操作提示 ----
        hint = QLabel("请选择如何处理这些已签名文件：")
        hint.setObjectName("dialogMuted")
        dlg.content_layout.addWidget(hint)
        dlg.content_layout.addStretch()

        # ---- 按钮区 ----
        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(10)
        btn_layout.addStretch()

        btn_cancel = QPushButton("取消处理")
        btn_cancel.setObjectName("dialogSecondaryBtn")
        btn_process_all = QPushButton("仍然处理全部")
        btn_process_all.setObjectName("dialogSecondaryBtn")
        btn_skip = QPushButton("跳过已签名文件")
        btn_skip.setObjectName("dialogPrimaryBtn")
        for btn in (btn_cancel, btn_process_all, btn_skip):
            btn.setFixedHeight(36)
            btn.setCursor(Qt.PointingHandCursor)

        def finish(action):
            dlg.signed_action = action
            dlg.accept()

        btn_skip.clicked.connect(lambda: finish("skip"))
        btn_process_all.clicked.connect(lambda: finish("process_all"))
        btn_cancel.clicked.connect(lambda: finish("cancel"))

        btn_layout.addWidget(btn_cancel)
        btn_layout.addWidget(btn_process_all)
        btn_layout.addWidget(btn_skip)
        dlg.content_layout.addLayout(btn_layout)

        dlg.adjustSize()
        dlg.setMinimumWidth(500)
        dlg.exec()
        return dlg.signed_action

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

        processing_mode = str(self.app_settings.value("Settings/ProcessingMode", "smart") or "smart").lower()
        self.radio_force_processing.setChecked(processing_mode == "force")
        self.radio_smart_processing.setChecked(processing_mode != "force")

        self.is_applying_preset = False

        # 默认恢复上次会话的勾选状态（自定义），不自动套用预设
        self.custom_selection_before_preset = set(self.get_selected_options())
        self.active_preset_key = None
        self._set_preset_button_state(None)

    def closeEvent(self, event):
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

    def persist_parallel_settings(self):
        if not hasattr(self, "app_settings"):
            return
        self.app_settings.setValue("Settings/ParallelProcessingEnabled", self.settings_dialog.cb_parallel_processing.isChecked())
        self.app_settings.setValue("Settings/ParallelWorkerCount", self.settings_dialog.spin_parallel_workers.value())

    def persist_processing_mode(self):
        if not hasattr(self, "app_settings"):
            return
        self.app_settings.setValue("Settings/ProcessingMode", self.get_processing_mode())

    def on_theme_mode_changed(self, mode):
        """切换外观主题：更新分段按钮选中态、应用配色并持久化。"""
        self.settings_dialog.set_theme_mode(mode)
        if self.theme_manager is not None:
            self.theme_manager.set_mode(mode)
        if hasattr(self, "app_settings"):
            self.app_settings.setValue("Settings/ThemeMode", mode)
        QTimer.singleShot(0, self.apply_native_title_bar_theme)

    def apply_native_title_bar_theme(self, palette=None):
        """同步 Windows 原生标题栏颜色；非 Windows 平台会静默跳过。"""
        apply_windows_title_bar_theme(self, palette or active_palette())

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
        value = self.app_settings.value("Presets/FavoriteOptions", [])
        if isinstance(value, str):
            return [item for item in value.split(",") if item]
        return list(value or [])

    def save_favorite_preset(self):
        favorite_options = self.get_processing_options()
        self.app_settings.setValue("Presets/FavoriteOptions", favorite_options)
        self.active_preset_key = "favorite"
        self._set_preset_button_state("favorite")
        self.custom_selection_before_preset = set(favorite_options)
        self.refresh_selection_summary()
        self.show_success_message("✅ 已保存", "当前处理规则已保存为我的常用。")

    def restore_custom_selection(self):
        checkbox_groups = {}
        for opt_id, cb in self.all_checkboxes.items():
            if opt_id in ["处理完成后自动打开输出文件夹", "覆盖原始文件 (不推荐)"]:
                continue
            checkbox_groups.setdefault(id(cb), {"checkbox": cb, "keys": set()})["keys"].add(opt_id)

        self.is_applying_preset = True
        try:
            for group in checkbox_groups.values():
                should_check = any(key in self.custom_selection_before_preset for key in group["keys"])
                group["checkbox"].setChecked(should_check)
        finally:
            self.is_applying_preset = False

        self.active_preset_key = None
        self._set_preset_button_state(None)
        self.persist_all_settings()
        self.refresh_selection_summary()

    def toggle_preset(self, preset_key):
        if preset_key == "favorite" and not self.get_favorite_preset_options():
            self.show_warning_message("⚠️ 尚未保存", "请先点击“保存为常用”保存当前规则组合。")
            self._set_preset_button_state(self.active_preset_key)
            return

        if self.active_preset_key == preset_key:
            self.restore_custom_selection()
            return

        if self.active_preset_key is None:
            self.custom_selection_before_preset = set(self.get_selected_options())

        self.apply_preset(preset_key)

    def clear_selected_options(self):
        checkbox_groups = {}
        for opt_id, cb in self.all_checkboxes.items():
            if opt_id in ["处理完成后自动打开输出文件夹", "覆盖原始文件 (不推荐)"]:
                continue
            checkbox_groups.setdefault(id(cb), cb)

        self.is_applying_preset = True
        try:
            for cb in checkbox_groups.values():
                cb.setChecked(False)
        finally:
            self.is_applying_preset = False

        self.active_preset_key = None
        self._set_preset_button_state(None)
        self.custom_selection_before_preset = set()
        self.persist_all_settings()
        self.refresh_selection_summary()

    def apply_preset(self, preset_key, persist=True):
        if preset_key == "favorite":
            target_options = set(self.get_favorite_preset_options())
        else:
            preset = self.PRESET_OPTIONS.get(preset_key)
            if not preset:
                return
            target_options = set(preset["options"])

        if not target_options:
            return

        checkbox_groups = {}
        for opt_id, cb in self.all_checkboxes.items():
            if opt_id in ["处理完成后自动打开输出文件夹", "覆盖原始文件 (不推荐)"]:
                continue
            checkbox_groups.setdefault(id(cb), {"checkbox": cb, "keys": set()})["keys"].add(opt_id)

        self.is_applying_preset = True
        try:
            for group in checkbox_groups.values():
                should_check = any(key in target_options for key in group["keys"])
                group["checkbox"].setChecked(should_check)
        finally:
            self.is_applying_preset = False

        self.active_preset_key = preset_key
        self._set_preset_button_state(preset_key)
        self.persist_all_settings()
        self.refresh_selection_summary()

    def on_checkbox_toggled(self, _checked):
        if self.is_applying_preset:
            return

        sender_cb = self.sender()
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
        for i, page in enumerate(self.settings_pages):
            if i == index:
                page.show()
            else:
                page.hide()

    def get_selected_options(self):
        selected = []
        for opt_id, cb in self.all_checkboxes.items():
            if cb.isChecked():
                selected.append(opt_id)
        return selected

    def clear_tree_ui(self):
        self.tree.clear()

    def update_counters_ui(self, count):
        self.current_file_count = count
        self.list_title.setText(f"待处理队列 ({count})")
        if count == 0:
            self.queue_meta_label.setText("当前队列为空")
        else:
            self.queue_meta_label.setText(f"已加入{count}个PDF")
        if hasattr(self, "btn_precheck") and self.btn_precheck.property("precheckResultCurrent") is True:
            self.btn_precheck.setProperty("precheckResultCurrent", False)
            self.btn_precheck.show()
        self.refresh_selection_summary()

    def refresh_selection_summary(self):
        selected_count = len([
            opt for opt in self.get_selected_options()
            if opt not in ["处理完成后自动打开输出文件夹", "覆盖原始文件 (不推荐)"]
        ])
        preset_titles = {key: value["title"] for key, value in self.PRESET_OPTIONS.items()}
        preset_titles["favorite"] = "我的常用"
        preset_key = self.active_preset_key if isinstance(self.active_preset_key, str) else ""
        preset_text = preset_titles.get(preset_key, "自定义选择")
        total_files = self.current_file_count

        if selected_count == 0:
            self.selection_summary_label.setText("尚未选择任何处理规则")
        else:
            self.selection_summary_label.setText(f"已选择 {selected_count} 条规则")

        if self.active_preset_key:
            self.preset_summary_label.setText(f"当前已应用 {preset_text} 预设，并可继续手动微调规则。")
        else:
            self.preset_summary_label.setText("当前为自定义规则组合，可随时切换到 eCTD 预设。")

        self.info_label.setText(f"{total_files} 个文件 · {selected_count} 条规则 · {preset_text}")

        is_prechecking = hasattr(self, "btn_precheck") and self.btn_precheck.property("precheckMode") is True

        if is_prechecking:
            self.btn_start.setEnabled(False)
            self.btn_start.setToolTip("预检进行中，请稍候")
            self.btn_precheck.setEnabled(False)
            self.btn_precheck.setToolTip("预检进行中，请稍候")
            if hasattr(self, "btn_retry_failed"):
                self.btn_retry_failed.setEnabled(False)
                self.btn_retry_failed.hide()
                self.btn_retry_failed.setToolTip("预检进行中，请稍候")
            if hasattr(self, "btn_apply_precheck"):
                self.btn_apply_precheck.setEnabled(False)
                self.btn_apply_precheck.hide()
                self.btn_apply_precheck.setToolTip("预检进行中，请稍候")
            if hasattr(self, "btn_process_precheck_suggested"):
                self.btn_process_precheck_suggested.setEnabled(False)
                self.btn_process_precheck_suggested.hide()
                self.btn_process_precheck_suggested.setToolTip("预检进行中，请稍候")
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

        overwrite_cb = self.all_checkboxes.get("覆盖原始文件 (不推荐)")
        if overwrite_cb and overwrite_cb.isChecked():
            self.risk_hint_label.setText("当前启用了覆盖原始文件，执行前请确认已有备份。")
            self.risk_hint_label.setProperty("danger", True)
        else:
            self.risk_hint_label.setText("")
            self.risk_hint_label.setProperty("danger", False)

        self.style().unpolish(self.risk_hint_label)
        self.style().polish(self.risk_hint_label)

    def _create_section_label(self, text):
        lbl = QLabel(text)
        lbl.setObjectName("sectionCaption")
        return lbl

    def _create_checkbox(self, opt_id, title, desc, checked):
        container = QWidget()
        layout = QVBoxLayout(container)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(4)

        top_layout = QHBoxLayout()
        top_layout.setContentsMargins(0, 0, 0, 0)
        top_layout.setSpacing(8)

        cb = QCheckBox()
        cb.setChecked(checked)
        cb.setFocusPolicy(Qt.NoFocus)
        cb.toggled.connect(self.on_checkbox_toggled)
        self.all_checkboxes[opt_id] = cb

        title_lbl = QLabel(title)
        title_lbl.setWordWrap(True)
        title_lbl.setObjectName("optionTitle")
        title_lbl.mousePressEvent = lambda event, checkbox=cb: checkbox.toggle()

        top_layout.addWidget(cb, 0, Qt.AlignTop)
        top_layout.addWidget(title_lbl, 1)
        layout.addLayout(top_layout)

        if desc:
            desc_lbl = QLabel(desc)
            desc_lbl.setWordWrap(True)
            desc_lbl.setObjectName("optionDesc")
            layout.addWidget(desc_lbl)

        return container

    def apply_stylesheet(self):
        # 视觉样式全部集中在 theme.py 的应用级 QSS 中，由 ThemeManager 应用到
        # QApplication 并级联到本窗口。此处保留方法名以兼容旧调用点：主题未初始化
        # (如单元测试直接实例化 MainWindow) 时，回退到用当前生效配色本地渲染一次，
        # 确保窗口不至于完全无样式。
        app = QApplication.instance()
        if app is not None and app.styleSheet():
            # 已由 ThemeManager 设置了应用级样式，无需在窗口层重复。
            return
        self.setStyleSheet(build_app_qss(active_palette()))
