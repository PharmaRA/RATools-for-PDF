from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QFrame, QLabel, QPushButton, QTableWidget, QTableWidgetItem,
    QHeaderView, QCheckBox, QStackedWidget, QScrollArea, QButtonGroup,
    QDialog, QTextEdit
)
from PySide6.QtCore import Qt, Signal


class LogDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("处理日志记录")
        self.resize(650, 450)

        layout = QVBoxLayout(self)
        self.text_edit = QTextEdit()
        self.text_edit.setReadOnly(True)
        layout.addWidget(self.text_edit)

        btn_layout = QHBoxLayout()
        self.btn_export = QPushButton("⬇️ 导出为 CSV")
        self.btn_export.setStyleSheet(
            "background-color: #2563EB; color: white; border-radius: 6px; padding: 8px 16px; font-weight: bold; border: none;")
        self.btn_close = QPushButton("关闭")
        self.btn_close.setStyleSheet(
            "background-color: #E5E7EB; color: #374151; border-radius: 6px; padding: 8px 16px; font-weight: bold; border: none;")
        self.btn_close.clicked.connect(self.accept)

        btn_layout.addStretch()
        btn_layout.addWidget(self.btn_export)
        btn_layout.addWidget(self.btn_close)
        layout.addLayout(btn_layout)

        self.setStyleSheet("""
            QDialog { background-color: #F9FAFB; font-family: "Segoe UI", "Microsoft YaHei", sans-serif; }
            QTextEdit { background-color: white; border: 1px solid #E5E7EB; border-radius: 8px; padding: 12px; color: #374151; font-family: Consolas, "Courier New", monospace; font-size: 12px; }
        """)


class DropZoneLabel(QLabel):
    files_dropped = Signal(list)

    def __init__(self, text):
        super().__init__(text)
        self.setAcceptDrops(True)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            self.setStyleSheet(
                "border: 2px dashed #2563EB; background-color: #EFF6FF; border-radius: 12px; color: #2563EB;")
        else:
            event.ignore()

    def dragLeaveEvent(self, event):
        self.setStyleSheet("")

    def dropEvent(self, event):
        self.setStyleSheet("")
        urls = event.mimeData().urls()
        paths = [url.toLocalFile() for url in urls if url.isLocalFile()]
        if paths:
            self.files_dropped.emit(paths)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("RATools for PDF")
        self.resize(1100, 750)
        self.setMinimumSize(900, 600)

        self.all_checkboxes = {}

        self.MODULES_DATA = [
            {
                "icon": "👀",
                "title": "初始视图与文档属性",
                "options": [
                    "修改打开页面为第一页", "修改页面布局为默认",
                    "修改放大率为默认", "修改导览标签",
                    "PDF若存在书签则收起", "根据文件名在PDF文档属性中自动添加文件标题"
                ]
            },
            {
                "icon": "📄",
                "title": "页面与字体标准化",
                "options": [
                    "一键批量将页面切换成A4", "一键批量将页面切换成Letter",
                    "一键批量嵌入所有非标准字体（中文）", "一键批量嵌入所有非标准字体（英文）"
                ]
            },
            {
                "icon": "🔖",
                "title": "书签管理与优化",
                "options": [
                    "修改书签设置为承前缩放", "修改书签的设置为在新窗口中打开",
                    "删除书签的外部链接", "删除失效的书签（即未分配任何操作的书签）",
                    "删除未知动作的书签（即GoTo, GoToR和Launch之外的书签）"
                ]
            },
            {
                "icon": "🔗",
                "title": "超链接处理与外观控制",
                "options": [
                    "将外链接中的绝对路径转相对路径", "修改超链接的设置为承前缩放",
                    "修改超链接的设置为在新窗口中打开", "修改超链接文本至蓝色字体",
                    "修改超链接文本至黑色边框", "超链接有边框则蓝框黑字",
                    "超链接无边框且蓝字则蓝框黑字", "删除超链接边框"
                ]
            },
            {
                "icon": "🛡️",
                "title": "违规内容清理与安全性",
                "options": [
                    "删除外部链接（网页、邮箱地址）", "删除外部链接（网页、邮箱地址）且将文字改成黑色",
                    "删除失效的链接（即未分配任何操作的链接）", "删除无效的超链接，且将文字改成黑色",
                    "删除未知动作的链接（即GoTo, GoToRi和Launch之外的书签之外的链接）",
                    "删除JavaScript, 3D内容或者动态内容", "删除文档附件",
                    "删除文档标签", "删除PDF注释", "删除文档说明", "删除所有链接和书签"
                ]
            },
            {
                "icon": "📦",
                "title": "文件级优化与输出",
                "options": [
                    "PDF版本转换", "修改文件为快速网页浏览",
                    "文件名修改为符合电子申报/eCTD要求的格式"
                ]
            }
        ]

        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        header = QFrame()
        header.setObjectName("header")
        header.setFixedHeight(56)
        header_layout = QHBoxLayout(header)
        header_layout.setContentsMargins(24, 0, 24, 0)
        title_label = QLabel("📄 RATools for PDF")
        title_label.setObjectName("titleLabel")
        header_layout.addWidget(title_label)
        header_layout.addStretch()
        main_layout.addWidget(header)

        middle_container = QFrame()
        middle_layout = QHBoxLayout(middle_container)
        middle_layout.setContentsMargins(0, 0, 0, 0)
        middle_layout.setSpacing(0)

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

        settings_btn = QPushButton("⚙️  全局设置")
        settings_btn.setObjectName("navBtn")
        left_layout.addWidget(settings_btn)
        middle_layout.addWidget(left_sidebar)

        main_view = QFrame()
        main_view.setObjectName("mainView")
        main_view_layout = QVBoxLayout(main_view)
        main_view_layout.setContentsMargins(24, 24, 24, 24)
        main_view_layout.setSpacing(24)

        self.drop_zone = DropZoneLabel("☁️\n点击或将 PDF 文件拖拽到此处\n支持批量选择文件夹")
        self.drop_zone.setObjectName("dropZone")
        self.drop_zone.setAlignment(Qt.AlignCenter)
        self.drop_zone.setFixedHeight(128)
        main_view_layout.addWidget(self.drop_zone)

        list_container = QFrame()
        list_container.setObjectName("listContainer")
        list_layout = QVBoxLayout(list_container)
        list_layout.setContentsMargins(0, 0, 0, 0)
        list_layout.setSpacing(0)

        list_header = QFrame()
        list_header.setObjectName("listHeader")
        list_header_layout = QHBoxLayout(list_header)
        list_header_layout.setContentsMargins(16, 8, 16, 8)
        self.list_title = QLabel("待处理列表 (0)")
        self.list_title.setStyleSheet("font-weight: bold; color: #374151;")
        self.add_folder_btn = QPushButton("添加文件夹")
        self.add_folder_btn.setObjectName("textBtn")
        list_header_layout.addWidget(self.list_title)
        list_header_layout.addStretch()
        list_header_layout.addWidget(self.add_folder_btn)
        list_layout.addWidget(list_header)

        self.table = QTableWidget(0, 3)
        self.table.setHorizontalHeaderLabels(["文件名", "路径", "状态"])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Interactive)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setStretchLastSection(False)
        self.table.setColumnWidth(0, 260)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setShowGrid(False)
        self.table.verticalHeader().setVisible(False)
        list_layout.addWidget(self.table)
        main_view_layout.addWidget(list_container)

        middle_layout.addWidget(main_view)

        right_sidebar = QFrame()
        right_sidebar.setObjectName("rightSidebar")
        right_sidebar.setFixedWidth(320)
        right_layout = QVBoxLayout(right_sidebar)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(0)

        right_header = QFrame()
        right_header.setObjectName("rightHeader")
        rh_layout = QVBoxLayout(right_header)
        self.rh_title = QLabel(f"{self.MODULES_DATA[0]['title']} 设置")
        self.rh_title.setStyleSheet("font-weight: bold; font-size: 14px;")
        rh_desc = QLabel("勾选需要执行的处理规则")
        rh_desc.setStyleSheet("color: #6B7280; font-size: 12px;")
        rh_layout.addWidget(self.rh_title)
        rh_layout.addWidget(rh_desc)
        right_layout.addWidget(right_header)

        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll_area.setObjectName("settingsScroll")
        self.settings_stack = QStackedWidget()

        # === 定义专属的高级 IO 操作按钮 ===
        btn_style = "background-color: #F3F4F6; color: #374151; border-radius: 6px; padding: 6px 12px; font-weight: bold; border: 1px solid #D1D5DB;"

        self.btn_export_bookmarks = QPushButton("📤 批量导出书签 (CSV)")
        self.btn_import_bookmarks = QPushButton("📥 批量导入书签 (CSV)")
        self.btn_export_links = QPushButton("📤 批量导出链接 (JSON)")
        self.btn_import_links = QPushButton("📥 批量导入链接 (JSON)")

        for btn in [self.btn_export_bookmarks, self.btn_import_bookmarks, self.btn_export_links, self.btn_import_links]:
            btn.setStyleSheet(btn_style)
            btn.setCursor(Qt.PointingHandCursor)

        # 动态生成页面，并将按钮注入对应模块
        for mod in self.MODULES_DATA:
            page = QWidget()
            page_layout = QVBoxLayout(page)
            page_layout.setContentsMargins(20, 20, 20, 20)
            page_layout.setSpacing(16)

            page_layout.addWidget(self._create_section_label("处理规则选项"))
            for opt in mod["options"]:
                page_layout.addWidget(self._create_checkbox(opt, "", False))

            # 书签模块注入 IO 按钮
            if mod["title"] == "书签管理与优化":
                page_layout.addSpacing(12)
                page_layout.addWidget(self._create_section_label("高级数据交换"))
                btn_layout = QVBoxLayout()
                btn_layout.setSpacing(8)  # 设置纵向排列的间距
                btn_layout.addWidget(self.btn_export_bookmarks)
                btn_layout.addWidget(self.btn_import_bookmarks)
                page_layout.addLayout(btn_layout)

            # 链接模块注入 IO 按钮
            elif mod["title"] == "超链接处理与外观控制":
                page_layout.addSpacing(12)
                page_layout.addWidget(self._create_section_label("高级数据交换"))
                btn_layout = QVBoxLayout()
                btn_layout.setSpacing(8)  # 设置纵向排列的间距
                btn_layout.addWidget(self.btn_export_links)
                btn_layout.addWidget(self.btn_import_links)
                page_layout.addLayout(btn_layout)

            page_layout.addStretch()
            self.settings_stack.addWidget(page)

        scroll_area.setWidget(self.settings_stack)
        right_layout.addWidget(scroll_area)

        middle_layout.addWidget(right_sidebar)
        main_layout.addWidget(middle_container)

        footer = QFrame()
        footer.setObjectName("footer")
        footer.setFixedHeight(64)
        footer_layout = QHBoxLayout(footer)
        footer_layout.setContentsMargins(24, 0, 24, 0)

        self.btn_clear = QPushButton("🗑️ 清空列表")
        self.btn_clear.setObjectName("actionBtn")
        self.btn_log = QPushButton("📋 查看/导出日志")
        self.btn_log.setObjectName("actionBtn")
        footer_layout.addWidget(self.btn_clear)
        footer_layout.addWidget(self.btn_log)
        footer_layout.addStretch()

        self.info_label = QLabel("共计 <b>0</b> 个文件")
        self.btn_start = QPushButton("▶ 开始批量处理")
        self.btn_start.setObjectName("startBtn")
        footer_layout.addWidget(self.info_label)
        footer_layout.addSpacing(16)
        footer_layout.addWidget(self.btn_start)
        main_layout.addWidget(footer)

        self.nav_btn_group.idClicked.connect(self.switch_settings_page)
        self.apply_stylesheet()

    def switch_settings_page(self, index):
        self.settings_stack.setCurrentIndex(index)
        self.rh_title.setText(f"{self.MODULES_DATA[index]['title']} 设置")

    def get_selected_options(self):
        selected = []
        for title, cb in self.all_checkboxes.items():
            if cb.isChecked():
                selected.append(title)
        return selected

    def add_table_row(self, name, path, status):
        row_index = self.table.rowCount()
        self.table.insertRow(row_index)
        name_item = QTableWidgetItem(name)
        name_item.setToolTip(name)
        path_item = QTableWidgetItem(path)
        path_item.setToolTip(path)
        status_item = QTableWidgetItem(status)
        status_item.setForeground(Qt.darkGray)
        self.table.setItem(row_index, 0, name_item)
        self.table.setItem(row_index, 1, path_item)
        self.table.setItem(row_index, 2, status_item)

    def update_table_row_status(self, row_index, status_text, color=Qt.black):
        item = self.table.item(row_index, 2)
        if item:
            item.setText(status_text)
            item.setForeground(color)

    def clear_table_ui(self):
        self.table.setRowCount(0)

    def update_counters_ui(self, count):
        self.list_title.setText(f"待处理列表 ({count})")
        self.info_label.setText(f"共计 <b>{count}</b> 个文件")

    def _create_section_label(self, text):
        lbl = QLabel(text)
        lbl.setStyleSheet("color: #9CA3AF; font-size: 12px; font-weight: bold; margin-bottom: 4px;")
        return lbl

    def _create_checkbox(self, title, desc, checked):
        container = QWidget()
        layout = QVBoxLayout(container)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(4)
        top_layout = QHBoxLayout()
        top_layout.setContentsMargins(0, 0, 0, 0)
        top_layout.setSpacing(8)
        cb = QCheckBox()
        cb.setChecked(checked)
        self.all_checkboxes[title] = cb
        title_lbl = QLabel(title)
        title_lbl.setWordWrap(True)
        title_lbl.setStyleSheet("font-weight: 500; color: #374151;")
        title_lbl.mousePressEvent = lambda event, checkbox=cb: checkbox.toggle()
        top_layout.addWidget(cb, 0, Qt.AlignTop)
        top_layout.addWidget(title_lbl, 1)
        layout.addLayout(top_layout)
        if desc:
            desc_lbl = QLabel(desc)
            desc_lbl.setWordWrap(True)
            desc_lbl.setStyleSheet("color: #6B7280; font-size: 11px; margin-left: 24px;")
            layout.addWidget(desc_lbl)
        return container

    def apply_stylesheet(self):
        qss = """
        QMainWindow { background-color: #F9FAFB; font-family: "Segoe UI", "Microsoft YaHei", sans-serif; font-size: 13px; }
        #header, #leftSidebar, #rightSidebar, #footer { background-color: white; }
        #header { border-bottom: 1px solid #E5E7EB; }
        #leftSidebar { border-right: 1px solid #E5E7EB; }
        #rightSidebar { border-left: 1px solid #E5E7EB; }
        #footer { border-top: 1px solid #E5E7EB; }
        #titleLabel { font-size: 18px; font-weight: bold; color: #111827; }
        #navTitle { color: #9CA3AF; font-size: 11px; font-weight: bold; letter-spacing: 1px; margin-bottom: 4px; }
        #navBtn { text-align: left; padding: 10px 12px; border: none; border-radius: 8px; color: #4B5563; font-weight: 500; background-color: transparent; }
        #navBtn:hover { background-color: #F3F4F6; }
        #navBtn:checked { background-color: #EFF6FF; color: #1D4ED8; font-weight: bold; }
        #mainView { background-color: #F9FAFB; }
        #dropZone { border: 2px dashed #D1D5DB; border-radius: 12px; background-color: white; color: #6B7280; font-weight: 500; }
        #dropZone:hover { border-color: #60A5FA; background-color: #EFF6FF; color: #2563EB; }
        #listContainer { background-color: white; border: 1px solid #E5E7EB; border-radius: 12px; }
        #listHeader { border-bottom: 1px solid #E5E7EB; background-color: #F9FAFB; border-top-left-radius: 12px; border-top-right-radius: 12px; }
        #textBtn { color: #2563EB; border: none; background: transparent; font-weight: 500; }
        #textBtn:hover { color: #1D4ED8; }
        QTableWidget { border: none; background-color: white; color: #374151; gridline-color: #F3F4F6; }
        QTableWidget::item { padding: 4px; border-bottom: 1px solid #F3F4F6; }
        QHeaderView::section { background-color: white; border: none; border-bottom: 1px solid #E5E7EB; padding: 8px; color: #6B7280; font-weight: 500; text-align: left; }
        #rightHeader { border-bottom: 1px solid #E5E7EB; padding: 16px 20px; }
        #settingsScroll { border: none; background-color: transparent; }
        #settingsScroll > QWidget > QWidget { background-color: white; }

        QScrollBar:vertical { border: none; background: transparent; width: 8px; margin: 0px; }
        QScrollBar::handle:vertical { background: #D1D5DB; min-height: 30px; border-radius: 4px; }
        QScrollBar::handle:vertical:hover { background: #9CA3AF; }
        QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0px; background: none; }
        QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical { background: transparent; }

        QScrollBar:horizontal { border: none; background: transparent; height: 8px; margin: 0px; }
        QScrollBar::handle:horizontal { background: #D1D5DB; min-width: 30px; border-radius: 4px; }
        QScrollBar::handle:horizontal:hover { background: #9CA3AF; }
        QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal { width: 0px; background: none; }
        QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal { background: transparent; }

        QCheckBox::indicator { width: 16px; height: 16px; border-radius: 4px; border: 1px solid #D1D5DB; background: white; margin-top: 1px;}
        QCheckBox::indicator:checked { background: #2563EB; border-color: #2563EB; }
        #actionBtn { padding: 8px 16px; border: 1px solid transparent; border-radius: 8px; background-color: transparent; color: #4B5563; font-weight: 500; }
        #actionBtn:hover { background-color: #F3F4F6; }
        #startBtn { padding: 10px 24px; background-color: #2563EB; color: white; border-radius: 8px; font-weight: bold; border: none; }
        #startBtn:hover { background-color: #1D4ED8; }
        #startBtn:pressed { background-color: #1E40AF; }
        """
        self.setStyleSheet(qss)