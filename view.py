import os
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QFrame, QLabel, QPushButton, QTreeWidget, QTreeWidgetItem,
    QHeaderView, QCheckBox, QScrollArea, QButtonGroup,
    QDialog, QTextEdit, QSizePolicy
)
from PySide6.QtCore import Qt, Signal, QPoint, QSettings
from PySide6.QtGui import QIcon  # 引入 QIcon

from app_paths import get_app_dir, get_resource_path


# ================== 自定义无边框拖拽对话框基类 ==================
class FramelessDraggableDialog(QDialog):
    def __init__(self, title_text, parent=None):
        super().__init__(parent)
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.Dialog)
        self.setAttribute(Qt.WA_TranslucentBackground)  # 支持圆角透明背景

        self.main_layout = QVBoxLayout(self)
        self.main_layout.setContentsMargins(0, 0, 0, 0)
        self.main_layout.setSpacing(0)

        # 整体圆角和边框容器
        self.bg_frame = QFrame()
        self.bg_frame.setStyleSheet("""
            QFrame {
                background-color: white;
                border: 1px solid #D1D5DB;
                border-radius: 8px;
            }
        """)
        bg_layout = QVBoxLayout(self.bg_frame)
        bg_layout.setContentsMargins(0, 0, 0, 0)
        bg_layout.setSpacing(0)

        # 顶部自定义标题栏
        self.title_bar = QFrame()
        self.title_bar.setFixedHeight(40)
        self.title_bar.setStyleSheet("""
            background-color: #F9FAFB; 
            border: none;
            border-bottom: 1px solid #E5E7EB; 
            border-top-left-radius: 8px; 
            border-top-right-radius: 8px;
        """)
        tb_layout = QHBoxLayout(self.title_bar)
        tb_layout.setContentsMargins(16, 0, 8, 0)

        title_lbl = QLabel(title_text)
        title_lbl.setStyleSheet("font-weight: bold; color: #374151; font-size: 13px; border: none;")
        tb_layout.addWidget(title_lbl)
        tb_layout.addStretch()

        btn_close = QPushButton("✕")
        btn_close.setFixedSize(30, 30)
        btn_close.setStyleSheet("""
            QPushButton { background: transparent; border: none; font-size: 14px; color: #9CA3AF; border-radius: 4px; } 
            QPushButton:hover { background-color: #E5E7EB; color: #EF4444; }
        """)
        btn_close.clicked.connect(self.reject)
        tb_layout.addWidget(btn_close)

        bg_layout.addWidget(self.title_bar)

        # 内部内容区
        self.content_widget = QWidget()
        self.content_widget.setStyleSheet("border: none; background-color: transparent;")
        self.content_layout = QVBoxLayout(self.content_widget)
        self.content_layout.setContentsMargins(24, 24, 24, 24)
        bg_layout.addWidget(self.content_widget)

        self.main_layout.addWidget(self.bg_frame)

    def mousePressEvent(self, event):
        """接管鼠标按下事件：若点击在高度 40px 以内的标题栏区域，则记录起始坐标"""
        if event.button() == Qt.LeftButton and event.position().y() < 40:
            self.drag_pos = event.globalPosition().toPoint() - self.frameGeometry().topLeft()
            event.accept()

    def mouseMoveEvent(self, event):
        """接管鼠标移动事件：应用拖拽偏移"""
        if event.buttons() == Qt.LeftButton and hasattr(self, 'drag_pos'):
            self.move(event.globalPosition().toPoint() - self.drag_pos)
            event.accept()

    def mouseReleaseEvent(self, event):
        """接管释放事件：清除记录"""
        if hasattr(self, 'drag_pos'):
            del self.drag_pos
            event.accept()


# ================== 具体的业务对话框 ==================

class CustomMessageBox(FramelessDraggableDialog):
    """用于完全替代原生 QMessageBox 的统一提示框"""

    def __init__(self, title_text, message_text, msg_type="info", show_cancel=False, parent=None):
        super().__init__(title_text, parent)
        self.resize(400, 200)

        icons = {
            "info": "ℹ️",
            "success": "✅",
            "warning": "⚠️",
            "error": "❌",
            "question": "❓"
        }
        icon_char = icons.get(msg_type, "ℹ️")

        content_h_layout = QHBoxLayout()
        content_h_layout.setSpacing(16)

        icon_lbl = QLabel(icon_char)
        icon_lbl.setStyleSheet("font-size: 36px; border: none; background: transparent;")
        icon_lbl.setAlignment(Qt.AlignTop)

        msg_lbl = QLabel(message_text)
        msg_lbl.setWordWrap(True)
        msg_lbl.setStyleSheet("color: #374151; font-size: 13px; border: none; line-height: 1.5;")
        msg_lbl.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)

        content_h_layout.addWidget(icon_lbl)
        content_h_layout.addWidget(msg_lbl, 1)

        self.content_layout.addLayout(content_h_layout)
        self.content_layout.addStretch()

        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(12)
        btn_layout.addStretch()

        if show_cancel:
            self.btn_cancel = QPushButton("取 消")
            self.btn_cancel.setFixedSize(80, 32)
            self.btn_cancel.setStyleSheet(
                "background-color: #F3F4F6; color: #374151; border-radius: 6px; font-weight: bold; border: 1px solid #D1D5DB;")
            self.btn_cancel.clicked.connect(self.reject)
            btn_layout.addWidget(self.btn_cancel)

        self.btn_ok = QPushButton("确 定")
        self.btn_ok.setFixedSize(80, 32)
        if msg_type in ["error", "warning"]:
            self.btn_ok.setStyleSheet(
                "background-color: #EF4444; color: white; border-radius: 6px; font-weight: bold; border: none;")
        else:
            self.btn_ok.setStyleSheet(
                "background-color: #2563EB; color: white; border-radius: 6px; font-weight: bold; border: none;")
        self.btn_ok.clicked.connect(self.accept)
        btn_layout.addWidget(self.btn_ok)

        self.content_layout.addLayout(btn_layout)


class LogDialog(FramelessDraggableDialog):
    def __init__(self, parent=None):
        super().__init__("📝 处理日志记录", parent)
        self.resize(650, 480)

        self.text_edit = QTextEdit()
        self.text_edit.setReadOnly(True)
        self.text_edit.setStyleSheet(
            "background-color: white; border: 1px solid #E5E7EB; border-radius: 8px; padding: 12px; color: #374151; font-family: Consolas, 'Courier New', monospace; font-size: 12px;")
        self.content_layout.addWidget(self.text_edit)

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
        self.content_layout.addLayout(btn_layout)


class SettingsDialog(FramelessDraggableDialog):
    def __init__(self, parent=None):
        super().__init__("⚙️ 全局设置", parent)
        self.resize(400, 240)

        self.content_layout.setSpacing(16)

        title = QLabel("常规选项")
        title.setStyleSheet("color: #9CA3AF; font-size: 12px; font-weight: bold; border: none;")
        self.content_layout.addWidget(title)

        cb_style = """
            QCheckBox { font-size: 13px; color: #374151; spacing: 8px; border: none; }
            QCheckBox::indicator { width: 16px; height: 16px; border-radius: 4px; border: 1px solid #D1D5DB; background: white; margin-top: 1px;}
            QCheckBox::indicator:checked { background: #2563EB; border-color: #2563EB; }
        """

        self.cb_auto_open = QCheckBox("处理完成后自动打开输出文件夹")
        self.cb_auto_open.setChecked(True)
        self.cb_auto_open.setStyleSheet(cb_style)

        self.cb_overwrite = QCheckBox("覆盖原始文件 (不推荐)")
        self.cb_overwrite.setChecked(False)
        self.cb_overwrite.setStyleSheet(cb_style.replace("color: #374151;", "color: #EF4444;"))

        self.content_layout.addWidget(self.cb_auto_open)
        self.content_layout.addWidget(self.cb_overwrite)

        danger_hint = QLabel("危险操作：直接覆盖源 PDF。建议仅在已有备份且确认规则无误后使用。")
        danger_hint.setWordWrap(True)
        danger_hint.setObjectName("dangerHint")
        self.content_layout.addWidget(danger_hint)

        self.content_layout.addStretch()

        btn_close = QPushButton("确 定")
        btn_close.setFixedHeight(36)
        btn_close.setStyleSheet(
            "background-color: #2563EB; color: white; border-radius: 6px; font-weight: bold; border: none;")
        btn_close.clicked.connect(self.accept)
        self.content_layout.addWidget(btn_close)


class AboutDialog(FramelessDraggableDialog):
    def __init__(self, parent=None):
        super().__init__("ℹ️ 关于软件", parent)
        self.resize(520, 360)
        self.content_layout.setSpacing(16)

        hero_card = QFrame()
        hero_card.setStyleSheet(
            "QFrame {"
            "background-color: #F8FBFF;"
            "border: 1px solid #D7E7F8;"
            "border-radius: 10px;"
            "}"
        )
        hero_layout = QVBoxLayout(hero_card)
        hero_layout.setContentsMargins(18, 16, 18, 16)
        hero_layout.setSpacing(6)

        brand_title = QLabel("RATools for PDF")
        brand_title.setStyleSheet("font-size: 22px; font-weight: bold; color: #1D4ED8; border: none;")
        version_badge = QLabel("Version 1.0.0")
        version_badge.setStyleSheet(
            "background-color: white; color: #2563EB; border: 1px solid #BFDBFE;"
            "border-radius: 999px; padding: 4px 10px; font-weight: 600;"
        )
        version_badge.setAlignment(Qt.AlignCenter)
        version_badge.setMaximumWidth(110)

        hero_layout.addWidget(brand_title)
        hero_layout.addWidget(version_badge, 0, Qt.AlignLeft)
        self.content_layout.addWidget(hero_card)

        intro_text = QLabel(
            "用于RA递交资料整理的PDF处理工具，"
            "帮助用户以更稳定的方式完成eCTD场景下常见的批量标准化操作。"
        )
        intro_text.setWordWrap(True)
        intro_text.setStyleSheet("color: #374151; font-size: 13px; line-height: 1.7; border: none;")
        self.content_layout.addWidget(intro_text)

        features_title = QLabel("核心功能")
        features_title.setStyleSheet("color: #111827; font-size: 13px; font-weight: bold; border: none;")
        self.content_layout.addWidget(features_title)

        features_text = QLabel(
            "• 批量导入PDF文件或文件夹\n"
            "• 按模块勾选规则，支持中国eCTD/美国eCTD预设\n"
            "• 覆盖文档属性、书签、链接、动态内容与附件等常见合规项\n"
            "• 输出处理日志，便于复核与追踪"
        )
        features_text.setWordWrap(True)
        features_text.setStyleSheet("color: #475569; font-size: 12px; line-height: 1.8; border: none;")
        self.content_layout.addWidget(features_text)

        tech_card = QFrame()
        tech_card.setStyleSheet(
            "QFrame {"
            "background-color: #F8FAFC;"
            "border: 1px solid #E2E8F0;"
            "border-radius: 10px;"
            "}"
        )
        tech_layout = QVBoxLayout(tech_card)
        tech_layout.setContentsMargins(16, 14, 16, 14)
        tech_layout.setSpacing(4)

        tech_title = QLabel("技术与许可")
        tech_title.setStyleSheet("color: #111827; font-size: 13px; font-weight: bold; border: none;")
        tech_detail = QLabel(
            "基于PySide6、PyMuPDF及Ghostscript等项目构建\n"
            "遵循GNU GPL v3开源协议"
        )
        tech_detail.setWordWrap(True)
        tech_detail.setStyleSheet("color: #64748B; font-size: 12px; line-height: 1.7; border: none;")
        tech_layout.addWidget(tech_title)
        tech_layout.addWidget(tech_detail)
        self.content_layout.addWidget(tech_card)

        self.content_layout.addStretch()

        btn_close = QPushButton("关 闭")
        btn_close.setFixedHeight(36)
        btn_close.setStyleSheet(
            "background-color: #E5E7EB; color: #374151; border-radius: 6px; font-weight: bold; border: none;")
        btn_close.clicked.connect(self.accept)
        self.content_layout.addWidget(btn_close)


# ================== 自定义组件与主窗口 ==================
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
        self.setWindowIcon(QIcon(get_resource_path("icon.png")))

        self.resize(1100, 750)
        self.setMinimumSize(900, 600)

        self.all_checkboxes = {}
        self.current_file_count = 0

        self.settings_dialog = SettingsDialog(self)
        self.all_checkboxes["处理完成后自动打开输出文件夹"] = self.settings_dialog.cb_auto_open
        self.all_checkboxes["覆盖原始文件 (不推荐)"] = self.settings_dialog.cb_overwrite
        self.settings_dialog.cb_overwrite.toggled.connect(lambda _checked: self.refresh_selection_summary())

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
                    {"id": "page_size_a4", "title": "强制转为 A4 尺寸", "desc": "统一将所有页面裁切/调整为标准的 A4 纸张尺寸"},
                    {"id": "page_size_letter", "title": "强制转为 Letter 尺寸", "desc": "统一将所有页面裁切/调整为标准的 Letter (信纸) 尺寸"},
                    {"id": "embed_nonstandard_fonts", "title": "嵌入全部非标准字体", "desc": "利用 Ghostscript 引擎将文档中使用的所有非标准字体完全嵌入"}
                ]
            },
            {
                "icon": "🔖",
                "title": "书签管理",
                "options": [
                    {"id": "bookmark_inherit_zoom", "title": "书签设为承前缩放", "desc": "点击书签跳转时，保持当前页面的缩放比例不变 (Inherit Zoom)"},
                    {"id": "bookmark_open_new_window", "title": "书签动作：新窗口打开", "desc": "配置书签的链接跳转默认在新的 PDF 浏览器窗口中打开"},
                    {"id": "bookmark_remove_external_links", "title": "清理书签外部链接", "desc": "移除书签中指向网页或外部文件的 URI 动作"},
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
                    {"id": "link_black_border", "title": "链接区域加黑框", "desc": "为所有的有效超链接区域添加 1px 的黑色实线边框"},
                    {"id": "link_bordered_to_blue_border", "title": "标准化有框链接", "desc": "若超链接已存在边框，则统一转为蓝框黑字样式"},
                    {"id": "link_unbordered_blue_to_blue_border", "title": "标准化无框蓝字链接", "desc": "若超链接无边框且文字为蓝色，则统一转为蓝框黑字样式"},
                    {"id": "link_remove_border", "title": "清除所有链接边框", "desc": "移除文档内所有超链接的可见边框，保持页面排版干净"}
                ]
            },
            {
                "icon": "🛡️",
                "title": "内容合规与安全性",
                "options": [
                    {"id": "cleanup_remove_external_uri", "title": "删除外部 URI 链接", "desc": "清理指向外部网站、邮箱等所有 URI 类型的超链接"},
                    {"id": "cleanup_remove_external_uri_and_text_black", "title": "删除外部 URI 链接并去色", "desc": "清理 URI 链接的同时，将该链接对应的文本颜色重置为黑色"},
                    {"id": "cleanup_remove_invalid_links", "title": "清理失效超链接", "desc": "自动扫描并移除所有未分配有效动作 (Action) 的空链接"},
                    {"id": "cleanup_remove_invalid_links_and_text_black", "title": "清理失效链接并去色", "desc": "移除空链接，并将该区域相关的文本颜色恢复为普通黑色"},
                    {"id": "cleanup_remove_unknown_action_links", "title": "清理非标准动作链接", "desc": "仅保留内部/外部跳转和执行动作，移除其它所有的特殊行为"},
                    {"id": "cleanup_remove_dynamic_content", "title": "彻底清除动态内容 (JS/3D)", "desc": "删除文档内所有的 JavaScript 脚本、3D 模型等交互元素以满足安全合规"},
                    {"id": "cleanup_remove_attachments", "title": "移除所有内嵌附件", "desc": "清理 PDF 内部打包的所有附加文件 (.zip, .xml 等)"},
                    {"id": "cleanup_remove_tags", "title": "移除结构化标签", "desc": "删除 PDF 结构树 (StructTreeRoot) 和标记信息 (MarkInfo)"},
                    {"id": "cleanup_remove_annotations", "title": "清理所有高亮/批注", "desc": "删除文本框、高亮、画笔等所有非链接类型的交互式注释"},
                    {"id": "cleanup_remove_metadata", "title": "清空文档元数据", "desc": "移除所有标题、作者、创建时间等 PieceInfo 和 Metadata"},
                    {"id": "cleanup_remove_all_links_bookmarks", "title": "移除全部链接和书签", "desc": "一键清除文档内所有的导航书签与页面超链接"}
                ]
            },
            {
                "icon": "📦",
                "title": "文件级优化与输出",
                "options": [
                    {"id": "convert_pdf_version", "title": "PDF 版本转换", "desc": "将PDF版本修改为1.7版本"},
                    {"id": "fast_web_view", "title": "启用线性化 (快速网页浏览)", "desc": "优化文档结构以支持 Web 环境下的流式加载和边下边看"},
                    {"id": "filename_ectd_format", "title": "eCTD 文件名合规格式化", "desc": "自动将输出文件名转为小写、去除空格并替换非法字符"}
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
        self.list_title.setStyleSheet("font-weight: 700; color: #1F2937;")
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
        self.rh_title.setStyleSheet("font-weight: bold; font-size: 14px;")

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

        btn_style = "background-color: #F3F4F6; color: #374151; border-radius: 6px; padding: 6px 12px; font-weight: bold; border: 1px solid #D1D5DB;"
        self.btn_export_bookmarks = QPushButton("📤 批量导出书签 (CSV)")
        self.btn_import_bookmarks = QPushButton("📥 批量导入书签 (CSV)")
        self.btn_export_links = QPushButton("📤 批量导出链接 (JSON)")
        self.btn_import_links = QPushButton("📥 批量导入链接 (JSON)")

        for btn in [self.btn_export_bookmarks, self.btn_import_bookmarks, self.btn_export_links, self.btn_import_links]:
            btn.setStyleSheet(btn_style)
            btn.setCursor(Qt.PointingHandCursor)

        for mod in self.MODULES_DATA:
            page = QWidget()
            page_layout = QVBoxLayout(page)
            page_layout.setContentsMargins(20, 18, 20, 20)
            page_layout.setSpacing(14)

            for opt in mod["options"]:
                page_layout.addWidget(self._create_checkbox(opt["id"], opt["title"], opt["desc"], False))

            if mod["title"] == "书签管理与优化":
                page_layout.addSpacing(12)
                page_layout.addWidget(self._create_section_label("导出/导入书签"))
                btn_layout = QVBoxLayout()
                btn_layout.setSpacing(8)
                btn_layout.addWidget(self.btn_export_bookmarks)
                btn_layout.addWidget(self.btn_import_bookmarks)
                page_layout.addLayout(btn_layout)

            elif mod["title"] == "超链接处理与外观控制":
                page_layout.addSpacing(12)
                page_layout.addWidget(self._create_section_label("导出/导入链接"))
                btn_layout = QVBoxLayout()
                btn_layout.setSpacing(8)
                btn_layout.addWidget(self.btn_export_links)
                btn_layout.addWidget(self.btn_import_links)
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

        self.btn_clear_selected_options = QPushButton("全部取消")
        self.btn_clear_selected_options.setObjectName("actionBtn")

        self.preset_summary_label = QLabel("默认载入中国 eCTD 预设，可按需微调。")
        self.preset_summary_label.setObjectName("presetSummary")

        self.preset_btn_group = QButtonGroup(self)
        self.preset_btn_group.setExclusive(True)
        self.preset_btn_group.addButton(self.btn_preset_china)
        self.preset_btn_group.addButton(self.btn_preset_us)

        preset_layout.addSpacing(8)
        preset_layout.addWidget(self.preset_summary_label)
        preset_layout.addStretch()
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
        self.btn_start = QPushButton("▶ 开始批量处理")
        self.btn_start.setObjectName("startBtn")
        footer_layout.addWidget(self.info_label)
        footer_layout.addSpacing(16)
        footer_layout.addWidget(self.processing_hint_label)
        footer_layout.addSpacing(16)
        footer_layout.addWidget(self.risk_hint_label)
        footer_layout.addSpacing(16)
        footer_layout.addWidget(self.btn_start)
        main_layout.addWidget(footer)

        self.nav_btn_group.idClicked.connect(self.switch_settings_page)

        # 立刻触发一次以展示第一页内容
        self.switch_settings_page(0)
        self.apply_stylesheet()

        # ================= 初始化 QSettings 持久化存储 =================
        current_dir = get_app_dir()
        ini_path = os.path.join(current_dir, "settings.ini")
        self.app_settings = QSettings(ini_path, QSettings.IniFormat)

        self.settings_key_map = {
            "处理完成后自动打开输出文件夹": "Settings/AutoOpenOutput",
            "覆盖原始文件 (不推荐)": "Settings/OverwriteOriginal"
        }
        for i, mod in enumerate(self.MODULES_DATA):
            for j, opt in enumerate(mod["options"]):
                self.settings_key_map[opt["id"]] = f"Modules/Mod_{i}_Opt_{j}"

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

    def show_confirm_message(self, title, message):
        dlg = CustomMessageBox(title, message, msg_type="question", show_cancel=True, parent=self)
        return dlg.exec() == QDialog.Accepted

    def load_all_settings(self):
        for opt_id, cb in self.all_checkboxes.items():
            key = self.settings_key_map.get(opt_id)
            if key:
                val = self.app_settings.value(key)
                if val is not None:
                    is_checked = str(val).lower() == 'true'
                    cb.setChecked(is_checked)

        # 兼容旧版本：如果用户之前勾选过“中文/英文字体嵌入”，迁移到新的统一选项
        merged_font_opt = self.all_checkboxes.get("embed_nonstandard_fonts")
        if merged_font_opt and not merged_font_opt.isChecked():
            old_cn = str(self.app_settings.value("Modules/Mod_1_Opt_2", "false")).lower() == 'true'
            old_en = str(self.app_settings.value("Modules/Mod_1_Opt_3", "false")).lower() == 'true'
            if old_cn or old_en:
                merged_font_opt.setChecked(True)

        # 默认恢复上次会话的勾选状态（自定义），不自动套用预设
        self.custom_selection_before_preset = set(self.get_selected_options())
        self.active_preset_key = None
        self._set_preset_button_state(None)

    def closeEvent(self, event):
        for opt_id, cb in self.all_checkboxes.items():
            key = self.settings_key_map.get(opt_id)
            if key:
                self.app_settings.setValue(key, cb.isChecked())
        super().closeEvent(event)

    def get_selected_preset(self):
        return self.active_preset_key

    def _set_preset_button_state(self, preset_key):
        self.preset_btn_group.setExclusive(False)
        self.btn_preset_china.setChecked(preset_key == "china")
        self.btn_preset_us.setChecked(preset_key == "us")
        self.preset_btn_group.setExclusive(True)

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
        self.refresh_selection_summary()

    def toggle_preset(self, preset_key):
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
        self.refresh_selection_summary()

    def apply_preset(self, preset_key, persist=True):
        preset = self.PRESET_OPTIONS.get(preset_key)
        if not preset:
            return

        target_options = set(preset["options"])

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
        self.refresh_selection_summary()

    def on_checkbox_toggled(self, _checked):
        if self.is_applying_preset:
            return

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
        self.refresh_selection_summary()

    def refresh_selection_summary(self):
        selected_count = len([
            opt for opt in self.get_selected_options()
            if opt not in ["处理完成后自动打开输出文件夹", "覆盖原始文件 (不推荐)"]
        ])
        preset_titles = {key: value["title"] for key, value in self.PRESET_OPTIONS.items()}
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
        lbl.setStyleSheet("color: #9CA3AF; font-size: 12px; font-weight: bold; margin-bottom: 4px;")
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
        QMainWindow { background-color: #F4F6F8; font-family: "Segoe UI", "Microsoft YaHei", sans-serif; font-size: 13px; color: #1F2937; }

        #header, #leftSidebar, #rightSidebar, #footer, #presetBar { background-color: white; }
        #header { border-bottom: 1px solid #DCE3EA; }
        #leftSidebar { border-right: 1px solid #E5E7EB; }
        #rightSidebar { border-left: 1px solid #E5E7EB; }
        #footer { border-top: 1px solid #DCE3EA; }
        #presetBar { border-top: 1px solid #E5E7EB; border-bottom: 1px solid #E5E7EB; }
        #presetLabel { color: #52606D; font-size: 12px; font-weight: 700; }
        #presetSummary { color: #6B7280; font-size: 12px; }

        #topBtn { background: transparent; border: none; font-weight: 600; color: #4B5563; padding: 6px 12px; border-radius: 6px; }
        #topBtn:hover { background-color: #F3F4F6; color: #111827; }

        #navTitle { color: #94A3B8; font-size: 11px; font-weight: bold; letter-spacing: 1px; margin-bottom: 6px; }
        #navBtn { text-align: left; padding: 11px 12px; border: 1px solid transparent; border-radius: 10px; color: #475569; font-weight: 600; background-color: transparent; }
        #navBtn:hover { background-color: #F8FAFC; border-color: #E2E8F0; }
        #navBtn:checked { background-color: #E8F1FF; color: #155EEF; font-weight: 700; border-color: #BFDBFE; }
        #mainView { background-color: #F4F6F8; }
        #importCard, #listContainer { background-color: white; border: 1px solid #DCE3EA; border-radius: 16px; }
        #sectionTitle { font-size: 14px; font-weight: 700; color: #0F172A; }
        #sectionHint { color: #64748B; font-size: 12px; }
        #dropZone { border: 2px dashed #B8C6DB; border-radius: 14px; background-color: #F8FBFF; color: #52606D; font-weight: 600; }
        #dropZone:hover { border-color: #60A5FA; background-color: #EFF6FF; color: #1D4ED8; }
        #secondaryBtn { padding: 8px 14px; border: 1px solid #DCE3EA; border-radius: 9px; background-color: white; color: #334155; font-weight: 600; }
        #secondaryBtn:hover { background-color: #F8FAFC; border-color: #CBD5E1; }
        #mutedLabel { color: #64748B; font-size: 12px; }
        #listHeader { border-bottom: 1px solid #E5E7EB; background-color: #F8FAFC; border-top-left-radius: 16px; border-top-right-radius: 16px; }

        QTreeWidget { border: none; background-color: white; color: #334155; outline: none; border-bottom-left-radius: 16px; border-bottom-right-radius: 16px; alternate-background-color: #FBFDFF; }
        QTreeWidget::item { padding: 7px; border-bottom: 1px solid #F1F5F9; }
        QTreeWidget::item:selected { background-color: #E8F1FF; color: #155EEF; }
        QTreeWidget::branch:selected { background: transparent; }
        QTreeWidget::branch:hover { background: transparent; }
        QHeaderView::section { background-color: white; border: none; border-bottom: 1px solid #E5E7EB; padding: 8px; color: #64748B; font-weight: 600; text-align: left; }

        #rightHeader { border-bottom: 1px solid #E5E7EB; }
        #selectionSummary { color: #475569; background-color: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 8px 10px; qproperty-alignment: 'AlignCenter'; }
        #dangerHint { color: #B42318; background-color: #FEF3F2; border: 1px solid #FECACA; border-radius: 8px; padding: 8px 10px; }

        #settingsScroll { border: none; background-color: transparent; margin: 0; padding: 0; }
        #settingsScroll > QWidget { background-color: transparent; }
        #settingsScroll > QWidget > QWidget { background-color: transparent; margin: 0; padding: 0; }

        QScrollBar:vertical { border: none; background: transparent; width: 8px; margin: 0px; }
        QScrollBar::handle:vertical { background: #CBD5E1; min-height: 30px; border-radius: 4px; }
        QScrollBar::handle:vertical:hover { background: #94A3B8; }
        QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0px; background: none; }
        QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical { background: transparent; }

        QScrollBar:horizontal { border: none; background: #EAF0F6; height: 14px; margin: 2px 8px 4px 8px; border-radius: 7px; }
        QScrollBar::handle:horizontal { background: #94A3B8; min-width: 44px; border-radius: 7px; margin: 1px; }
        QScrollBar::handle:horizontal:hover { background: #64748B; }
        QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal { width: 0px; background: none; }
        QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal { background: #EAF0F6; border-radius: 7px; }

        QCheckBox { outline: none; }
        QCheckBox:focus { outline: none; }
        QCheckBox::indicator { width: 16px; height: 16px; border-radius: 4px; border: 1px solid #CBD5E1; background: white; margin-top: 1px; }
        QCheckBox::indicator:checked { background: #2563EB; border-color: #2563EB; }
        #actionBtn { padding: 8px 16px; border: 1px solid #D1D5DB; border-radius: 10px; background-color: white; color: #475569; font-weight: 600; }
        #actionBtn:hover { background-color: #F8FAFC; border-color: #9CA3AF; }
        #presetBtn { padding: 6px 14px; border: 1px solid #D1D5DB; border-radius: 8px; background-color: white; color: #475569; font-weight: 600; }
        #presetBtn:hover { background-color: #F9FAFB; border-color: #9CA3AF; }
        #presetBtn:checked { background-color: #E8F1FF; border-color: #93C5FD; color: #155EEF; }
        #presetBtn:focus { outline: none; }
        #footerSummary { color: #0F172A; font-weight: 700; }
        #processingHint { color: #155EEF; min-width: 0px; }
        #footerHint { color: #64748B; min-width: 0px; }
        #footerHint[danger="true"] { color: #B42318; font-weight: 600; }
        #startBtn { padding: 10px 24px; background-color: #2563EB; color: white; border-radius: 10px; font-weight: bold; border: none; }
        #startBtn[stopMode="true"] { background-color: #DC2626; }
        #startBtn[stopMode="true"]:hover { background-color: #B91C1C; }
        #startBtn[stopMode="true"]:pressed { background-color: #991B1B; }
        #startBtn:hover { background-color: #1D4ED8; }
        #startBtn:pressed { background-color: #1E40AF; }
        """
        self.setStyleSheet(qss)
