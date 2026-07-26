
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QFrame, QHBoxLayout, QLabel, QPushButton, QVBoxLayout,
)

from ratools_pdf.config.features import ENABLE_UPDATE_CHECK
from ratools_pdf.config.version import get_display_version
from ratools_pdf.ui.log_view_model import (
    build_log_summary_items,
    filter_log_summary_items,
)
from ratools_pdf.ui.dialogs.base import FramelessDraggableDialog


class AboutDialog(FramelessDraggableDialog):
    def __init__(self, parent=None):
        super().__init__("ℹ️ 关于软件", parent)
        self.resize(520, 460)
        self.content_layout.setSpacing(16)
        self.latest_release_url = ""

        hero_card = QFrame()
        hero_card.setObjectName("aboutHeroCard")
        hero_layout = QVBoxLayout(hero_card)
        hero_layout.setContentsMargins(18, 16, 18, 16)
        hero_layout.setSpacing(6)

        brand_title = QLabel("RATools for PDF")
        brand_title.setObjectName("aboutBrandTitle")
        version_badge = QLabel(get_display_version())
        version_badge.setObjectName("aboutBadge")
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
        intro_text.setObjectName("aboutIntro")
        self.content_layout.addWidget(intro_text)

        features_title = QLabel("核心功能")
        features_title.setObjectName("aboutTitle")
        self.content_layout.addWidget(features_title)

        features_text = QLabel(
            "• 批量导入PDF文件或文件夹\n"
            "• 按模块勾选规则，支持中国eCTD/美国eCTD预设\n"
            "• 覆盖文档属性、书签、链接、动态内容与附件等常见合规项\n"
            "• 输出处理日志，便于复核与追踪"
        )
        features_text.setWordWrap(True)
        features_text.setObjectName("aboutText")
        self.content_layout.addWidget(features_text)

        tech_card = QFrame()
        tech_card.setObjectName("aboutInfoCard")
        tech_layout = QVBoxLayout(tech_card)
        tech_layout.setContentsMargins(16, 14, 16, 14)
        tech_layout.setSpacing(4)

        tech_title = QLabel("技术与许可")
        tech_title.setObjectName("aboutTitle")
        tech_detail = QLabel(
            "基于PySide6、PyMuPDF及qpdf等项目构建\n"
            "项目源码遵循GNU AGPL v3开源协议\n"
            "第三方组件许可与源码说明见 THIRD_PARTY_NOTICES.md"
        )
        tech_detail.setWordWrap(True)
        tech_detail.setObjectName("aboutText")
        tech_layout.addWidget(tech_title)
        tech_layout.addWidget(tech_detail)
        self.content_layout.addWidget(tech_card)

        if ENABLE_UPDATE_CHECK:
            update_card = QFrame()
            update_card.setObjectName("aboutInfoCard")
            update_layout = QVBoxLayout(update_card)
            update_layout.setContentsMargins(16, 14, 16, 14)
            update_layout.setSpacing(8)

            update_title = QLabel("更新")
            update_title.setObjectName("aboutTitle")
            self.update_status_label = QLabel("可手动检查 GitHub Releases 中的最新版本。")
            self.update_status_label.setWordWrap(True)
            self.update_status_label.setObjectName("aboutText")

            update_button_layout = QHBoxLayout()
            update_button_layout.setContentsMargins(0, 4, 0, 0)
            update_button_layout.setSpacing(8)
            self.btn_check_updates = QPushButton("检查更新")
            self.btn_check_updates.setObjectName("dialogPrimaryBtn")
            self.btn_open_release = QPushButton("打开发布页")
            self.btn_open_release.setObjectName("dialogSecondaryBtn")
            self.btn_open_release.hide()
            update_button_layout.addWidget(self.btn_check_updates)
            update_button_layout.addWidget(self.btn_open_release)
            update_button_layout.addStretch()

            update_layout.addWidget(update_title)
            update_layout.addWidget(self.update_status_label)
            update_layout.addLayout(update_button_layout)
            self.content_layout.addWidget(update_card)

        self.content_layout.addStretch()

        btn_close = QPushButton("关 闭")
        btn_close.setObjectName("dialogSecondaryBtn")
        btn_close.setFixedHeight(36)
        btn_close.clicked.connect(self.accept)
        self.content_layout.addWidget(btn_close)

    def set_update_checking(self):
        # NoUpdate 变体（ENABLE_UPDATE_CHECK=False）不创建更新区控件
        if not hasattr(self, "btn_check_updates"):
            return
        self.btn_check_updates.setEnabled(False)
        self.update_status_label.setText("正在检查更新...")
        self.btn_open_release.hide()

    def set_update_result(self, message, release_url=""):
        if not hasattr(self, "btn_check_updates"):
            return
        self.btn_check_updates.setEnabled(True)
        self.update_status_label.setText(message)
        self.latest_release_url = release_url
        self.btn_open_release.setVisible(bool(release_url))


# ================== 自定义组件与主窗口 ==================
