"""多选一提示对话框：重大更新提醒、已签名文件处理确认。

两者共用"若干按钮各代表一个动作，关闭即返回动作字符串"的模式，
由 MainWindow 的包装方法调用（历史上内联在主窗口里）。
"""

import os

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QFrame, QHBoxLayout, QLabel, QPushButton, QScrollArea, QVBoxLayout

from ratools_pdf.ui.dialogs.base import FramelessDraggableDialog


class _ActionChoiceDialog(FramelessDraggableDialog):
    """底部按钮各绑定一个动作字符串；exec 后读 chosen_action。"""

    def __init__(self, title_text, default_action, parent=None):
        super().__init__(title_text, parent)
        self.chosen_action = default_action

    def add_action_buttons(self, buttons):
        """buttons: [(text, action, objectName), ...] 依序排列，最后一个通常是主按钮。"""
        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(10)
        btn_layout.addStretch()
        for text, action, object_name in buttons:
            btn = QPushButton(text)
            btn.setObjectName(object_name)
            btn.setFixedHeight(36)
            btn.setCursor(Qt.PointingHandCursor)
            btn.clicked.connect(lambda _checked=False, a=action: self._finish(a))
            btn_layout.addWidget(btn)
        self.content_layout.addLayout(btn_layout)

    def _finish(self, action):
        self.chosen_action = action
        self.accept()


class MajorUpdatePromptDialog(_ActionChoiceDialog):
    """重大更新提醒：返回 "open" / "later" / "ignore"。"""

    def __init__(self, current_version, release, parent=None):
        super().__init__("重要更新可用", default_action="later", parent=parent)
        self.resize(460, 300)
        self.content_layout.setSpacing(14)

        message = QLabel(
            f"发现重要更新：{release.version_text}\n"
            f"当前版本：{current_version}\n"
            f"发布标题：{release.title}\n"
            f"发布时间：{release.published_at or '未知'}"
        )
        message.setWordWrap(True)
        message.setObjectName("aboutText")
        self.content_layout.addWidget(message)
        self.content_layout.addStretch()

        self.add_action_buttons([
            ("忽略此版本", "ignore", "dialogSecondaryBtn"),
            ("稍后提醒", "later", "dialogSecondaryBtn"),
            ("查看更新", "open", "dialogPrimaryBtn"),
        ])


class SignedFilesPromptDialog(_ActionChoiceDialog):
    """已签名文件三选一：返回 "skip" / "process_all" / "cancel"。"""

    def __init__(self, signed_files, parent=None):
        super().__init__("检测到已签名文件", default_action="cancel", parent=parent)
        count = len(signed_files)
        self.content_layout.setSpacing(16)

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

        self.content_layout.addWidget(warn_card)

        # ---- 文件列表标题 ----
        caption = QLabel("已签名文件：")
        caption.setObjectName("signedListCaption")
        self.content_layout.addWidget(caption)

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
        self.content_layout.addWidget(scroll)

        # ---- 底部操作提示 ----
        hint = QLabel("请选择如何处理这些已签名文件：")
        hint.setObjectName("dialogMuted")
        self.content_layout.addWidget(hint)
        self.content_layout.addStretch()

        self.add_action_buttons([
            ("取消处理", "cancel", "dialogSecondaryBtn"),
            ("仍然处理全部", "process_all", "dialogSecondaryBtn"),
            ("跳过已签名文件", "skip", "dialogPrimaryBtn"),
        ])

        self.adjustSize()
        self.setMinimumWidth(500)
