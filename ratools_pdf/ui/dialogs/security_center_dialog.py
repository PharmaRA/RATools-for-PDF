"""PDF 安全与加密中心：256-bit/128-bit AES 加密、打开与权限密码、精细权限矩阵。"""

import os
from typing import Optional

from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtWidgets import (
    QButtonGroup,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QRadioButton,
    QVBoxLayout,
    QWidget,
)

from ratools_pdf.pdf import qpdf
from ratools_pdf.ui.dialogs.base import FramelessDraggableDialog


class SecurityWorker(QThread):
    """后台执行 PDF 加密与权限保护任务的异步工作线程。"""

    progress = Signal(str)
    finished_signal = Signal(bool, str, str)  # success, message, target_path

    def __init__(
        self,
        input_pdf: str,
        output_pdf: str,
        user_password: Optional[str],
        owner_password: Optional[str],
        bits: int,
        print_perm: str,
        modify_perm: str,
        extract: bool,
        accessibility: bool,
        cleartext_metadata: bool,
        linearize: bool,
        parent=None,
    ):
        super().__init__(parent)
        self.input_pdf = input_pdf
        self.output_pdf = output_pdf
        self.user_password = user_password
        self.owner_password = owner_password
        self.bits = bits
        self.print_perm = print_perm
        self.modify_perm = modify_perm
        self.extract = extract
        self.accessibility = accessibility
        self.cleartext_metadata = cleartext_metadata
        self.linearize = linearize

    def run(self):
        try:
            self.progress.emit("正在应用安全加密配置...")
            res = qpdf.encrypt_pdf(
                input_pdf=self.input_pdf,
                output_pdf=self.output_pdf,
                user_password=self.user_password,
                owner_password=self.owner_password,
                bits=self.bits,
                print_perm=self.print_perm,
                modify_perm=self.modify_perm,
                extract=self.extract,
                accessibility=self.accessibility,
                cleartext_metadata=self.cleartext_metadata,
                linearize=self.linearize,
            )
            if res.is_success:
                self.finished_signal.emit(
                    True, f"✅ PDF 已成功加密并保存：\n{self.output_pdf}", self.output_pdf
                )
            else:
                self.finished_signal.emit(
                    False, f"❌ 加密失败：{res.stderr or '未知错误'}", ""
                )
        except Exception as e:
            self.finished_signal.emit(False, f"❌ 加密异常：{str(e)}", "")


class SecurityCenterDialog(FramelessDraggableDialog):
    """PDF 安全与权限配置中心对话框。"""

    def __init__(self, initial_file: Optional[str] = None, parent=None):
        super().__init__("🔒 PDF 安全与权限配置中心", parent)
        self.resize(720, 620)
        self.worker: Optional[SecurityWorker] = None
        self.last_output_path = ""

        self.content_layout.setSpacing(12)

        self._build_file_picker(initial_file)
        self._build_crypto_settings()
        self._build_permission_matrix()
        self._build_compliance_options()
        self._build_footer()

    def _build_file_picker(self, initial_file: Optional[str]):
        """源文件与目标文件选择。"""
        picker_card = QFrame()
        picker_card.setObjectName("wizardCard")
        grid = QGridLayout(picker_card)
        grid.setContentsMargins(14, 12, 14, 12)
        grid.setSpacing(10)

        grid.addWidget(QLabel("源 PDF 文件："), 0, 0)
        self.txt_input = QLineEdit()
        self.txt_input.setPlaceholderText("选择需要加密保护的 PDF 文件...")
        self.txt_input.textChanged.connect(self._auto_fill_output)
        self.btn_browse_input = QPushButton("浏览...")
        self.btn_browse_input.clicked.connect(self._on_browse_input_clicked)
        grid.addWidget(self.txt_input, 0, 1)
        grid.addWidget(self.btn_browse_input, 0, 2)

        grid.addWidget(QLabel("输出另存为："), 1, 0)
        self.txt_output = QLineEdit()
        self.txt_output.setPlaceholderText("加密后的 PDF 另存路径...")
        self.btn_browse_output = QPushButton("浏览...")
        self.btn_browse_output.clicked.connect(self._on_browse_output_clicked)
        grid.addWidget(self.txt_output, 1, 1)
        grid.addWidget(self.btn_browse_output, 1, 2)

        self.content_layout.addWidget(picker_card)

        if initial_file and os.path.exists(initial_file):
            self.txt_input.setText(initial_file)

    def _build_crypto_settings(self):
        """算法选择与密码输入。"""
        crypto_card = QFrame()
        crypto_card.setObjectName("wizardCard")
        vbox = QVBoxLayout(crypto_card)
        vbox.setContentsMargins(14, 12, 14, 12)
        vbox.setSpacing(10)

        # 算法
        row_algo = QHBoxLayout()
        row_algo.addWidget(QLabel("加密算法："))
        self.combo_algo = QComboBox()
        self.combo_algo.addItem("256-bit AES (推荐：Acrobat X 及更高版本，最高安全性)", 256)
        self.combo_algo.addItem("128-bit AES (兼容旧版阅读器)", 128)
        row_algo.addWidget(self.combo_algo, stretch=1)
        vbox.addLayout(row_algo)

        # 密码
        grid_pass = QGridLayout()
        grid_pass.setSpacing(8)

        self.cb_user_pass = QCheckBox("设置打开密码 (User Password)：")
        self.txt_user_pass = QLineEdit()
        self.txt_user_pass.setEchoMode(QLineEdit.Password)
        self.txt_user_pass.setPlaceholderText("留空表示任何人均可直接打开阅读")
        grid_pass.addWidget(self.cb_user_pass, 0, 0)
        grid_pass.addWidget(self.txt_user_pass, 0, 1)

        self.cb_owner_pass = QCheckBox("设置权限密码 (Owner Password)：")
        self.cb_owner_pass.setChecked(True)
        self.txt_owner_pass = QLineEdit()
        self.txt_owner_pass.setEchoMode(QLineEdit.Password)
        self.txt_owner_pass.setPlaceholderText("用于限制或修改打印/编辑等权限")
        grid_pass.addWidget(self.cb_owner_pass, 1, 0)
        grid_pass.addWidget(self.txt_owner_pass, 1, 1)

        vbox.addLayout(grid_pass)
        self.content_layout.addWidget(crypto_card)

    def _build_permission_matrix(self):
        """精细权限控制矩阵。"""
        perm_card = QFrame()
        perm_card.setObjectName("wizardCard")
        vbox = QVBoxLayout(perm_card)
        vbox.setContentsMargins(14, 12, 14, 12)
        vbox.setSpacing(10)

        lbl = QLabel("权限细则控制（限制非授权用户的操作）：")
        lbl.setStyleSheet("font-weight: 700;")
        vbox.addWidget(lbl)

        # 打印权限
        h_print = QHBoxLayout()
        h_print.addWidget(QLabel("打印权限："))
        self.print_group = QButtonGroup(self)
        self.rb_print_full = QRadioButton("允许高质量打印")
        self.rb_print_low = QRadioButton("仅允许低分辨率 (150 dpi)")
        self.rb_print_none = QRadioButton("完全禁止打印")
        self.rb_print_full.setChecked(True)

        self.print_group.addButton(self.rb_print_full, 0)
        self.print_group.addButton(self.rb_print_low, 1)
        self.print_group.addButton(self.rb_print_none, 2)

        h_print.addWidget(self.rb_print_full)
        h_print.addWidget(self.rb_print_low)
        h_print.addWidget(self.rb_print_none)
        h_print.addStretch()
        vbox.addLayout(h_print)

        # 修改权限
        h_mod = QHBoxLayout()
        h_mod.addWidget(QLabel("修改权限："))
        self.combo_modify = QComboBox()
        self.combo_modify.addItem("完全禁止修改 (推荐定稿归档)", "none")
        self.combo_modify.addItem("仅允许页面装配 (插入、删除、旋转页面)", "assembly")
        self.combo_modify.addItem("仅允许填写表单字段与电子签名", "form")
        self.combo_modify.addItem("允许批注与填写表单", "annotate")
        self.combo_modify.addItem("允许完全修改", "all")
        h_mod.addWidget(self.combo_modify, stretch=1)
        vbox.addLayout(h_mod)

        # 提取与辅助功能
        h_extract = QHBoxLayout()
        self.cb_extract = QCheckBox("允许复制与提取文本/图像内容 (Extract)")
        self.cb_accessibility = QCheckBox("允许屏幕阅读器辅助功能 (Accessibility)")
        self.cb_accessibility.setChecked(True)
        h_extract.addWidget(self.cb_extract)
        h_extract.addWidget(self.cb_accessibility)
        h_extract.addStretch()
        vbox.addLayout(h_extract)

        self.content_layout.addWidget(perm_card)

    def _build_compliance_options(self):
        """合规与优化特性。"""
        row = QHBoxLayout()
        row.setSpacing(16)

        self.cb_cleartext_meta = QCheckBox(
            "保持元数据明文未加密 (Cleartext Metadata，便于 Windows 检索与档案库索引)"
        )
        self.cb_cleartext_meta.setChecked(True)
        self.cb_linearize = QCheckBox("启用 Web 快速视图 (线性化)")
        self.cb_linearize.setChecked(True)

        row.addWidget(self.cb_cleartext_meta)
        row.addWidget(self.cb_linearize)
        row.addStretch()
        self.content_layout.addLayout(row)

    def _build_footer(self):
        """底部状态栏与主操作按钮。"""
        footer = QWidget()
        layout = QHBoxLayout(footer)
        layout.setContentsMargins(0, 8, 0, 0)
        layout.setSpacing(12)

        self.lbl_status = QLabel("就绪")
        self.lbl_status.setStyleSheet("color: gray; font-size: 12px;")
        layout.addWidget(self.lbl_status)
        layout.addStretch()

        self.btn_open_target = QPushButton("打开目标")
        self.btn_open_target.setVisible(False)
        self.btn_open_target.clicked.connect(self._on_open_target_clicked)
        layout.addWidget(self.btn_open_target)

        self.btn_cancel = QPushButton("关闭")
        self.btn_cancel.clicked.connect(self.reject)
        layout.addWidget(self.btn_cancel)

        self.btn_execute = QPushButton("应用安全配置")
        self.btn_execute.setObjectName("choiceToggleBtn")
        self.btn_execute.setStyleSheet(
            "QPushButton { background-color: #2563EB; color: white; font-weight: bold; border-radius: 8px; padding: 8px 18px; }"
            "QPushButton:hover { background-color: #1D4ED8; }"
            "QPushButton:disabled { background-color: #93C5FD; }"
        )
        self.btn_execute.clicked.connect(self._on_execute_clicked)
        layout.addWidget(self.btn_execute)

        self.content_layout.addWidget(footer)

    def _auto_fill_output(self, in_path: str):
        if hasattr(self, "txt_output") and in_path and in_path.lower().endswith(".pdf") and not self.txt_output.text():
            dir_name = os.path.dirname(in_path)
            base_name = os.path.splitext(os.path.basename(in_path))[0]
            self.txt_output.setText(os.path.join(dir_name, f"{base_name}_encrypted.pdf"))

    def _on_browse_input_clicked(self):
        f, _ = QFileDialog.getOpenFileName(self, "选择输入 PDF 文件", "", "PDF 文件 (*.pdf)")
        if f:
            self.txt_input.setText(f)

    def _on_browse_output_clicked(self):
        f, _ = QFileDialog.getSaveFileName(
            self, "保存加密后 PDF 文件", self.txt_output.text() or "encrypted.pdf", "PDF 文件 (*.pdf)"
        )
        if f:
            self.txt_output.setText(f)

    def _on_execute_clicked(self):
        in_path = self.txt_input.text().strip()
        out_path = self.txt_output.text().strip()

        if not in_path or not os.path.exists(in_path):
            QMessageBox.warning(self, "提示", "请选择有效的输入 PDF 文件！")
            return
        if not out_path:
            QMessageBox.warning(self, "提示", "请指定输出 PDF 文件路径！")
            return

        user_pass = self.txt_user_pass.text().strip() if self.cb_user_pass.isChecked() else None
        owner_pass = self.txt_owner_pass.text().strip() if self.cb_owner_pass.isChecked() else None

        if not user_pass and not owner_pass:
            reply = QMessageBox.question(
                self,
                "确认",
                "未设置打开密码与权限密码，仅应用无密码加密策略。是否继续？",
                QMessageBox.Yes | QMessageBox.No,
            )
            if reply != QMessageBox.Yes:
                return

        bits = self.combo_algo.currentData()
        print_id = self.print_group.checkedId()
        print_perm = "full" if print_id == 0 else ("low" if print_id == 1 else "none")
        modify_perm = self.combo_modify.currentData()

        self.btn_execute.setEnabled(False)
        self.btn_cancel.setEnabled(False)
        self.lbl_status.setText("⏳ 正在加密...")
        self.btn_open_target.setVisible(False)

        self.worker = SecurityWorker(
            input_pdf=in_path,
            output_pdf=out_path,
            user_password=user_pass,
            owner_password=owner_pass,
            bits=bits,
            print_perm=print_perm,
            modify_perm=modify_perm,
            extract=self.cb_extract.isChecked(),
            accessibility=self.cb_accessibility.isChecked(),
            cleartext_metadata=self.cb_cleartext_meta.isChecked(),
            linearize=self.cb_linearize.isChecked(),
            parent=self,
        )
        self.worker.progress.connect(lambda msg: self.lbl_status.setText(msg))
        self.worker.finished_signal.connect(self._on_worker_finished)
        self.worker.start()

    def _on_worker_finished(self, success: bool, message: str, target_path: str):
        self.btn_execute.setEnabled(True)
        self.btn_cancel.setEnabled(True)
        self.lbl_status.setText("加密完成" if success else "加密失败")

        if success:
            self.last_output_path = target_path
            self.btn_open_target.setVisible(True)
            QMessageBox.information(self, "完成", message)
        else:
            QMessageBox.critical(self, "错误", message)

    def _on_open_target_clicked(self):
        if self.last_output_path and os.path.exists(self.last_output_path):
            os.startfile(os.path.dirname(self.last_output_path))
