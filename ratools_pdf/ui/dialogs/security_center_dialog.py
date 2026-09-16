"""PDF 安全与密码/权限控制中心：
- 选项卡 1：🔒 PDF 加密与权限保护（256-bit/128-bit AES 加密、打开与权限密码、精细权限矩阵、明文元数据保留）
- 选项卡 2：🔓 PDF 解密与权限脱壳（支持智能状态探针检测、仅受限免密脱壳、已知密码彻底解锁、多文件批量解密）
"""

import os
from typing import List, Optional, Tuple

from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtWidgets import (
    QAbstractItemView,
    QButtonGroup,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QRadioButton,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from ratools_pdf.pdf import qpdf
from ratools_pdf.ui.dialogs.base import FramelessDraggableDialog


class SecurityWorker(QThread):
    """后台执行 PDF 加密任务的异步工作线程。"""

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


class DecryptWorker(QThread):
    """后台执行批量解密脱壳任务的异步工作线程。"""

    progress = Signal(str)
    file_finished = Signal(int, bool, str)  # row_idx, success, message
    finished_all = Signal(int, int, str)    # success_count, fail_count, output_dir

    def __init__(
        self,
        tasks: List[Tuple[int, str]],  # (row_idx, file_path)
        output_dir: str,
        password: Optional[str] = None,
        suffix: str = "_decrypted",
        linearize: bool = True,
        object_streams: bool = True,
        parent=None,
    ):
        super().__init__(parent)
        self.tasks = tasks
        self.output_dir = output_dir
        self.password = password
        self.suffix = suffix
        self.linearize = linearize
        self.object_streams = object_streams

    def run(self):
        success_cnt = 0
        fail_cnt = 0
        total = len(self.tasks)
        os.makedirs(self.output_dir, exist_ok=True)

        for i, (row_idx, in_path) in enumerate(self.tasks):
            base_name = os.path.basename(in_path)
            name_no_ext, ext = os.path.splitext(base_name)
            self.progress.emit(f"正在解密 ({i+1}/{total}): {base_name}...")

            out_name = f"{name_no_ext}{self.suffix}{ext}" if self.suffix else base_name
            out_path = os.path.join(self.output_dir, out_name)

            is_same = os.path.abspath(in_path) == os.path.abspath(out_path)
            target_out = (out_path + ".tmp.pdf") if is_same else out_path

            try:
                res = qpdf.decrypt_pdf(
                    input_pdf=in_path,
                    output_pdf=target_out,
                    password=self.password,
                    linearize=self.linearize,
                    object_streams="generate" if self.object_streams else None,
                )
                if res.is_success:
                    if is_same:
                        os.replace(target_out, out_path)
                    success_cnt += 1
                    self.file_finished.emit(row_idx, True, "解密成功")
                else:
                    if is_same and os.path.exists(target_out):
                        os.remove(target_out)
                    fail_cnt += 1
                    self.file_finished.emit(row_idx, False, res.stderr or "解密失败")
            except Exception as e:
                if is_same and os.path.exists(target_out):
                    try:
                        os.remove(target_out)
                    except Exception:
                        pass
                fail_cnt += 1
                self.file_finished.emit(row_idx, False, str(e))

        self.finished_all.emit(success_cnt, fail_cnt, self.output_dir)


class SecurityCenterDialog(FramelessDraggableDialog):
    """PDF 安全与密码/权限控制中心对话框（加密 + 解密脱壳）。"""

    def __init__(self, initial_file: Optional[str] = None, parent=None):
        super().__init__("🔒 PDF 安全与密码/权限控制中心", parent)
        self.resize(800, 650)
        self.setMinimumSize(740, 560)
        self.encrypt_worker: Optional[SecurityWorker] = None
        self.decrypt_worker: Optional[DecryptWorker] = None
        self.last_output_path = ""
        self.last_decrypt_dir = ""

        self.content_layout.setSpacing(10)

        self.tabs = QTabWidget()
        self._build_decrypt_tab()
        self._build_encrypt_tab(initial_file)
        self.content_layout.addWidget(self.tabs, stretch=1)

        self._build_footer()

        if initial_file and os.path.exists(initial_file):
            self.add_decrypt_files([initial_file])

    # =========================================================================
    # 选项卡 1：🔒 PDF 加密与权限保护
    # =========================================================================
    def _build_encrypt_tab(self, initial_file: Optional[str]):
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(8, 12, 8, 8)
        layout.setSpacing(12)

        # 1. 文件选择
        picker_card = QFrame()
        picker_card.setObjectName("wizardCard")
        grid = QGridLayout(picker_card)
        grid.setContentsMargins(14, 12, 14, 12)
        grid.setSpacing(10)

        grid.addWidget(QLabel("源 PDF 文件："), 0, 0)
        self.txt_input = QLineEdit()
        self.txt_input.setObjectName("settingsPathEdit")
        self.txt_input.setFixedHeight(34)
        self.txt_input.setPlaceholderText("选择需要加密保护的 PDF 文件...")
        self.txt_input.textChanged.connect(self._auto_fill_output)
        self.btn_browse_input = QPushButton("浏览...")
        self.btn_browse_input.setObjectName("dialogSecondaryBtn")
        self.btn_browse_input.setFixedHeight(34)
        self.btn_browse_input.setCursor(Qt.PointingHandCursor)
        self.btn_browse_input.clicked.connect(self._on_browse_input_clicked)
        grid.addWidget(self.txt_input, 0, 1)
        grid.addWidget(self.btn_browse_input, 0, 2)

        grid.addWidget(QLabel("输出另存为："), 1, 0)
        self.txt_output = QLineEdit()
        self.txt_output.setObjectName("settingsPathEdit")
        self.txt_output.setFixedHeight(34)
        self.txt_output.setPlaceholderText("加密后的 PDF 另存路径...")
        self.btn_browse_output = QPushButton("浏览...")
        self.btn_browse_output.setObjectName("dialogSecondaryBtn")
        self.btn_browse_output.setFixedHeight(34)
        self.btn_browse_output.setCursor(Qt.PointingHandCursor)
        self.btn_browse_output.clicked.connect(self._on_browse_output_clicked)
        grid.addWidget(self.txt_output, 1, 1)
        grid.addWidget(self.btn_browse_output, 1, 2)
        layout.addWidget(picker_card)

        # 2. 算法与密码
        crypto_card = QFrame()
        crypto_card.setObjectName("wizardCard")
        vbox_crypto = QVBoxLayout(crypto_card)
        vbox_crypto.setContentsMargins(14, 12, 14, 12)
        vbox_crypto.setSpacing(10)

        row_algo = QHBoxLayout()
        row_algo.addWidget(QLabel("加密算法："))
        self.combo_algo = QComboBox()
        self.combo_algo.setFixedHeight(34)
        self.combo_algo.addItem("256-bit AES (推荐：最高安全等级，Acrobat X 及更高版本)", 256)
        self.combo_algo.addItem("128-bit AES (兼容旧版阅读器)", 128)
        row_algo.addWidget(self.combo_algo, stretch=1)
        vbox_crypto.addLayout(row_algo)

        grid_pass = QGridLayout()
        grid_pass.setSpacing(8)

        self.cb_user_pass = QCheckBox("设置打开密码 (User Password)：")
        self.txt_user_pass = QLineEdit()
        self.txt_user_pass.setObjectName("settingsPathEdit")
        self.txt_user_pass.setFixedHeight(32)
        self.txt_user_pass.setEchoMode(QLineEdit.Password)
        self.txt_user_pass.setPlaceholderText("留空表示任何人均可直接打开阅读")
        grid_pass.addWidget(self.cb_user_pass, 0, 0)
        grid_pass.addWidget(self.txt_user_pass, 0, 1)

        self.cb_owner_pass = QCheckBox("设置权限密码 (Owner Password)：")
        self.cb_owner_pass.setChecked(True)
        self.txt_owner_pass = QLineEdit()
        self.txt_owner_pass.setObjectName("settingsPathEdit")
        self.txt_owner_pass.setFixedHeight(32)
        self.txt_owner_pass.setEchoMode(QLineEdit.Password)
        self.txt_owner_pass.setPlaceholderText("用于限制或修改打印/编辑等权限")
        grid_pass.addWidget(self.cb_owner_pass, 1, 0)
        grid_pass.addWidget(self.txt_owner_pass, 1, 1)
        vbox_crypto.addLayout(grid_pass)
        layout.addWidget(crypto_card)

        # 3. 权限矩阵
        perm_card = QFrame()
        perm_card.setObjectName("wizardCard")
        vbox_perm = QVBoxLayout(perm_card)
        vbox_perm.setContentsMargins(14, 12, 14, 12)
        vbox_perm.setSpacing(10)

        lbl_perm = QLabel("权限细则控制（限制未取得权限密码的普通用户操作）：")
        lbl_perm.setStyleSheet("font-weight: 700;")
        vbox_perm.addWidget(lbl_perm)

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
        vbox_perm.addLayout(h_print)

        h_mod = QHBoxLayout()
        h_mod.addWidget(QLabel("修改权限："))
        self.combo_modify = QComboBox()
        self.combo_modify.setFixedHeight(34)
        self.combo_modify.addItem("完全禁止修改 (推荐定稿归档)", "none")
        self.combo_modify.addItem("仅允许页面装配 (插入、删除、旋转页面)", "assembly")
        self.combo_modify.addItem("仅允许填写表单字段与电子签名", "form")
        self.combo_modify.addItem("允许批注与填写表单", "annotate")
        self.combo_modify.addItem("允许完全修改", "all")
        h_mod.addWidget(self.combo_modify, stretch=1)
        vbox_perm.addLayout(h_mod)

        h_extract = QHBoxLayout()
        self.cb_extract = QCheckBox("允许复制与提取文本/图像内容 (Extract)")
        self.cb_accessibility = QCheckBox("允许屏幕阅读器辅助功能 (Accessibility)")
        self.cb_accessibility.setChecked(True)
        h_extract.addWidget(self.cb_extract)
        h_extract.addWidget(self.cb_accessibility)
        h_extract.addStretch()
        vbox_perm.addLayout(h_extract)
        layout.addWidget(perm_card)

        # 4. 合规与优化
        row_opts = QHBoxLayout()
        row_opts.setSpacing(16)
        self.cb_cleartext_meta = QCheckBox(
            "保持元数据明文未加密 (Cleartext Metadata，便于 Windows 检索与档案库索引)"
        )
        self.cb_cleartext_meta.setChecked(True)
        self.cb_linearize = QCheckBox("启用 Web 快速视图 (线性化)")
        self.cb_linearize.setChecked(True)
        row_opts.addWidget(self.cb_cleartext_meta)
        row_opts.addWidget(self.cb_linearize)
        row_opts.addStretch()
        layout.addLayout(row_opts)

        layout.addStretch()
        self.tabs.addTab(tab, "🔒 PDF 加密与权限保护")

        if initial_file and os.path.exists(initial_file):
            self.txt_input.setText(initial_file)

    # =========================================================================
    # 选项卡 2：🔓 PDF 解密与权限脱壳
    # =========================================================================
    def _build_decrypt_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(8, 12, 8, 8)
        layout.setSpacing(10)

        # 1. 待解密文件表格
        table_container = QWidget()
        h_table = QHBoxLayout(table_container)
        h_table.setContentsMargins(0, 0, 0, 0)
        h_table.setSpacing(10)

        self.decrypt_table = QTableWidget()
        self.decrypt_table.setObjectName("previewTable")
        self.decrypt_table.setColumnCount(6)
        self.decrypt_table.setHorizontalHeaderLabels(
            ["#", "文件名", "总页数", "当前安全状态", "诊断建议", "完整路径"]
        )
        self.decrypt_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Fixed)
        self.decrypt_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.decrypt_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.decrypt_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.decrypt_table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeToContents)
        self.decrypt_table.setColumnWidth(0, 40)
        self.decrypt_table.setColumnHidden(5, True)
        self.decrypt_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.decrypt_table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.decrypt_table.setAlternatingRowColors(True)

        self.decrypt_table.setAcceptDrops(True)
        self.decrypt_table.dragEnterEvent = self._decrypt_drag_enter
        self.decrypt_table.dragMoveEvent = self._decrypt_drag_move
        self.decrypt_table.dropEvent = self._decrypt_drop

        h_table.addWidget(self.decrypt_table, stretch=1)

        v_btn_bar = QVBoxLayout()
        v_btn_bar.setSpacing(8)
        self.btn_dec_add = QPushButton("+ 添加文件")
        self.btn_dec_add.setObjectName("dialogSecondaryBtn")
        self.btn_dec_add.setFixedWidth(100)
        self.btn_dec_add.setFixedHeight(30)
        self.btn_dec_add.setCursor(Qt.PointingHandCursor)
        self.btn_dec_add.clicked.connect(self._on_dec_add_clicked)

        self.btn_dec_remove = QPushButton("- 移除选中")
        self.btn_dec_remove.setObjectName("dialogSecondaryBtn")
        self.btn_dec_remove.setFixedWidth(100)
        self.btn_dec_remove.setFixedHeight(30)
        self.btn_dec_remove.setCursor(Qt.PointingHandCursor)
        self.btn_dec_remove.clicked.connect(self._on_dec_remove_clicked)

        self.btn_dec_clear = QPushButton("清空")
        self.btn_dec_clear.setObjectName("dialogSecondaryBtn")
        self.btn_dec_clear.setFixedWidth(100)
        self.btn_dec_clear.setFixedHeight(30)
        self.btn_dec_clear.setCursor(Qt.PointingHandCursor)
        self.btn_dec_clear.clicked.connect(self._on_dec_clear_clicked)

        v_btn_bar.addWidget(self.btn_dec_add)
        v_btn_bar.addWidget(self.btn_dec_remove)
        v_btn_bar.addWidget(self.btn_dec_clear)
        v_btn_bar.addStretch()
        h_table.addLayout(v_btn_bar)
        layout.addWidget(table_container, stretch=1)

        # 2. 解密密码配置卡片
        pass_card = QFrame()
        pass_card.setObjectName("wizardCard")
        v_pass = QVBoxLayout(pass_card)
        v_pass.setContentsMargins(14, 10, 14, 10)
        v_pass.setSpacing(8)

        row_pw = QHBoxLayout()
        row_pw.setSpacing(10)
        self.cb_dec_use_pass = QCheckBox("包含已知打开/权限密码：")
        self.txt_dec_pass = QLineEdit()
        self.txt_dec_pass.setObjectName("settingsPathEdit")
        self.txt_dec_pass.setFixedHeight(32)
        self.txt_dec_pass.setEchoMode(QLineEdit.Password)
        self.txt_dec_pass.setPlaceholderText("批量文件若包含统一打开密码可在此输入，否则留空自动免密脱壳")
        row_pw.addWidget(self.cb_dec_use_pass)
        row_pw.addWidget(self.txt_dec_pass, stretch=1)
        v_pass.addLayout(row_pw)

        hint_lbl = QLabel(
            "💡 说明：对于仅受权限限制（禁止打印/编辑/复制）但无打开密码的 PDF，无需输入密码即可直接免密脱壳；\n"
            "若文档受打开密码保护，必须勾选并输入密码方可成功解锁解密。"
        )
        hint_lbl.setStyleSheet("color: #6B7280; font-size: 11px; line-height: 140%;")
        hint_lbl.setWordWrap(True)
        v_pass.addWidget(hint_lbl)
        layout.addWidget(pass_card)

        # 3. 输出与优化设置卡片
        out_card = QFrame()
        out_card.setObjectName("wizardCard")
        v_out = QVBoxLayout(out_card)
        v_out.setContentsMargins(14, 10, 14, 10)
        v_out.setSpacing(8)

        row_out_dir = QHBoxLayout()
        row_out_dir.addWidget(QLabel("解密输出目录："))
        self.txt_dec_out_dir = QLineEdit()
        self.txt_dec_out_dir.setObjectName("settingsPathEdit")
        self.txt_dec_out_dir.setFixedHeight(34)
        self.txt_dec_out_dir.setPlaceholderText("选择解密后 PDF 存放的文件夹目录...")
        self.btn_dec_browse_dir = QPushButton("浏览...")
        self.btn_dec_browse_dir.setObjectName("dialogSecondaryBtn")
        self.btn_dec_browse_dir.setFixedHeight(34)
        self.btn_dec_browse_dir.setCursor(Qt.PointingHandCursor)
        self.btn_dec_browse_dir.clicked.connect(self._on_dec_browse_dir_clicked)
        row_out_dir.addWidget(self.txt_dec_out_dir, stretch=1)
        row_out_dir.addWidget(self.btn_dec_browse_dir)
        v_out.addLayout(row_out_dir)

        row_naming = QHBoxLayout()
        row_naming.addWidget(QLabel("文件命名："))
        self.naming_group = QButtonGroup(self)
        self.rb_name_suffix = QRadioButton("自动追加 '_decrypted' 后缀 (推荐，保留原件)")
        self.rb_name_origin = QRadioButton("保持原文件名 (覆盖输出)")
        self.rb_name_suffix.setChecked(True)
        self.naming_group.addButton(self.rb_name_suffix, 0)
        self.naming_group.addButton(self.rb_name_origin, 1)
        row_naming.addWidget(self.rb_name_suffix)
        row_naming.addWidget(self.rb_name_origin)
        row_naming.addStretch()
        v_out.addLayout(row_naming)

        row_opts2 = QHBoxLayout()
        self.cb_dec_linearize = QCheckBox("启用 Web 快速视图 (线性化)")
        self.cb_dec_linearize.setChecked(True)
        self.cb_dec_objstms = QCheckBox("压缩生成对象流 (PDF 1.5+)")
        self.cb_dec_objstms.setChecked(True)
        row_opts2.addWidget(self.cb_dec_linearize)
        row_opts2.addWidget(self.cb_dec_objstms)
        row_opts2.addStretch()
        v_out.addLayout(row_opts2)
        layout.addWidget(out_card)

        self.tabs.addTab(tab, "🔓 PDF 解密与权限脱壳")

    # =========================================================================
    # 底部状态栏与主执行按钮
    # =========================================================================
    def _build_footer(self):
        footer = QWidget()
        layout = QHBoxLayout(footer)
        layout.setContentsMargins(0, 6, 0, 0)
        layout.setSpacing(12)

        self.lbl_status = QLabel("就绪")
        self.lbl_status.setStyleSheet("color: gray; font-size: 12px;")
        layout.addWidget(self.lbl_status)
        layout.addStretch()

        self.btn_open_target = QPushButton("打开输出目录")
        self.btn_open_target.setObjectName("dialogSecondaryBtn")
        self.btn_open_target.setFixedHeight(34)
        self.btn_open_target.setCursor(Qt.PointingHandCursor)
        self.btn_open_target.setVisible(False)
        self.btn_open_target.clicked.connect(self._on_open_target_clicked)
        layout.addWidget(self.btn_open_target)

        self.btn_cancel = QPushButton("关闭")
        self.btn_cancel.setObjectName("dialogSecondaryBtn")
        self.btn_cancel.setFixedHeight(34)
        self.btn_cancel.setCursor(Qt.PointingHandCursor)
        self.btn_cancel.clicked.connect(self.reject)
        layout.addWidget(self.btn_cancel)

        self.btn_execute = QPushButton("开始处理")
        self.btn_execute.setObjectName("dialogPrimaryBtn")
        self.btn_execute.setFixedHeight(34)
        self.btn_execute.setCursor(Qt.PointingHandCursor)
        self.btn_execute.clicked.connect(self._on_execute_clicked)
        layout.addWidget(self.btn_execute)

        self.content_layout.addWidget(footer)
        self.tabs.currentChanged.connect(self._on_tab_changed)
        self._on_tab_changed(0)

    def _on_tab_changed(self, idx: int):
        self.btn_open_target.setVisible(False)
        if idx == 0:
            self.btn_execute.setText("开始批量解密")
            cnt = self.decrypt_table.rowCount()
            self.lbl_status.setText(f"已选 {cnt} 个待解密文件" if cnt > 0 else "请添加待解密 PDF 文件")
        else:
            self.btn_execute.setText("应用安全加密")
            self.lbl_status.setText("就绪 (加密配置模式)")

    # =========================================================================
    # 加密业务逻辑
    # =========================================================================
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

    def _execute_encryption(self):
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

        self.encrypt_worker = SecurityWorker(
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
        self.encrypt_worker.progress.connect(lambda msg: self.lbl_status.setText(msg))
        self.encrypt_worker.finished_signal.connect(self._on_encrypt_finished)
        self.encrypt_worker.start()

    def _on_encrypt_finished(self, success: bool, message: str, target_path: str):
        self.btn_execute.setEnabled(True)
        self.btn_cancel.setEnabled(True)
        self.lbl_status.setText("加密完成" if success else "加密失败")

        if success:
            self.last_output_path = target_path
            self.last_decrypt_dir = os.path.dirname(target_path)
            self.btn_open_target.setVisible(True)
            QMessageBox.information(self, "完成", message)
        else:
            QMessageBox.critical(self, "错误", message)

    # =========================================================================
    # 解密脱壳业务逻辑
    # =========================================================================
    def add_decrypt_files(self, file_paths: List[str]):
        """向解密表格添加文件并执行探针状态诊断。"""
        pdf_paths = [p for p in file_paths if p.lower().endswith(".pdf") and os.path.exists(p)]
        if not pdf_paths:
            return

        self.decrypt_table.blockSignals(True)
        for p in pdf_paths:
            row = self.decrypt_table.rowCount()
            self.decrypt_table.insertRow(row)

            # 0: 序号
            it_idx = QTableWidgetItem(str(row + 1))
            it_idx.setTextAlignment(Qt.AlignCenter)
            it_idx.setFlags(it_idx.flags() & ~Qt.ItemIsEditable)
            self.decrypt_table.setItem(row, 0, it_idx)

            # 1: 文件名
            it_name = QTableWidgetItem(os.path.basename(p))
            it_name.setToolTip(p)
            it_name.setFlags(it_name.flags() & ~Qt.ItemIsEditable)
            self.decrypt_table.setItem(row, 1, it_name)

            # 2: 页数
            pages = qpdf.get_pdf_page_count(p)
            it_p = QTableWidgetItem(str(pages))
            it_p.setTextAlignment(Qt.AlignCenter)
            it_p.setFlags(it_p.flags() & ~Qt.ItemIsEditable)
            self.decrypt_table.setItem(row, 2, it_p)

            # 探针检测安全状态
            probe = qpdf.probe_pdf_encryption_status(p)
            status_text = probe.get("title", "")
            desc_text = probe.get("description", "")
            status_code = probe.get("status", "")

            # 3: 当前状态
            it_status = QTableWidgetItem(status_text)
            it_status.setTextAlignment(Qt.AlignCenter)
            if status_code == "password_required":
                it_status.setForeground(Qt.red)
            elif status_code == "restricted_no_password":
                it_status.setForeground(Qt.darkYellow)
            else:
                it_status.setForeground(Qt.darkGreen)
            it_status.setFlags(it_status.flags() & ~Qt.ItemIsEditable)
            self.decrypt_table.setItem(row, 3, it_status)

            # 4: 诊断建议
            it_desc = QTableWidgetItem(desc_text)
            it_desc.setFlags(it_desc.flags() & ~Qt.ItemIsEditable)
            self.decrypt_table.setItem(row, 4, it_desc)

            # 5: 完整路径
            it_path = QTableWidgetItem(p)
            self.decrypt_table.setItem(row, 5, it_path)

        self.decrypt_table.blockSignals(False)

        # 刷新序号与默认输出目录
        for r in range(self.decrypt_table.rowCount()):
            it = self.decrypt_table.item(r, 0)
            if it:
                it.setText(str(r + 1))

        if not self.txt_dec_out_dir.text() and self.decrypt_table.rowCount() > 0:
            first_p = self.decrypt_table.item(0, 5).text()
            self.txt_dec_out_dir.setText(os.path.dirname(first_p))

        cnt = self.decrypt_table.rowCount()
        if hasattr(self, "lbl_status"):
            self.lbl_status.setText(f"已选 {cnt} 个待解密文件")

    def _on_dec_add_clicked(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "选择待解密 PDF 文件", "", "PDF 文件 (*.pdf)"
        )
        if files:
            self.add_decrypt_files(files)

    def _on_dec_remove_clicked(self):
        row = self.decrypt_table.currentRow()
        if row >= 0:
            self.decrypt_table.removeRow(row)
            for r in range(self.decrypt_table.rowCount()):
                it = self.decrypt_table.item(r, 0)
                if it:
                    it.setText(str(r + 1))
            cnt = self.decrypt_table.rowCount()
            self.lbl_status.setText(f"已选 {cnt} 个待解密文件" if cnt > 0 else "请添加待解密 PDF 文件")

    def _on_dec_clear_clicked(self):
        self.decrypt_table.setRowCount(0)
        self.lbl_status.setText("已清空列表")

    def _on_dec_browse_dir_clicked(self):
        folder = QFileDialog.getExistingDirectory(self, "选择解密输出文件夹")
        if folder:
            self.txt_dec_out_dir.setText(folder)

    def _decrypt_drag_enter(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def _decrypt_drag_move(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def _decrypt_drop(self, event):
        urls = event.mimeData().urls()
        files = [u.toLocalFile() for u in urls if u.toLocalFile().lower().endswith(".pdf")]
        if files:
            self.add_decrypt_files(files)
            event.acceptProposedAction()

    def _execute_decryption(self):
        if self.decrypt_table.rowCount() == 0:
            QMessageBox.warning(self, "提示", "请先添加至少一个待解密 PDF 文件！")
            return

        out_dir = self.txt_dec_out_dir.text().strip()
        if not out_dir:
            QMessageBox.warning(self, "提示", "请指定解密输出保存目录！")
            return

        password = self.txt_dec_pass.text().strip() if self.cb_dec_use_pass.isChecked() else None
        suffix = "_decrypted" if self.rb_name_suffix.isChecked() else ""

        tasks = []
        for r in range(self.decrypt_table.rowCount()):
            tasks.append((r, self.decrypt_table.item(r, 5).text()))

        self.btn_execute.setEnabled(False)
        self.btn_cancel.setEnabled(False)
        self.btn_open_target.setVisible(False)
        self.lbl_status.setText("⏳ 正在批量解密脱壳...")

        self.decrypt_worker = DecryptWorker(
            tasks=tasks,
            output_dir=out_dir,
            password=password,
            suffix=suffix,
            linearize=self.cb_dec_linearize.isChecked(),
            object_streams=self.cb_dec_objstms.isChecked(),
            parent=self,
        )
        self.decrypt_worker.progress.connect(lambda msg: self.lbl_status.setText(msg))
        self.decrypt_worker.file_finished.connect(self._on_dec_file_finished)
        self.decrypt_worker.finished_all.connect(self._on_dec_finished_all)
        self.decrypt_worker.start()

    def _on_dec_file_finished(self, row_idx: int, success: bool, msg: str):
        it_status = self.decrypt_table.item(row_idx, 3)
        it_desc = self.decrypt_table.item(row_idx, 4)
        if it_status:
            it_status.setText("✅ 已解密" if success else "❌ 失败")
            it_status.setForeground(Qt.darkGreen if success else Qt.red)
        if it_desc:
            it_desc.setText(msg)

    def _on_dec_finished_all(self, success_cnt: int, fail_cnt: int, out_dir: str):
        self.btn_execute.setEnabled(True)
        self.btn_cancel.setEnabled(True)
        self.lbl_status.setText(f"解密完成：成功 {success_cnt} 个，失败 {fail_cnt} 个")
        self.last_decrypt_dir = out_dir
        self.btn_open_target.setVisible(True)

        if fail_cnt == 0:
            QMessageBox.information(
                self, "完成", f"✅ 全部 {success_cnt} 个文档已成功解密脱壳！\n保存至：{out_dir}"
            )
        else:
            QMessageBox.warning(
                self,
                "部分完成",
                f"解密处理结束：成功 {success_cnt} 个，失败 {fail_cnt} 个。\n请检查失败文档是否需要指定有效打开密码。",
            )

    # =========================================================================
    # 统一分发执行
    # =========================================================================
    def _on_execute_clicked(self):
        if self.tabs.currentIndex() == 0:
            self._execute_decryption()
        else:
            self._execute_encryption()

    def _on_open_target_clicked(self):
        target = self.last_decrypt_dir or (
            os.path.dirname(self.last_output_path) if self.last_output_path else ""
        )
        if target and os.path.exists(target):
            os.startfile(target)
