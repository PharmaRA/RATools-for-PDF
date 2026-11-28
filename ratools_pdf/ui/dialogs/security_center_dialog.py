"""PDF 安全与密码/权限控制中心：
- 选项卡 1：🔒 PDF 加密与权限保护（256-bit/128-bit AES 加密、打开与权限密码、精细权限矩阵、明文元数据保留）
- 选项卡 2：🔓 PDF 解密与权限脱壳（支持智能状态探针检测、仅受限免密脱壳、已知密码彻底解锁、多文件批量解密）
"""

import os
import tempfile
from typing import List, Optional, Tuple

from PySide6.QtCore import QSize, Qt, QThread, QTimer, Signal
from PySide6.QtGui import QColor
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
    QPushButton,
    QRadioButton,
    QScrollArea,
    QSizePolicy,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from ratools_pdf.pdf import qpdf
from ratools_pdf.services.file_safety import check_unlinked_path, recycle_owned_path
from ratools_pdf.services.output_paths import common_base_dir, relative_source_directory, validate_output_plan
from ratools_pdf.ui import icons
from ratools_pdf.ui.dialogs.base import FramelessDraggableDialog
from ratools_pdf.ui.dialogs.message_box import CustomMessageBox as QMessageBox
from ratools_pdf.ui.theme import active_palette
from ratools_pdf.ui.widgets import ElidedLabel, NoFocusItemDelegate


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
        remove_restrictions: bool = True,
        remove_acroform: bool = False,
        linearize: bool = True,
        object_streams: bool = True,
        parent=None,
    ):
        super().__init__(parent)
        self.tasks = tasks
        self.output_dir = output_dir
        self.password = password
        self.suffix = suffix
        self.remove_restrictions = remove_restrictions
        self.remove_acroform = remove_acroform
        self.linearize = linearize
        self.object_streams = object_streams

    def run(self):
        success_cnt = 0
        fail_cnt = 0
        total = len(self.tasks)
        plan = []
        try:
            inputs = [path for _, path in self.tasks]
            common_base = common_base_dir(inputs)
            for in_path in inputs:
                name_no_ext, ext = os.path.splitext(os.path.basename(in_path))
                out_name = f"{name_no_ext}{self.suffix}{ext}" if self.suffix else os.path.basename(in_path)
                subdir = relative_source_directory(in_path, common_base)
                plan.append((in_path, os.path.join(self.output_dir, subdir, out_name)))
            validate_output_plan(plan, inputs, allow_in_place=not self.suffix)
            os.makedirs(self.output_dir, exist_ok=True)
        except Exception as exc:
            for row_idx, _ in self.tasks:
                self.file_finished.emit(row_idx, False, str(exc))
            self.finished_all.emit(0, total, self.output_dir)
            return

        for i, ((row_idx, in_path), (_, out_path)) in enumerate(zip(self.tasks, plan)):
            base_name = os.path.basename(in_path)
            self.progress.emit(f"正在解密 ({i+1}/{total}): {base_name}...")
            staging_dir = None
            try:
                check_unlinked_path(out_path)
                parent_dir = os.path.dirname(os.path.abspath(out_path))
                os.makedirs(parent_dir, exist_ok=True)
                staging_dir = tempfile.mkdtemp(prefix='.ratools-decrypt-', dir=parent_dir)
                target_out = os.path.join(staging_dir, 'result.pdf')
                res = qpdf.decrypt_pdf(
                    input_pdf=in_path,
                    output_pdf=target_out,
                    password=self.password,
                    remove_restrictions=self.remove_restrictions,
                    remove_acroform=self.remove_acroform,
                    linearize=self.linearize,
                    object_streams="generate" if self.object_streams else None,
                )
                if res.is_success:
                    check_unlinked_path(out_path)
                    os.replace(target_out, out_path)
                    success_cnt += 1
                    self.file_finished.emit(row_idx, True, "解密成功")
                else:
                    fail_cnt += 1
                    self.file_finished.emit(row_idx, False, res.stderr or "解密失败")
            except Exception as e:
                fail_cnt += 1
                self.file_finished.emit(row_idx, False, str(e))
            finally:
                if staging_dir:
                    try:
                        recycle_owned_path(staging_dir, staging_dir)
                    except Exception:
                        self.progress.emit(f"临时文件未能移入回收站，已保留于: {staging_dir}")

        self.finished_all.emit(success_cnt, fail_cnt, self.output_dir)


class SecurityCenterDialog(FramelessDraggableDialog):
    """PDF 加密与解密工作台对话框（批量解密脱壳 + 加密保护）。"""

    def __init__(self, initial_file: Optional[str] = None, parent=None):
        super().__init__("PDF 加密与解密", parent)
        self.setWindowTitle("PDF 加密与解密")
        self.resize(840, 680)
        self.setMinimumSize(760, 580)
        self.encrypt_worker: Optional[SecurityWorker] = None
        self.decrypt_worker: Optional[DecryptWorker] = None
        self.last_output_path = ""
        self.last_decrypt_dir = ""

        self.content_layout.setContentsMargins(20, 14, 20, 16)
        self.content_layout.setSpacing(10)

        self.tabs = QTabWidget()
        self.tabs.setObjectName("securityTabs")
        self.tabs.tabBar().setDrawBase(False)
        self.tabs.tabBar().setFocusPolicy(Qt.NoFocus)
        self._build_decrypt_tab()
        self._build_encrypt_tab(initial_file)
        self.content_layout.addWidget(self.tabs, stretch=1)

        self._build_footer()

        if initial_file and os.path.exists(initial_file):
            self.add_decrypt_files([initial_file])

    def _attach_password_toggle(self, edit: QLineEdit):
        pal = active_palette()
        action = edit.addAction(
            icons.svg_icon("eye", pal.text_muted, size=14, disabled_color=pal.text_faint),
            QLineEdit.TrailingPosition,
        )
        action.setToolTip("显示/隐藏密码")

        def _toggle():
            is_hidden = edit.echoMode() == QLineEdit.Password
            edit.setEchoMode(QLineEdit.Normal if is_hidden else QLineEdit.Password)
            cur_pal = active_palette()
            action.setIcon(
                icons.svg_icon(
                    "eye_off" if is_hidden else "eye",
                    cur_pal.text_muted,
                    size=14,
                    disabled_color=cur_pal.text_faint,
                )
            )

        action.triggered.connect(_toggle)

    @staticmethod
    def _make_form_label(text: str, min_width: int = 72) -> QLabel:
        lbl = QLabel(text)
        lbl.setObjectName("dialogCaption")
        lbl.setMinimumWidth(min_width)
        return lbl

    # =========================================================================
    # 选项卡 1：PDF 加密与权限保护
    # =========================================================================
    def _build_encrypt_tab(self, initial_file: Optional[str]):
        pal = active_palette()
        tab = QWidget()
        tab_layout = QVBoxLayout(tab)
        tab_layout.setContentsMargins(0, 0, 0, 0)
        tab_layout.setSpacing(0)

        scroll_area = QScrollArea()
        scroll_area.setObjectName("settingsScroll")
        scroll_area.setWidgetResizable(True)
        scroll_area.setFrameShape(QFrame.NoFrame)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)

        scroll_body = QWidget()
        layout = QVBoxLayout(scroll_body)
        layout.setContentsMargins(1, 2, 1, 4)
        layout.setSpacing(12)

        # 1. 文件位置
        picker_card = QFrame()
        picker_card.setObjectName("securitySectionCard")
        grid = QGridLayout(picker_card)
        grid.setContentsMargins(14, 12, 14, 12)
        grid.setHorizontalSpacing(10)
        grid.setVerticalSpacing(10)

        grid.addWidget(self._make_form_label("源 PDF", 74), 0, 0)
        self.txt_input = QLineEdit()
        self.txt_input.setObjectName("settingsPathEdit")
        self.txt_input.setFixedHeight(34)
        self.txt_input.setPlaceholderText("选择需要加密保护的源 PDF 文件…")
        self.txt_input.textChanged.connect(self._auto_fill_output)
        self.txt_input.textChanged.connect(self._update_encrypt_status)
        self.btn_browse_input = QPushButton("浏览…")
        self.btn_browse_input.setObjectName("dialogSecondaryBtn")
        self.btn_browse_input.setFixedHeight(34)
        self.btn_browse_input.setCursor(Qt.PointingHandCursor)
        self.btn_browse_input.setIconSize(QSize(14, 14))
        self.btn_browse_input.setIcon(
            icons.svg_icon("folder_open", pal.text_body, size=14, disabled_color=pal.text_faint)
        )
        self.btn_browse_input.clicked.connect(self._on_browse_input_clicked)
        grid.addWidget(self.txt_input, 0, 1)
        grid.addWidget(self.btn_browse_input, 0, 2)

        grid.addWidget(self._make_form_label("输出路径", 74), 1, 0)
        self.txt_output = QLineEdit()
        self.txt_output.setObjectName("settingsPathEdit")
        self.txt_output.setFixedHeight(34)
        self.txt_output.setPlaceholderText("加密后 PDF 文件的保存路径…")
        self.btn_browse_output = QPushButton("浏览…")
        self.btn_browse_output.setObjectName("dialogSecondaryBtn")
        self.btn_browse_output.setFixedHeight(34)
        self.btn_browse_output.setCursor(Qt.PointingHandCursor)
        self.btn_browse_output.setIconSize(QSize(14, 14))
        self.btn_browse_output.setIcon(
            icons.svg_icon("folder_open", pal.text_body, size=14, disabled_color=pal.text_faint)
        )
        self.btn_browse_output.clicked.connect(self._on_browse_output_clicked)
        grid.addWidget(self.txt_output, 1, 1)
        grid.addWidget(self.btn_browse_output, 1, 2)
        grid.setColumnStretch(1, 1)
        layout.addWidget(picker_card)

        # 2. 算法与密码
        crypto_card = QFrame()
        crypto_card.setObjectName("securitySectionCard")
        grid_crypto = QGridLayout(crypto_card)
        grid_crypto.setContentsMargins(14, 12, 14, 12)
        grid_crypto.setHorizontalSpacing(10)
        grid_crypto.setVerticalSpacing(10)

        grid_crypto.addWidget(self._make_form_label("加密算法", 74), 0, 0)
        self.combo_algo = QComboBox()
        self.combo_algo.setFixedHeight(34)
        self.combo_algo.addItem("256-bit AES（推荐，兼容 Acrobat X 及以上）", 256)
        self.combo_algo.addItem("128-bit AES（兼容旧版阅读器）", 128)
        grid_crypto.addWidget(self.combo_algo, 0, 1)

        self.cb_user_pass = QCheckBox("打开密码")
        self.cb_user_pass.setToolTip("设置打开文档时需输入的密码（User Password）")
        self.txt_user_pass = QLineEdit()
        self.txt_user_pass.setObjectName("settingsPathEdit")
        self.txt_user_pass.setFixedHeight(34)
        self.txt_user_pass.setEchoMode(QLineEdit.Password)
        self.txt_user_pass.setPlaceholderText("留空表示任何人均可直接打开阅读")
        self.txt_user_pass.setEnabled(False)
        self._attach_password_toggle(self.txt_user_pass)
        self.cb_user_pass.toggled.connect(self.txt_user_pass.setEnabled)
        grid_crypto.addWidget(self.cb_user_pass, 1, 0)
        grid_crypto.addWidget(self.txt_user_pass, 1, 1)

        self.cb_owner_pass = QCheckBox("权限密码")
        self.cb_owner_pass.setToolTip("设置修改打印、复制或编辑权限时需输入的密码（Owner Password）")
        self.cb_owner_pass.setChecked(True)
        self.txt_owner_pass = QLineEdit()
        self.txt_owner_pass.setObjectName("settingsPathEdit")
        self.txt_owner_pass.setFixedHeight(34)
        self.txt_owner_pass.setEchoMode(QLineEdit.Password)
        self.txt_owner_pass.setPlaceholderText("用于保护打印、编辑与复制限制策略")
        self._attach_password_toggle(self.txt_owner_pass)
        self.cb_owner_pass.toggled.connect(self.txt_owner_pass.setEnabled)
        grid_crypto.addWidget(self.cb_owner_pass, 2, 0)
        grid_crypto.addWidget(self.txt_owner_pass, 2, 1)
        grid_crypto.setColumnStretch(1, 1)
        layout.addWidget(crypto_card)

        # 3. 权限细则与输出选项
        perm_card = QFrame()
        perm_card.setObjectName("securitySectionCard")
        vbox_perm = QVBoxLayout(perm_card)
        vbox_perm.setContentsMargins(14, 12, 14, 12)
        vbox_perm.setSpacing(10)

        lbl_perm = QLabel("权限控制与输出选项")
        lbl_perm.setObjectName("dialogSubTitle")
        vbox_perm.addWidget(lbl_perm)

        grid_perm = QGridLayout()
        grid_perm.setContentsMargins(0, 0, 0, 0)
        grid_perm.setHorizontalSpacing(10)
        grid_perm.setVerticalSpacing(10)

        grid_perm.addWidget(self._make_form_label("打印权限", 74), 0, 0)
        h_print = QHBoxLayout()
        h_print.setContentsMargins(0, 0, 0, 0)
        h_print.setSpacing(14)
        self.print_group = QButtonGroup(self)
        self.rb_print_full = QRadioButton("高质量打印")
        self.rb_print_low = QRadioButton("低分辨率 (150 dpi)")
        self.rb_print_none = QRadioButton("禁止打印")
        self.rb_print_full.setChecked(True)
        self.print_group.addButton(self.rb_print_full, 0)
        self.print_group.addButton(self.rb_print_low, 1)
        self.print_group.addButton(self.rb_print_none, 2)
        h_print.addWidget(self.rb_print_full)
        h_print.addWidget(self.rb_print_low)
        h_print.addWidget(self.rb_print_none)
        h_print.addStretch()
        grid_perm.addLayout(h_print, 0, 1)

        grid_perm.addWidget(self._make_form_label("修改权限", 74), 1, 0)
        self.combo_modify = QComboBox()
        self.combo_modify.setFixedHeight(34)
        self.combo_modify.addItem("完全禁止修改（推荐定稿归档）", "none")
        self.combo_modify.addItem("仅允许页面装配（插入、删除、旋转页面）", "assembly")
        self.combo_modify.addItem("仅允许填写表单字段与电子签名", "form")
        self.combo_modify.addItem("允许批注与填写表单", "annotate")
        self.combo_modify.addItem("允许完全修改", "all")
        grid_perm.addWidget(self.combo_modify, 1, 1)
        grid_perm.setColumnStretch(1, 1)
        vbox_perm.addLayout(grid_perm)

        checks_grid = QGridLayout()
        checks_grid.setContentsMargins(0, 2, 0, 0)
        checks_grid.setHorizontalSpacing(20)
        checks_grid.setVerticalSpacing(8)

        self.cb_extract = QCheckBox("允许复制与提取内容")
        self.cb_extract.setToolTip("允许未持有权限密码的用户复制文本或提取图像 (Extract)")
        self.cb_accessibility = QCheckBox("允许无障碍辅助阅读")
        self.cb_accessibility.setToolTip("允许屏幕阅读器等无障碍工具访问文档内容 (Accessibility)")
        self.cb_accessibility.setChecked(True)
        self.cb_cleartext_meta = QCheckBox("保持元数据明文（便于检索）")
        self.cb_cleartext_meta.setToolTip("保持文档元数据明文不加密 (Cleartext Metadata)，便于 Windows 检索与档案库索引")
        self.cb_cleartext_meta.setChecked(True)
        self.cb_linearize = QCheckBox("启用 Web 快速视图（线性化）")
        self.cb_linearize.setChecked(True)

        checks_grid.addWidget(self.cb_extract, 0, 0)
        checks_grid.addWidget(self.cb_accessibility, 0, 1)
        checks_grid.addWidget(self.cb_cleartext_meta, 1, 0)
        checks_grid.addWidget(self.cb_linearize, 1, 1)
        checks_grid.setColumnStretch(0, 1)
        checks_grid.setColumnStretch(1, 1)
        vbox_perm.addLayout(checks_grid)
        layout.addWidget(perm_card)

        layout.addStretch()
        scroll_area.setWidget(scroll_body)
        tab_layout.addWidget(scroll_area)

        self.tabs.addTab(tab, "加密与权限保护")

        if initial_file and os.path.exists(initial_file):
            self.txt_input.setText(initial_file)

    # =========================================================================
    # 选项卡 2：PDF 解密与权限脱壳
    # =========================================================================
    def _build_decrypt_tab(self):
        pal = active_palette()
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(1, 1, 1, 1)
        layout.setSpacing(10)

        # 1. 待解密文件列表卡片（单层边框 + 顶栏操作 + 无网格表格）
        self.table_card = QFrame()
        self.table_card.setObjectName("securityTableCard")
        self.table_card.setAcceptDrops(True)
        self.table_card.dragEnterEvent = self._decrypt_drag_enter
        self.table_card.dragMoveEvent = self._decrypt_drag_move
        self.table_card.dragLeaveEvent = self._decrypt_drag_leave
        self.table_card.dropEvent = self._decrypt_drop

        table_card_layout = QVBoxLayout(self.table_card)
        table_card_layout.setContentsMargins(1, 1, 1, 1)
        table_card_layout.setSpacing(0)

        table_header = QFrame()
        table_header.setObjectName("securityTableHeader")
        th_layout = QHBoxLayout(table_header)
        th_layout.setContentsMargins(12, 8, 10, 8)
        th_layout.setSpacing(6)

        self.lbl_dec_list_title = QLabel("待解密文件 (0)")
        self.lbl_dec_list_title.setObjectName("dialogSubTitle")
        list_hint = QLabel("支持拖拽 PDF 或文件夹")
        list_hint.setObjectName("dialogMuted")
        th_layout.addWidget(self.lbl_dec_list_title)
        th_layout.addWidget(list_hint)
        th_layout.addStretch()

        self.btn_dec_add = QPushButton("添加文件")
        self.btn_dec_add.setObjectName("compactSecondaryBtn")
        self.btn_dec_add.setCursor(Qt.PointingHandCursor)
        self.btn_dec_add.setIconSize(QSize(13, 13))
        self.btn_dec_add.setIcon(
            icons.svg_icon("plus", pal.text_body, size=13, disabled_color=pal.text_faint)
        )
        self.btn_dec_add.clicked.connect(self._on_dec_add_clicked)

        self.btn_dec_add_folder = QPushButton("添加文件夹")
        self.btn_dec_add_folder.setObjectName("compactSecondaryBtn")
        self.btn_dec_add_folder.setCursor(Qt.PointingHandCursor)
        self.btn_dec_add_folder.setIconSize(QSize(13, 13))
        self.btn_dec_add_folder.setIcon(
            icons.svg_icon("folder", pal.text_body, size=13, disabled_color=pal.text_faint)
        )
        self.btn_dec_add_folder.clicked.connect(self._on_dec_add_folder_clicked)

        self.btn_dec_remove = QPushButton("移除")
        self.btn_dec_remove.setObjectName("ghostBtn")
        self.btn_dec_remove.setCursor(Qt.PointingHandCursor)
        self.btn_dec_remove.setToolTip("移除表格中选中的文件")
        self.btn_dec_remove.clicked.connect(self._on_dec_remove_clicked)

        self.btn_dec_clear = QPushButton("清空")
        self.btn_dec_clear.setObjectName("ghostBtn")
        self.btn_dec_clear.setCursor(Qt.PointingHandCursor)
        self.btn_dec_clear.setIconSize(QSize(13, 13))
        self.btn_dec_clear.setIcon(
            icons.svg_icon("trash", pal.text_muted, size=13, disabled_color=pal.text_faint)
        )
        self.btn_dec_clear.clicked.connect(self._on_dec_clear_clicked)

        th_layout.addWidget(self.btn_dec_add)
        th_layout.addWidget(self.btn_dec_add_folder)
        th_layout.addWidget(self.btn_dec_remove)
        th_layout.addWidget(self.btn_dec_clear)
        table_card_layout.addWidget(table_header)

        self.decrypt_table = QTableWidget()
        self.decrypt_table.setObjectName("securityTable")
        self.decrypt_table.setFrameShape(QFrame.NoFrame)
        self.decrypt_table.setShowGrid(False)
        self.decrypt_table.verticalHeader().setVisible(False)
        self.decrypt_table.verticalHeader().setDefaultSectionSize(34)
        self.decrypt_table.setItemDelegate(NoFocusItemDelegate(self.decrypt_table))
        self.decrypt_table.setColumnCount(6)
        self.decrypt_table.setHorizontalHeaderLabels(
            ["#", "文件名", "页数", "安全状态", "诊断建议", "完整路径"]
        )
        header = self.decrypt_table.horizontalHeader()
        header.setDefaultAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        header.setSectionResizeMode(0, QHeaderView.Fixed)
        header.setSectionResizeMode(1, QHeaderView.Interactive)
        header.setSectionResizeMode(2, QHeaderView.Fixed)
        header.setSectionResizeMode(3, QHeaderView.Fixed)
        header.setSectionResizeMode(4, QHeaderView.Stretch)
        self.decrypt_table.setColumnWidth(0, 38)
        self.decrypt_table.setColumnWidth(1, 220)
        self.decrypt_table.setColumnWidth(2, 56)
        self.decrypt_table.setColumnWidth(3, 112)
        self.decrypt_table.setColumnHidden(5, True)
        self.decrypt_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.decrypt_table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.decrypt_table.setAlternatingRowColors(True)

        self.decrypt_table.setAcceptDrops(True)
        self.decrypt_table.dragEnterEvent = self._decrypt_drag_enter
        self.decrypt_table.dragMoveEvent = self._decrypt_drag_move
        self.decrypt_table.dragLeaveEvent = self._decrypt_drag_leave
        self.decrypt_table.dropEvent = self._decrypt_drop
        self.decrypt_table.itemSelectionChanged.connect(self._sync_decrypt_table_controls)

        table_card_layout.addWidget(self.decrypt_table, stretch=1)
        layout.addWidget(self.table_card, stretch=1)

        # 2. 解密与解锁选项
        pass_card = QFrame()
        pass_card.setObjectName("securitySectionCard")
        v_pass = QVBoxLayout(pass_card)
        v_pass.setContentsMargins(14, 10, 14, 10)
        v_pass.setSpacing(8)

        row_pw = QHBoxLayout()
        row_pw.setSpacing(10)
        self.cb_dec_use_pass = QCheckBox("输入打开/权限密码")
        self.txt_dec_pass = QLineEdit()
        self.txt_dec_pass.setObjectName("settingsPathEdit")
        self.txt_dec_pass.setFixedHeight(32)
        self.txt_dec_pass.setEchoMode(QLineEdit.Password)
        self.txt_dec_pass.setPlaceholderText("设有打开密码时勾选并输入，仅受权限限制可留空")
        self.txt_dec_pass.setEnabled(False)
        self._attach_password_toggle(self.txt_dec_pass)
        self.cb_dec_use_pass.toggled.connect(self.txt_dec_pass.setEnabled)
        row_pw.addWidget(self.cb_dec_use_pass)
        row_pw.addWidget(self.txt_dec_pass, stretch=1)
        v_pass.addLayout(row_pw)

        hint_lbl = QLabel(
            "仅受权限限制（禁止打印/编辑/复制）的 PDF 无需密码即可免密解除；设有打开密码的文件请勾选并输入密码。"
        )
        hint_lbl.setObjectName("dialogMuted")
        hint_lbl.setWordWrap(True)
        hint_lbl.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Minimum)
        v_pass.addWidget(hint_lbl)

        row_sig_opts = QGridLayout()
        row_sig_opts.setContentsMargins(0, 2, 0, 0)
        row_sig_opts.setHorizontalSpacing(20)
        self.cb_dec_remove_restrictions = QCheckBox("解除数字签名限制（保留外观）")
        self.cb_dec_remove_restrictions.setToolTip(
            "解除数字签名限制并解锁后续编辑；采用 300 DPI 靶向外观固化技术，绝对保留印章与签字视觉效果且不破坏任何超链接"
        )
        self.cb_dec_remove_restrictions.setChecked(True)

        self.cb_dec_remove_acroform = QCheckBox("移除交互表单字典（AcroForm）")
        self.cb_dec_remove_acroform.setToolTip(
            "从文档 Catalog 中移除交互式表单字典 (--remove-acroform)，清除表单字段与签名锁定状态"
        )
        self.cb_dec_remove_acroform.setChecked(False)

        row_sig_opts.addWidget(self.cb_dec_remove_restrictions, 0, 0)
        row_sig_opts.addWidget(self.cb_dec_remove_acroform, 0, 1)
        row_sig_opts.setColumnStretch(0, 1)
        row_sig_opts.setColumnStretch(1, 1)
        v_pass.addLayout(row_sig_opts)
        layout.addWidget(pass_card)

        # 3. 输出设置
        out_card = QFrame()
        out_card.setObjectName("securitySectionCard")
        grid_out = QGridLayout(out_card)
        grid_out.setContentsMargins(14, 10, 14, 10)
        grid_out.setHorizontalSpacing(10)
        grid_out.setVerticalSpacing(8)

        grid_out.addWidget(self._make_form_label("输出目录", 72), 0, 0)
        self.txt_dec_out_dir = QLineEdit()
        self.txt_dec_out_dir.setObjectName("settingsPathEdit")
        self.txt_dec_out_dir.setFixedHeight(32)
        self.txt_dec_out_dir.setPlaceholderText("选择解密后 PDF 的保存目录…")
        self.btn_dec_browse_dir = QPushButton("浏览…")
        self.btn_dec_browse_dir.setObjectName("dialogSecondaryBtn")
        self.btn_dec_browse_dir.setFixedHeight(32)
        self.btn_dec_browse_dir.setCursor(Qt.PointingHandCursor)
        self.btn_dec_browse_dir.setIconSize(QSize(14, 14))
        self.btn_dec_browse_dir.setIcon(
            icons.svg_icon("folder_open", pal.text_body, size=14, disabled_color=pal.text_faint)
        )
        self.btn_dec_browse_dir.clicked.connect(self._on_dec_browse_dir_clicked)
        grid_out.addWidget(self.txt_dec_out_dir, 0, 1)
        grid_out.addWidget(self.btn_dec_browse_dir, 0, 2)

        grid_out.addWidget(self._make_form_label("文件命名", 72), 1, 0)
        row_naming = QHBoxLayout()
        row_naming.setContentsMargins(0, 0, 0, 0)
        row_naming.setSpacing(16)
        self.naming_group = QButtonGroup(self)
        self.rb_name_suffix = QRadioButton("追加 _decrypted 后缀（保留原件）")
        self.rb_name_origin = QRadioButton("保持原文件名（覆盖输出）")
        self.rb_name_suffix.setChecked(True)
        self.naming_group.addButton(self.rb_name_suffix, 0)
        self.naming_group.addButton(self.rb_name_origin, 1)
        row_naming.addWidget(self.rb_name_suffix)
        row_naming.addWidget(self.rb_name_origin)
        row_naming.addStretch()
        grid_out.addLayout(row_naming, 1, 1, 1, 2)

        grid_out.addWidget(self._make_form_label("输出优化", 72), 2, 0)
        row_opts2 = QHBoxLayout()
        row_opts2.setContentsMargins(0, 0, 0, 0)
        row_opts2.setSpacing(16)
        self.cb_dec_linearize = QCheckBox("启用 Web 快速视图（线性化）")
        self.cb_dec_linearize.setChecked(True)
        self.cb_dec_objstms = QCheckBox("生成对象流压缩（PDF 1.5+）")
        self.cb_dec_objstms.setChecked(True)
        row_opts2.addWidget(self.cb_dec_linearize)
        row_opts2.addWidget(self.cb_dec_objstms)
        row_opts2.addStretch()
        grid_out.addLayout(row_opts2, 2, 1, 1, 2)
        grid_out.setColumnStretch(1, 1)
        layout.addWidget(out_card)

        self._sync_decrypt_table_controls()
        self.tabs.addTab(tab, "解密与解除限制")

    # =========================================================================
    # 底部状态栏与主执行按钮
    # =========================================================================
    def _build_footer(self):
        pal = active_palette()
        footer = QFrame()
        footer.setObjectName("securityFooter")
        layout = QHBoxLayout(footer)
        layout.setContentsMargins(0, 10, 0, 0)
        layout.setSpacing(10)

        self.lbl_status = ElidedLabel("就绪")
        self.lbl_status.setObjectName("dialogMuted")
        layout.addWidget(self.lbl_status, stretch=1)

        self.btn_open_target = QPushButton("打开输出目录")
        self.btn_open_target.setObjectName("dialogSecondaryBtn")
        self.btn_open_target.setFixedHeight(34)
        self.btn_open_target.setCursor(Qt.PointingHandCursor)
        self.btn_open_target.setIconSize(QSize(14, 14))
        self.btn_open_target.setIcon(
            icons.svg_icon("folder_open", pal.text_body, size=14, disabled_color=pal.text_faint)
        )
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
        self.btn_execute.setIconSize(QSize(14, 14))
        self.btn_execute.clicked.connect(self._on_execute_clicked)
        layout.addWidget(self.btn_execute)

        self.content_layout.addWidget(footer)
        self.tabs.currentChanged.connect(self._on_tab_changed)
        self._on_tab_changed(0)

    def _sync_decrypt_table_controls(self):
        cnt = self.decrypt_table.rowCount()
        if hasattr(self, "lbl_dec_list_title"):
            self.lbl_dec_list_title.setText(f"待解密文件 ({cnt})")
        if hasattr(self, "btn_dec_clear"):
            self.btn_dec_clear.setEnabled(cnt > 0)
        if hasattr(self, "btn_dec_remove"):
            self.btn_dec_remove.setEnabled(cnt > 0 and self.decrypt_table.currentRow() >= 0)

    def _on_tab_changed(self, idx: int):
        pal = active_palette()
        self.btn_open_target.setVisible(False)
        if idx == 0:
            self.btn_execute.setText("开始批量解密")
            self.btn_execute.setIcon(
                icons.svg_icon("unlock", pal.text_on_primary, size=14, disabled_color=pal.surface)
            )
            cnt = self.decrypt_table.rowCount()
            self.lbl_status.setText(f"已选 {cnt} 个待解密文件" if cnt > 0 else "请添加待解密 PDF 文件")
        else:
            self.btn_execute.setText("应用安全加密")
            self.btn_execute.setIcon(
                icons.svg_icon("lock", pal.text_on_primary, size=14, disabled_color=pal.surface)
            )
            self._update_encrypt_status()
        self.refresh_backdrop()
        QTimer.singleShot(0, self, self.refresh_backdrop)

    def _update_encrypt_status(self):
        if not hasattr(self, "lbl_status") or self.tabs.currentIndex() != 1:
            return
        in_path = self.txt_input.text().strip()
        if not in_path:
            self.lbl_status.setText("请选择需要加密保护的源 PDF 文件")
        elif not os.path.exists(in_path):
            self.lbl_status.setText("所选源文件不存在")
        else:
            self.lbl_status.setText(f"已选择待加密文件：{os.path.basename(in_path)}")

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
        self.lbl_status.setText("正在加密...")
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
            pal = active_palette()
            if status_code == "password_required":
                it_status.setForeground(QColor(pal.danger_text))
            elif status_code == "restricted_no_password":
                it_status.setForeground(QColor(pal.warning_text))
            else:
                it_status.setForeground(QColor(pal.success_text))
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
        self._sync_decrypt_table_controls()
        if hasattr(self, "lbl_status"):
            self.lbl_status.setText(f"已选 {cnt} 个待解密文件")

    def _on_dec_add_clicked(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "选择待解密 PDF 文件", "", "PDF 文件 (*.pdf)"
        )
        if files:
            self.add_decrypt_files(files)

    def _on_dec_add_folder_clicked(self):
        folder = QFileDialog.getExistingDirectory(self, "选择包含 PDF 文件的文件夹")
        if not folder:
            return
        matched = []
        for root, _, fnames in os.walk(folder):
            for f in fnames:
                if f.lower().endswith(".pdf"):
                    matched.append(os.path.join(root, f))
        if matched:
            matched.sort()
            self.add_decrypt_files(matched)
        else:
            QMessageBox.information(self, "提示", "所选文件夹中未找到任何 PDF 文件！")

    def _on_dec_remove_clicked(self):
        row = self.decrypt_table.currentRow()
        if row >= 0:
            self.decrypt_table.removeRow(row)
            for r in range(self.decrypt_table.rowCount()):
                it = self.decrypt_table.item(r, 0)
                if it:
                    it.setText(str(r + 1))
            cnt = self.decrypt_table.rowCount()
            self._sync_decrypt_table_controls()
            self.lbl_status.setText(f"已选 {cnt} 个待解密文件" if cnt > 0 else "请添加待解密 PDF 文件")

    def _on_dec_clear_clicked(self):
        self.decrypt_table.setRowCount(0)
        self._sync_decrypt_table_controls()
        self.lbl_status.setText("已清空列表")

    def _on_dec_browse_dir_clicked(self):
        folder = QFileDialog.getExistingDirectory(self, "选择解密输出文件夹")
        if folder:
            self.txt_dec_out_dir.setText(folder)

    def _set_table_drag_active(self, active: bool):
        card = getattr(self, "table_card", None)
        if card is not None:
            card.setProperty("dragActive", "true" if active else "false")
            card.style().unpolish(card)
            card.style().polish(card)

    def _decrypt_drag_enter(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            self._set_table_drag_active(True)

    def _decrypt_drag_move(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def _decrypt_drag_leave(self, _event):
        self._set_table_drag_active(False)

    def _decrypt_drop(self, event):
        self._set_table_drag_active(False)
        urls = event.mimeData().urls()
        files = []
        for u in urls:
            local = u.toLocalFile()
            if os.path.isfile(local) and local.lower().endswith(".pdf"):
                files.append(local)
            elif os.path.isdir(local):
                for root, _, fnames in os.walk(local):
                    for f in fnames:
                        if f.lower().endswith(".pdf"):
                            files.append(os.path.join(root, f))
        if files:
            files.sort()
            self.add_decrypt_files(files)
            event.acceptProposedAction()

    def _execute_decryption(self):
        if self.decrypt_worker is not None:
            return
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
        self.lbl_status.setText("正在批量解密脱壳...")

        self._decrypt_summary = None
        self._decrypt_rows = {row for row, _ in tasks}
        self._decrypt_output_dir = out_dir
        try:
            self.decrypt_worker = DecryptWorker(
                tasks=tasks,
                output_dir=out_dir,
                password=password,
                suffix=suffix,
                remove_restrictions=self.cb_dec_remove_restrictions.isChecked(),
                remove_acroform=self.cb_dec_remove_acroform.isChecked(),
                linearize=self.cb_dec_linearize.isChecked(),
                object_streams=self.cb_dec_objstms.isChecked(),
                parent=self,
            )
            self.decrypt_worker.progress.connect(lambda msg: self.lbl_status.setText(msg))
            self.decrypt_worker.file_finished.connect(self._on_dec_file_finished)
            self.decrypt_worker.finished_all.connect(self._on_dec_finished_all)
            self.decrypt_worker.finished.connect(self._on_dec_thread_finished)
            self.decrypt_worker.start()
        except Exception as exc:
            for row, _ in tasks:
                self._on_dec_file_finished(row, False, f"解密启动失败: {exc}")
            self._decrypt_summary = (0, len(tasks), out_dir)
            self._finish_decryption_thread()

    def _on_dec_file_finished(self, row_idx: int, success: bool, msg: str):
        it_status = self.decrypt_table.item(row_idx, 3)
        it_desc = self.decrypt_table.item(row_idx, 4)
        pal = active_palette()
        if it_status:
            it_status.setText("已解密" if success else "失败")
            it_status.setForeground(QColor(pal.success_text if success else pal.danger_text))
        if it_desc:
            it_desc.setText(msg)

    def _on_dec_finished_all(self, success_cnt: int, fail_cnt: int, out_dir: str):
        self._decrypt_summary = (success_cnt, fail_cnt, out_dir)

    def _on_dec_thread_finished(self):
        if self.sender() is self.decrypt_worker:
            self._finish_decryption_thread()

    def _finish_decryption_thread(self):
        worker = self.decrypt_worker
        self.decrypt_worker = None
        if worker is not None:
            worker.deleteLater()
        summary = self._decrypt_summary
        if summary is None:
            for row in self._decrypt_rows:
                self._on_dec_file_finished(row, False, "解密线程未返回完成结果，请重试")
            summary = (0, len(self._decrypt_rows), self._decrypt_output_dir)
        success_cnt, fail_cnt, out_dir = summary
        self.btn_execute.setEnabled(True)
        self.btn_cancel.setEnabled(True)
        self.lbl_status.setText(f"解密完成：成功 {success_cnt} 个，失败 {fail_cnt} 个")
        self.last_decrypt_dir = out_dir if success_cnt else ""
        self.btn_open_target.setVisible(success_cnt > 0)

        if fail_cnt == 0:
            QMessageBox.information(
                self, "完成", f"✅ 全部 {success_cnt} 个文档已成功解密脱壳！\n保存至：{out_dir}"
            )
        else:
            QMessageBox.warning(
                self,
                "部分完成" if success_cnt else "解密失败",
                f"解密处理结束：成功 {success_cnt} 个，失败 {fail_cnt} 个。\n请查看列表中的失败原因，修正后重试。",
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
