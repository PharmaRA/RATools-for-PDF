"""图层套印与印章管理器：支持公章/骑缝章/水印前景覆盖与信头纸/表格底纹套打。"""

import os
from typing import Optional

from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtWidgets import (
    QButtonGroup,
    QCheckBox,
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


class OverlayWorker(QThread):
    """后台执行图层叠加或打底合成的异步工作线程。"""

    progress = Signal(str)
    finished_signal = Signal(bool, str, str)  # success, message, target_path

    def __init__(
        self,
        input_pdf: str,
        output_pdf: str,
        layer_pdf: str,
        is_overlay: bool,
        repeat: Optional[str] = "1-z",
        to_range: Optional[str] = None,
        linearize: bool = True,
        object_streams: bool = True,
        parent=None,
    ):
        super().__init__(parent)
        self.input_pdf = input_pdf
        self.output_pdf = output_pdf
        self.layer_pdf = layer_pdf
        self.is_overlay = is_overlay
        self.repeat = repeat
        self.to_range = to_range
        self.linearize = linearize
        self.object_streams = "generate" if object_streams else None

    def run(self):
        try:
            self.progress.emit("正在合成图层...")
            overlay_file = self.layer_pdf if self.is_overlay else None
            underlay_file = self.layer_pdf if not self.is_overlay else None

            res = qpdf.apply_overlay_underlay(
                input_pdf=self.input_pdf,
                output_pdf=self.output_pdf,
                overlay_file=overlay_file,
                underlay_file=underlay_file,
                to_range=self.to_range,
                repeat=self.repeat,
                linearize=self.linearize,
                object_streams=self.object_streams,
            )
            if res.is_success:
                mode_desc = "前景叠加 (Overlay)" if self.is_overlay else "背景打底 (Underlay)"
                self.finished_signal.emit(
                    True,
                    f"✅ 成功完成{mode_desc}合成：\n{self.output_pdf}",
                    self.output_pdf,
                )
            else:
                self.finished_signal.emit(
                    False, f"❌ 合成失败：{res.stderr or '未知错误'}", ""
                )
        except Exception as e:
            self.finished_signal.emit(False, f"❌ 执行异常：{str(e)}", "")


class OverlayUnderlayDialog(FramelessDraggableDialog):
    """图层套印与印章管理器对话框。"""

    def __init__(self, initial_file: Optional[str] = None, parent=None):
        super().__init__("🎨 图层套印与印章管理器", parent)
        self.resize(720, 520)
        self.worker: Optional[OverlayWorker] = None
        self.last_output_path = ""

        self.content_layout.setSpacing(12)

        self._build_mode_section()
        self._build_file_section(None)
        self._build_range_section()
        self._build_output_section()
        self._build_footer()

        if initial_file and os.path.exists(initial_file):
            self.txt_target.setText(initial_file)

    def _build_mode_section(self):
        """图层模式选择：前景 Overlay vs 背景 Underlay。"""
        card = QFrame()
        card.setObjectName("wizardCard")
        layout = QHBoxLayout(card)
        layout.setContentsMargins(14, 10, 14, 10)
        layout.setSpacing(24)

        lbl = QLabel("合成模式：")
        lbl.setStyleSheet("font-weight: 700;")
        layout.addWidget(lbl)

        self.mode_group = QButtonGroup(self)
        self.rb_overlay = QRadioButton("前景覆印 (Overlay)：公章、骑缝章、草稿/机密水印")
        self.rb_underlay = QRadioButton("背景打底 (Underlay)：信头纸套打、背景网格、空白表单模板")
        self.rb_overlay.setChecked(True)

        self.mode_group.addButton(self.rb_overlay, 0)
        self.mode_group.addButton(self.rb_underlay, 1)

        layout.addWidget(self.rb_overlay)
        layout.addWidget(self.rb_underlay)
        layout.addStretch()

        self.content_layout.addWidget(card)

    def _build_file_section(self, initial_file: Optional[str]):
        """目标主文档与图层模板文档选择。"""
        card = QFrame()
        card.setObjectName("wizardCard")
        grid = QGridLayout(card)
        grid.setContentsMargins(14, 12, 14, 12)
        grid.setSpacing(10)

        grid.addWidget(QLabel("目标主文档："), 0, 0)
        self.txt_target = QLineEdit()
        self.txt_target.setPlaceholderText("选择需要被盖章或加水印的 PDF 主文档...")
        self.txt_target.textChanged.connect(self._auto_fill_output)
        self.btn_browse_target = QPushButton("浏览...")
        self.btn_browse_target.clicked.connect(self._on_browse_target_clicked)
        grid.addWidget(self.txt_target, 0, 1)
        grid.addWidget(self.btn_browse_target, 0, 2)

        grid.addWidget(QLabel("印章/图层源："), 1, 0)
        self.txt_layer = QLineEdit()
        self.txt_layer.setPlaceholderText("选择包含印章、水印或信头纸模板的 PDF 文件...")
        self.btn_browse_layer = QPushButton("浏览...")
        self.btn_browse_layer.clicked.connect(self._on_browse_layer_clicked)
        grid.addWidget(self.txt_layer, 1, 1)
        grid.addWidget(self.btn_browse_layer, 1, 2)

        self.content_layout.addWidget(card)

        if initial_file and os.path.exists(initial_file):
            self.txt_target.setText(initial_file)

    def _build_range_section(self):
        """套印页面范围与重复策略。"""
        card = QFrame()
        card.setObjectName("wizardCard")
        vbox = QVBoxLayout(card)
        vbox.setContentsMargins(14, 10, 14, 10)
        vbox.setSpacing(8)

        lbl = QLabel("套印范围与重复规则：")
        lbl.setStyleSheet("font-weight: 700;")
        vbox.addWidget(lbl)

        row = QHBoxLayout()
        self.repeat_group = QButtonGroup(self)
        self.rb_repeat_all = QRadioButton("循环覆盖全部页面 (第 1 页印章重复应用于目标所有页)")
        self.rb_repeat_custom = QRadioButton("指定目标页面范围：")
        self.rb_repeat_all.setChecked(True)

        self.repeat_group.addButton(self.rb_repeat_all, 0)
        self.repeat_group.addButton(self.rb_repeat_custom, 1)

        self.txt_to_range = QLineEdit()
        self.txt_to_range.setPlaceholderText("例如: 1-5, 8, z")
        self.txt_to_range.setEnabled(False)
        self.rb_repeat_custom.toggled.connect(self.txt_to_range.setEnabled)

        row.addWidget(self.rb_repeat_all)
        row.addWidget(self.rb_repeat_custom)
        row.addWidget(self.txt_to_range)
        row.addStretch()
        vbox.addLayout(row)

        self.content_layout.addWidget(card)

    def _build_output_section(self):
        """输出设置。"""
        card = QFrame()
        card.setObjectName("wizardCard")
        grid = QGridLayout(card)
        grid.setContentsMargins(14, 10, 14, 10)
        grid.setSpacing(10)

        grid.addWidget(QLabel("合成输出另存为："), 0, 0)
        self.txt_output = QLineEdit()
        self.txt_output.setPlaceholderText("输出的 PDF 文件完整路径...")
        self.btn_browse_output = QPushButton("浏览...")
        self.btn_browse_output.clicked.connect(self._on_browse_output_clicked)
        grid.addWidget(self.txt_output, 0, 1)
        grid.addWidget(self.btn_browse_output, 0, 2)

        # 选项
        h_opts = QHBoxLayout()
        self.cb_linearize = QCheckBox("启用 Web 快速视图 (线性化)")
        self.cb_linearize.setChecked(True)
        self.cb_object_streams = QCheckBox("压缩生成对象流")
        self.cb_object_streams.setChecked(True)
        h_opts.addWidget(self.cb_linearize)
        h_opts.addWidget(self.cb_object_streams)
        h_opts.addStretch()
        grid.addLayout(h_opts, 1, 1, 1, 2)

        self.content_layout.addWidget(card)

    def _build_footer(self):
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

        self.btn_execute = QPushButton("开始套印合成")
        self.btn_execute.setObjectName("choiceToggleBtn")
        self.btn_execute.setStyleSheet(
            "QPushButton { background-color: #2563EB; color: white; font-weight: bold; border-radius: 8px; padding: 8px 18px; }"
            "QPushButton:hover { background-color: #1D4ED8; }"
            "QPushButton:disabled { background-color: #93C5FD; }"
        )
        self.btn_execute.clicked.connect(self._on_execute_clicked)
        layout.addWidget(self.btn_execute)

        self.content_layout.addWidget(footer)

    def _auto_fill_output(self, target_path: str):
        if hasattr(self, "txt_output") and target_path and target_path.lower().endswith(".pdf") and not self.txt_output.text():
            dir_name = os.path.dirname(target_path)
            base_name = os.path.splitext(os.path.basename(target_path))[0]
            self.txt_output.setText(os.path.join(dir_name, f"{base_name}_stamped.pdf"))

    def _on_browse_target_clicked(self):
        f, _ = QFileDialog.getOpenFileName(self, "选择目标主 PDF 文档", "", "PDF 文件 (*.pdf)")
        if f:
            self.txt_target.setText(f)

    def _on_browse_layer_clicked(self):
        f, _ = QFileDialog.getOpenFileName(self, "选择图层/印章 PDF 模板", "", "PDF 文件 (*.pdf)")
        if f:
            self.txt_layer.setText(f)

    def _on_browse_output_clicked(self):
        f, _ = QFileDialog.getSaveFileName(
            self, "保存合成后 PDF 文件", self.txt_output.text() or "stamped.pdf", "PDF 文件 (*.pdf)"
        )
        if f:
            self.txt_output.setText(f)

    def _on_execute_clicked(self):
        target = self.txt_target.text().strip()
        layer = self.txt_layer.text().strip()
        out = self.txt_output.text().strip()

        if not target or not os.path.exists(target):
            QMessageBox.warning(self, "提示", "请选择有效的目标主 PDF 文档！")
            return
        if not layer or not os.path.exists(layer):
            QMessageBox.warning(self, "提示", "请选择有效的图层/印章 PDF 模板！")
            return
        if not out:
            QMessageBox.warning(self, "提示", "请指定输出另存文件路径！")
            return

        is_overlay = self.rb_overlay.isChecked()
        is_repeat_all = self.rb_repeat_all.isChecked()
        to_range = self.txt_to_range.text().strip() if not is_repeat_all else None
        repeat = "1-z" if is_repeat_all else None

        self.btn_execute.setEnabled(False)
        self.btn_cancel.setEnabled(False)
        self.lbl_status.setText("⏳ 正在合成图层...")
        self.btn_open_target.setVisible(False)

        self.worker = OverlayWorker(
            input_pdf=target,
            output_pdf=out,
            layer_pdf=layer,
            is_overlay=is_overlay,
            repeat=repeat,
            to_range=to_range,
            linearize=self.cb_linearize.isChecked(),
            object_streams=self.cb_object_streams.isChecked(),
            parent=self,
        )
        self.worker.progress.connect(lambda msg: self.lbl_status.setText(msg))
        self.worker.finished_signal.connect(self._on_worker_finished)
        self.worker.start()

    def _on_worker_finished(self, success: bool, message: str, target_path: str):
        self.btn_execute.setEnabled(True)
        self.btn_cancel.setEnabled(True)
        self.lbl_status.setText("合成完成" if success else "合成失败")

        if success:
            self.last_output_path = target_path
            self.btn_open_target.setVisible(True)
            QMessageBox.information(self, "完成", message)
        else:
            QMessageBox.critical(self, "错误", message)

    def _on_open_target_clicked(self):
        if self.last_output_path and os.path.exists(self.last_output_path):
            os.startfile(os.path.dirname(self.last_output_path))
