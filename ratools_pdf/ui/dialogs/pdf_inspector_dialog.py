"""PDF 深度诊断与结构探针：语法体检(--check)、加密透视、线性化分析、JSON 语法树探针与受损 XRef 灾难修复。"""

import json
import os
from typing import Any, Dict, Optional

from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtGui import QFont
from PySide6.QtWidgets import (
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QTabWidget,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

from ratools_pdf.pdf import qpdf
from ratools_pdf.ui.dialogs.base import FramelessDraggableDialog


class DiagnosticWorker(QThread):
    """后台运行各维度 qpdf 诊断任务的异步工作线程。"""

    finished_signal = Signal(dict)  # results payload

    def __init__(self, pdf_path: str, parent=None):
        super().__init__(parent)
        self.pdf_path = pdf_path

    def run(self):
        payload: Dict[str, Any] = {
            "path": self.pdf_path,
            "check": {},
            "encryption": "",
            "linearization": "",
            "json_ast": None,
            "header_version": "",
            "page_count": 0,
        }
        try:
            payload["header_version"] = qpdf.read_pdf_header_version(self.pdf_path)
            payload["page_count"] = qpdf.get_pdf_page_count(self.pdf_path)
            payload["check"] = qpdf.check_pdf_syntax(self.pdf_path)
            payload["encryption"] = qpdf.qpdf_encryption_info(self.pdf_path)
            payload["linearization"] = qpdf.get_linearization_report(self.pdf_path)
            payload["json_ast"] = qpdf.inspect_pdf_json(self.pdf_path)
        except Exception as e:
            payload["error"] = str(e)
        self.finished_signal.emit(payload)


class PdfInspectorDialog(FramelessDraggableDialog):
    """PDF 深度诊断与结构探针面板。"""

    def __init__(self, initial_file: Optional[str] = None, parent=None):
        super().__init__("🔍 PDF 深度诊断与受损修复", parent)
        self.resize(880, 640)
        self.setMinimumSize(800, 560)
        self.worker: Optional[DiagnosticWorker] = None
        self.current_json_ast: Optional[Dict[str, Any]] = None

        self.content_layout.setSpacing(10)

        self._build_file_bar(initial_file)
        self._build_tabs()
        self._build_footer()

        if initial_file and os.path.exists(initial_file):
            self.run_diagnosis(initial_file)

    def _build_file_bar(self, initial_file: Optional[str]):
        bar = QFrame()
        bar.setObjectName("wizardCard")
        layout = QHBoxLayout(bar)
        layout.setContentsMargins(12, 8, 12, 8)
        layout.setSpacing(10)

        lbl = QLabel("目标文件：")
        lbl.setStyleSheet("font-weight: 700;")
        self.txt_path = QLineEdit()
        self.txt_path.setObjectName("settingsPathEdit")
        self.txt_path.setFixedHeight(34)
        self.txt_path.setPlaceholderText("选择需要进行深度诊断的 PDF 文件...")
        self.btn_browse = QPushButton("浏览...")
        self.btn_browse.setObjectName("dialogSecondaryBtn")
        self.btn_browse.setFixedHeight(34)
        self.btn_browse.setCursor(Qt.PointingHandCursor)
        self.btn_browse.clicked.connect(self._on_browse_clicked)
        self.btn_diagnose = QPushButton("开始诊断")
        self.btn_diagnose.setObjectName("dialogPrimaryBtn")
        self.btn_diagnose.setFixedHeight(34)
        self.btn_diagnose.setCursor(Qt.PointingHandCursor)
        self.btn_diagnose.clicked.connect(lambda: self.run_diagnosis(self.txt_path.text().strip()))

        layout.addWidget(lbl)
        layout.addWidget(self.txt_path, stretch=1)
        layout.addWidget(self.btn_browse)
        layout.addWidget(self.btn_diagnose)

        self.content_layout.addWidget(bar)

        if initial_file and os.path.exists(initial_file):
            self.txt_path.setText(initial_file)

    def _build_tabs(self):
        self.tabs = QTabWidget()

        # Tab 1: 语法健康体检
        self.tab_check = QWidget()
        v1 = QVBoxLayout(self.tab_check)
        v1.setContentsMargins(0, 10, 0, 0)
        v1.setSpacing(8)

        self.lbl_check_badge = QLabel("等待诊断...")
        self.lbl_check_badge.setStyleSheet(
            "font-size: 14px; font-weight: bold; padding: 6px 12px; background: #E5E7EB; border-radius: 6px;"
        )
        v1.addWidget(self.lbl_check_badge)

        self.txt_check_log = QPlainTextEdit()
        self.txt_check_log.setReadOnly(True)
        self.txt_check_log.setFont(QFont("Consolas", 10))
        v1.addWidget(self.txt_check_log)
        self.tabs.addTab(self.tab_check, "语法健康体检")

        # Tab 2: 加密与权限透视
        self.tab_enc = QWidget()
        v2 = QVBoxLayout(self.tab_enc)
        v2.setContentsMargins(0, 10, 0, 0)
        v2.setSpacing(8)

        self.lbl_enc_badge = QLabel("等待诊断...")
        self.lbl_enc_badge.setStyleSheet(
            "font-size: 13px; font-weight: bold; padding: 4px 10px; background: #E5E7EB; border-radius: 6px;"
        )
        v2.addWidget(self.lbl_enc_badge)

        self.txt_enc_log = QPlainTextEdit()
        self.txt_enc_log.setReadOnly(True)
        self.txt_enc_log.setFont(QFont("Consolas", 10))
        v2.addWidget(self.txt_enc_log)
        self.tabs.addTab(self.tab_enc, "加密与权限透视")

        # Tab 3: 线性化分析
        self.tab_lin = QWidget()
        v3 = QVBoxLayout(self.tab_lin)
        v3.setContentsMargins(0, 10, 0, 0)
        v3.setSpacing(8)

        self.lbl_lin_badge = QLabel("等待诊断...")
        self.lbl_lin_badge.setStyleSheet(
            "font-size: 13px; font-weight: bold; padding: 4px 10px; background: #E5E7EB; border-radius: 6px;"
        )
        v3.addWidget(self.lbl_lin_badge)

        self.txt_lin_log = QPlainTextEdit()
        self.txt_lin_log.setReadOnly(True)
        self.txt_lin_log.setFont(QFont("Consolas", 10))
        v3.addWidget(self.txt_lin_log)
        self.tabs.addTab(self.tab_lin, "线性化(Web快速视图)")

        # Tab 4: 底层 JSON 语法树探针
        self.tab_json = QWidget()
        v4 = QVBoxLayout(self.tab_json)
        v4.setContentsMargins(0, 10, 0, 0)
        v4.setSpacing(8)

        h_json_tools = QHBoxLayout()
        self.lbl_json_info = QLabel("JSON AST 语法树：")
        self.lbl_json_info.setStyleSheet("font-weight: 700;")
        self.btn_export_json = QPushButton("导出 JSON 文件...")
        self.btn_export_json.setObjectName("dialogSecondaryBtn")
        self.btn_export_json.setFixedHeight(30)
        self.btn_export_json.setCursor(Qt.PointingHandCursor)
        self.btn_export_json.setEnabled(False)
        self.btn_export_json.clicked.connect(self._on_export_json_clicked)
        h_json_tools.addWidget(self.lbl_json_info)
        h_json_tools.addStretch()
        h_json_tools.addWidget(self.btn_export_json)
        v4.addLayout(h_json_tools)

        self.tree_json = QTreeWidget()
        self.tree_json.setHeaderLabels(["键 / 节点", "值 / 属性"])
        self.tree_json.header().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.tree_json.header().setSectionResizeMode(1, QHeaderView.Stretch)
        v4.addWidget(self.tree_json)
        self.tabs.addTab(self.tab_json, "底层 JSON 探针")

        # Tab 5: 灾难抢救与受损重建
        self.tab_repair = QWidget()
        v5 = QVBoxLayout(self.tab_repair)
        v5.setContentsMargins(0, 10, 0, 0)
        v5.setSpacing(14)

        card_repair = QFrame()
        card_repair.setObjectName("wizardCard")
        v_rep = QVBoxLayout(card_repair)
        v_rep.setContentsMargins(16, 16, 16, 16)
        v_rep.setSpacing(12)

        lbl_rep_title = QLabel("🛠️ PDF 受损抢救与交叉引用表 (XRef) 重建")
        lbl_rep_title.setStyleSheet("font-size: 15px; font-weight: bold; color: #1E3A8A;")
        lbl_rep_desc = QLabel(
            "当 PDF 文档在 Adobe Acrobat 或其他阅读器中打不开、提示文件损坏、缺少 EOF 标志、\n"
            "或交叉引用表错位时，可通过 qpdf 底层容错解析引擎深度扫描全部内部对象流，\n"
            "强制修复并重新生成标准的 XRef 表与 Trailer 结构。"
        )
        lbl_rep_desc.setWordWrap(True)
        lbl_rep_desc.setStyleSheet("color: #4B5563; line-height: 140%;")

        self.btn_do_repair = QPushButton("尝试强制修复并另存为...")
        self.btn_do_repair.setObjectName("dialogPrimaryBtn")
        self.btn_do_repair.setFixedHeight(34)
        self.btn_do_repair.setCursor(Qt.PointingHandCursor)
        self.btn_do_repair.clicked.connect(self._on_repair_clicked)

        v_rep.addWidget(lbl_rep_title)
        v_rep.addWidget(lbl_rep_desc)
        v_rep.addWidget(self.btn_do_repair, alignment=Qt.AlignLeft)
        v5.addWidget(card_repair)
        v5.addStretch()

        self.tabs.addTab(self.tab_repair, "受损抢救修复")
        self.content_layout.addWidget(self.tabs, stretch=1)

    def _build_footer(self):
        footer = QWidget()
        layout = QHBoxLayout(footer)
        layout.setContentsMargins(0, 4, 0, 0)
        layout.setSpacing(12)

        self.lbl_status = QLabel("就绪")
        self.lbl_status.setStyleSheet("color: gray; font-size: 12px;")
        layout.addWidget(self.lbl_status)
        layout.addStretch()

        self.btn_close = QPushButton("关闭")
        self.btn_close.setObjectName("dialogSecondaryBtn")
        self.btn_close.setFixedHeight(34)
        self.btn_close.setCursor(Qt.PointingHandCursor)
        self.btn_close.clicked.connect(self.reject)
        layout.addWidget(self.btn_close)

        self.content_layout.addWidget(footer)

    def _on_browse_clicked(self):
        f, _ = QFileDialog.getOpenFileName(self, "选择诊断 PDF 文件", "", "PDF 文件 (*.pdf)")
        if f:
            self.txt_path.setText(f)
            self.run_diagnosis(f)

    def run_diagnosis(self, file_path: str):
        if not file_path or not os.path.exists(file_path):
            QMessageBox.warning(self, "提示", "请选择有效的 PDF 文件！")
            return

        self.btn_diagnose.setEnabled(False)
        self.lbl_status.setText("⏳ 正在进行深度诊断与 AST 解析...")

        self.worker = DiagnosticWorker(file_path, parent=self)
        self.worker.finished_signal.connect(self._on_diagnosis_finished)
        self.worker.start()

    def _on_diagnosis_finished(self, data: Dict[str, Any]):
        self.btn_diagnose.setEnabled(True)
        self.lbl_status.setText("诊断完成")

        # 1. 语法检查结果
        check_res = data.get("check", {})
        output_txt = check_res.get("output", "")
        self.txt_check_log.setPlainText(output_txt)

        if check_res.get("is_healthy"):
            self.lbl_check_badge.setText("🟢 语法完整性：正常无损 (0 错误, 0 警告)")
            self.lbl_check_badge.setStyleSheet(
                "font-size: 13px; font-weight: bold; padding: 6px 12px; background: #DCFCE7; color: #166534; border-radius: 6px;"
            )
        elif check_res.get("has_warnings"):
            self.lbl_check_badge.setText("🟡 语法完整性：存在非标警告 (可正常渲染)")
            self.lbl_check_badge.setStyleSheet(
                "font-size: 13px; font-weight: bold; padding: 6px 12px; background: #FEF9C3; color: #854D0E; border-radius: 6px;"
            )
        else:
            self.lbl_check_badge.setText("🔴 语法完整性：检测到结构损坏或语法错误")
            self.lbl_check_badge.setStyleSheet(
                "font-size: 13px; font-weight: bold; padding: 6px 12px; background: #FEE2E2; color: #991B1B; border-radius: 6px;"
            )

        # 2. 加密透视
        enc_txt = data.get("encryption", "")
        self.txt_enc_log.setPlainText(enc_txt)
        if "file is not encrypted" in enc_txt.lower():
            self.lbl_enc_badge.setText("🔓 该文档未加密 (公开无限制)")
            self.lbl_enc_badge.setStyleSheet(
                "font-size: 13px; font-weight: bold; padding: 4px 10px; background: #DCFCE7; color: #166534; border-radius: 6px;"
            )
        else:
            self.lbl_enc_badge.setText("🔒 该文档已加密或包含权限限制")
            self.lbl_enc_badge.setStyleSheet(
                "font-size: 13px; font-weight: bold; padding: 4px 10px; background: #FEF3C7; color: #92400E; border-radius: 6px;"
            )

        # 3. 线性化
        lin_txt = data.get("linearization", "")
        self.txt_lin_log.setPlainText(lin_txt)
        if "linearization data" in lin_txt.lower() and "not linearized" not in lin_txt.lower():
            self.lbl_lin_badge.setText("⚡ 已启用 Web 快速视图 (支持流式边下边看)")
            self.lbl_lin_badge.setStyleSheet(
                "font-size: 13px; font-weight: bold; padding: 4px 10px; background: #DCFCE7; color: #166534; border-radius: 6px;"
            )
        else:
            self.lbl_lin_badge.setText("⏳ 未线性化 (流式网络加载需下载全文件)")
            self.lbl_lin_badge.setStyleSheet(
                "font-size: 13px; font-weight: bold; padding: 4px 10px; background: #F3F4F6; color: #4B5563; border-radius: 6px;"
            )

        # 4. JSON AST 树
        self.current_json_ast = data.get("json_ast")
        self.tree_json.clear()
        if self.current_json_ast:
            self.btn_export_json.setEnabled(True)
            self._populate_json_tree(self.current_json_ast)
        else:
            self.btn_export_json.setEnabled(False)
            self.tree_json.addTopLevelItem(QTreeWidgetItem(["未能解析为 JSON AST", ""]))

    def _populate_json_tree(self, ast_dict: Dict[str, Any]):
        """将 JSON 字典递归渲染到 QTreeWidget。"""

        def add_node(parent_item, key, val, depth):
            if depth > 3:
                return
            if isinstance(val, dict):
                node = QTreeWidgetItem([str(key), f"字典 ({len(val)} 个键)"])
                parent_item.addChild(node) if parent_item else self.tree_json.addTopLevelItem(node)
                for k, v in val.items():
                    add_node(node, k, v, depth + 1)
            elif isinstance(val, list):
                node = QTreeWidgetItem([str(key), f"数组 ({len(val)} 项)"])
                parent_item.addChild(node) if parent_item else self.tree_json.addTopLevelItem(node)
                for i, v in enumerate(val[:60]):
                    add_node(node, f"[{i}]", v, depth + 1)
            else:
                node = QTreeWidgetItem([str(key), str(val)])
                parent_item.addChild(node) if parent_item else self.tree_json.addTopLevelItem(node)

        for k, v in ast_dict.items():
            add_node(None, k, v, 1)

    def _on_export_json_clicked(self):
        if not self.current_json_ast:
            return
        save_path, _ = QFileDialog.getSaveFileName(
            self, "导出 PDF AST JSON", "pdf_ast.json", "JSON 文件 (*.json)"
        )
        if save_path:
            try:
                with open(save_path, "w", encoding="utf-8") as f:
                    json.dump(self.current_json_ast, f, ensure_ascii=False, indent=2)
                QMessageBox.information(self, "导出成功", f"JSON 文件已保存至：\n{save_path}")
            except Exception as e:
                QMessageBox.critical(self, "导出失败", str(e))

    def _on_repair_clicked(self):
        in_path = self.txt_path.text().strip()
        if not in_path or not os.path.exists(in_path):
            QMessageBox.warning(self, "提示", "请选择有效的 PDF 文件！")
            return

        base_name = os.path.splitext(os.path.basename(in_path))[0]
        dir_name = os.path.dirname(in_path)
        default_out = os.path.join(dir_name, f"{base_name}_repaired.pdf")

        save_path, _ = QFileDialog.getSaveFileName(
            self, "另存修复后的 PDF 文件", default_out, "PDF 文件 (*.pdf)"
        )
        if not save_path:
            return

        try:
            res = qpdf.repair_pdf(in_path, save_path)
            if res.is_success:
                QMessageBox.information(
                    self, "修复成功", f"✅ 交叉引用表已强制重建，修复文件已保存：\n{save_path}"
                )
            else:
                QMessageBox.critical(self, "修复失败", res.stderr or "未知错误")
        except Exception as e:
            QMessageBox.critical(self, "修复失败", f"无法修复该文件：\n{str(e)}")
