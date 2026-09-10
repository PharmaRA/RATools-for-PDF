"""对话框包。各对话框一文件；本模块 re-export 保持
`from ratools_pdf.ui.dialogs import X` 的历史导入路径不变。"""

from ratools_pdf.ui.dialogs.about_dialog import AboutDialog
from ratools_pdf.ui.dialogs.base import FramelessDraggableDialog
from ratools_pdf.ui.dialogs.dpi_selection_dialog import DpiSelectionDialog
from ratools_pdf.ui.dialogs.font_embedding import ManualFontEmbeddingDialog
from ratools_pdf.ui.dialogs.io_wizard import IODataWizardDialog
from ratools_pdf.ui.dialogs.log_dialog import LogDialog
from ratools_pdf.ui.dialogs.message_box import CustomMessageBox
from ratools_pdf.ui.dialogs.overlay_dialog import OverlayUnderlayDialog
from ratools_pdf.ui.dialogs.page_assembly_dialog import PageAssemblyDialog
from ratools_pdf.ui.dialogs.pdf_inspector_dialog import PdfInspectorDialog
from ratools_pdf.ui.dialogs.security_center_dialog import SecurityCenterDialog
from ratools_pdf.ui.dialogs.settings_dialog import SettingsDialog

__all__ = [
    "AboutDialog",
    "CustomMessageBox",
    "DpiSelectionDialog",
    "FramelessDraggableDialog",
    "IODataWizardDialog",
    "LogDialog",
    "ManualFontEmbeddingDialog",
    "OverlayUnderlayDialog",
    "PageAssemblyDialog",
    "PdfInspectorDialog",
    "SecurityCenterDialog",
    "SettingsDialog",
]
