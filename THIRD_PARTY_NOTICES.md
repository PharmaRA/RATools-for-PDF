# Third-Party Notices

本文件汇总 `RATools-for-PDF` 主要第三方组件的授权信息与源码入口，便于源码分发和打包分发时一并提供。

`RATools-for-PDF` 项目自身源码采用 `GNU AGPL v3` 许可；以下第三方组件仍分别受其各自许可证约束。

| Component | Usage | License | Upstream |
| --- | --- | --- | --- |
| PySide6 | GUI runtime dependency declared in `requirements.txt` | `LGPL-3.0-only OR GPL-2.0-only OR GPL-3.0-only` | https://pypi.org/project/PySide6/ , https://code.qt.io/cgit/pyside/pyside-setup.git/ |
| PyMuPDF | PDF processing runtime dependency (`fitz`) declared in `requirements.txt` | `AGPL-3.0` or Artifex commercial license | https://pypi.org/project/PyMuPDF/ , https://github.com/pymupdf/PyMuPDF |
| Ghostscript 10.06.0 | Bundled Windows executable and DLLs in `plugins/ghostscript/bin/` | `AGPL-3.0` or Artifex commercial license | https://ghostscript.com/releases/ , https://github.com/ArtifexSoftware/ghostpdl |
| pefile | Optional build helper used by `patch_pe_subsystem.py` | `MIT` | https://pypi.org/project/pefile/ , https://github.com/erocarrera/pefile |

## Distribution Notes

1. 本仓库当前直接包含 `Ghostscript 10.06.0` 的 Windows 可执行文件与动态库。对外分发包含这些文件的源码包或二进制包时，应同时保留本文件与根目录 `LICENSE`。
2. 若你分发了包含 `PyMuPDF`、`Ghostscript` 或其本地修改版本的对象代码、冻结打包产物或安装包，应确保接收方可获得对应版本的源码以及你对这些组件所做的修改。
3. 本仓库已经提供 `RATools-for-PDF` 自身源码。若你发布了额外的 PyInstaller、Nuitka 或其他冻结构建产物，建议在分发目录中同步包含 `LICENSE` 和 `THIRD_PARTY_NOTICES.md`。
4. 上表中的上游地址用于定位官方项目、发布页与源码仓库。若你基于某个特定版本二次分发，请记录并保留你实际使用的精确版本信息。
