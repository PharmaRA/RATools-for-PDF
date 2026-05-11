# RATools-for-PDF

专为 RA 递交场景开发的桌面端 PDF 批量处理工具，适用于 eCTD 资料整理、合规清理和结构标准化。

项目基于 `PySide6` 构建图形界面，基于 `PyMuPDF` 实现 PDF 底层处理，并使用 `qpdf` 完成版本转换与线性化等结构级优化。项目源码采用 `GNU AGPL v3`，第三方组件声明见 `THIRD_PARTY_NOTICES.md`。

## 项目特性

- 面向 RA / eCTD 场景的 PDF 批量处理桌面工具
- 支持拖拽导入 PDF 文件或整个文件夹，并自动递归收集 PDF
- 内置中国 eCTD、美国 eCTD 两套快速预设
- 支持批量处理并保留原始目录层级输出
- 支持处理中停止整批任务或跳过当前文件
- 支持处理日志查看与导出，便于复核处理结果
- 支持书签、超链接的批量导出与批量导入
- 书签与超链接导入导出时可保留相对目录层级，避免同名 PDF 数据互相覆盖
- 支持默认输出目录、自动打开输出目录与覆盖原文件风险控制
- 支持在“关于”中手动检查 GitHub Releases 更新
- 支持对重大更新进行启动提醒，并提供无更新能力的 `NoUpdate` 发布变体

## 当前支持的核心能力

### 1. 初始视图与文档属性

- 设为首页打开
- 重置页面布局
- 重置缩放比例
- 设置导览标签（有书签时显示书签面板）
- 折叠所有书签
- 根据文件名自动写入 PDF 标题属性

### 2. 页面与字体标准化

- 批量转换页面为 A4
- 批量转换页面为 Letter
- 批量嵌入非标准字体（暂不可用）
- PDF 版本转换
- 启用线性化（快速网页浏览）

### 3. 书签管理

- 书签设为承前缩放
- 书签动作改为新窗口打开
- 删除书签中的外部链接
- 删除失效书签
- 删除未知动作书签
- 书签批量导出为 CSV
- 从 CSV 批量导入书签

### 4. 超链接处理

- 将外部文件链接的绝对路径转为相对路径
- 超链接设为承前缩放
- 超链接动作改为新窗口打开
- 链接文本设为蓝色
- 增加或删除链接边框
- 统一有框/无框蓝字链接样式
- 超链接批量导出为 JSON
- 从 JSON 批量导入超链接

### 5. 内容合规与安全性

- 删除外部 URI 链接
- 删除外部 URI 链接并将文字改为黑色
- 删除失效或无效链接
- 删除未知动作链接
- 删除 JavaScript、3D 或动态内容
- 删除文档附件
- 删除文档标签
- 删除 PDF 注释
- 删除文档元数据
- 一键删除所有链接和书签

### 6. 文件级优化与输出

- 按 eCTD 规则规范化输出文件名
- 处理完成后自动打开输出目录
- 可选覆盖原始文件（高风险操作）
- 批量输出到 `RATools_Output` 目录

### 7. 批处理控制与日志

- 实时显示当前处理进度、完成数量与当前文件名
- 可手动停止当前整批任务
- 可跳过当前卡住或无需继续处理的单个文件
- 日志支持窗口查看，并可导出为 `CSV` 或 `TXT`
- 处理中会锁定队列移除/清空操作，避免状态错乱

## 界面工作流

1. 启动程序后，将 PDF 文件或文件夹拖入左侧待处理队列。
2. 在右侧选择处理模块中的具体规则，或直接套用中国 / 美国 eCTD 预设。
3. 如有需要，可在“全局设置”中预先配置默认输出目录、自动打开输出目录等选项。
4. 点击“开始批量处理”。
5. 处理过程中可按需停止整批任务，或跳过当前文件。
6. 若未启用覆盖原文件，程序会要求选择输出根目录，并在其中生成 `RATools_Output`。
7. 处理结束后可查看日志，或按设置自动打开输出目录。

此外，文件树支持：

- 右键移除选中项
- 定位到文件位置
- 查看文件详情
- 双击使用系统默认程序打开 PDF

## 书签与链接数据导入导出

除规则处理外，程序还提供独立的数据 IO 功能：

- 书签导出：将 PDF 书签导出为 `CSV`
- 书签导入：从对应 `CSV` 恢复书签结构
- 链接导出：将页面中的链接区域与动作信息导出为 `JSON`
- 链接导入：从 `JSON` 重新写入链接信息

导入导出时会按源 PDF 的相对目录层级组织数据文件和输出结果。例如同时处理 `a/report.pdf` 与 `b/report.pdf` 时，会分别生成到对应子目录中，避免因文件同名导致 `CSV`、`JSON` 或导入后的 PDF 相互覆盖。

适合用于：

- 模板文档结构迁移
- 处理前后数据比对
- 批量修复书签或链接

## 技术栈

- Python 3
- PySide6
- PyMuPDF (`fitz`)
- qpdf（版本转换与线性化依赖）

## 项目结构

```text
RATools-for-PDF/
├─ main.py            # 程序入口
├─ main_no_update.py  # 无更新版程序入口
├─ app_features.py    # 功能开关（如更新能力）
├─ app_version.py     # 应用版本与发布元数据
├─ app_paths.py       # 运行目录/资源目录解析
├─ view.py            # 界面层，负责 UI、预设、设置加载
├─ controller.py      # 控制层，负责事件绑定、任务调度、线程处理
├─ pdf_processor.py   # PDF 核心处理引擎
├─ update_checker.py  # GitHub Releases 更新检查逻辑
├─ build_pyinstaller.bat
├─ build_nuitka.bat
├─ patch_pe_subsystem.py
├─ icon.png           # 程序图标
└─ README.md
```

## 安装与运行

### 1. 使用 GitHub Release（推荐）

如果你只是使用本工具，建议优先下载 GitHub Releases 中提供的 Windows 打包版本。

使用方式：

1. 从 GitHub Releases 下载最新的 Windows 发布包。
2. 根据使用场景选择带更新提醒版或 `NoUpdate` 无更新版。
3. 将整个发布包完整解压到本地目录。
4. 运行对应目录中的 `RATools-for-PDF.exe` 或 `RATools-for-PDF-NoUpdate.exe`。

说明：

- 请保留发布包中的完整目录结构，不要只单独拷贝 `.exe`
- Windows 发布包已包含运行所需的界面依赖和 qpdf 相关资源，通常无需额外安装 Python 或 qpdf
- 首次运行后，程序会在可执行文件所在目录生成 `settings.ini`
- `NoUpdate` 变体不会显示检查更新入口，也不会在启动时访问 GitHub 检查更新

### 2. 从源码运行

如果你需要本地开发、调试或自行修改程序，可直接运行源码。

安装项目依赖：

```bash
pip install -r requirements.txt
```

以下能力会调用 qpdf：

- PDF 版本转换
- 快速网页浏览 / 线性化

### 3. 启动源码程序

```bash
python main.py
```

如需以无更新能力模式启动源码程序，可使用：

```bash
python main_no_update.py
```

说明：

- 程序入口已内置 `multiprocessing.freeze_support()`，用于兼容冻结打包后的批处理子进程启动
- Windows 冻结构建下会主动尝试脱离控制台窗口，避免桌面版程序额外弹出黑框

## PyInstaller 打包（Windows 推荐）

如果你希望更快完成 Windows 桌面分发，推荐使用 `PyInstaller` 的 `onedir` 模式。

### 1. 一键打包

项目根目录已提供 Windows 打包脚本：

```bat
build_pyinstaller.bat
```

脚本会自动：

- 检查 `python` 是否可用
- 检查并安装 `PyInstaller`
- 使用 `onedir` 模式分别打包带更新版和 `NoUpdate` 无更新版
- 自动将生成的 exe 修正为 GUI 子系统
- 从 `app_version.py` 动态生成临时 Windows 版本信息并写入 exe
- 一并带上 `icon.png`
- 一并带上 `plugins/qpdf/` 目录及其运行时 DLL
- 移除 `qopensslbackend.dll` 以减少部分环境下的启动 DLL 冲突

首次使用该链路前，建议确保环境中已安装 `pefile`，因为 GUI 子系统修正脚本 `patch_pe_subsystem.py` 依赖它：

```bash
pip install pefile
```

### 2. 输出位置

打包完成后，程序输出在：

```text
dist/RATools-for-PDF_v0.3.3/RATools-for-PDF.exe
dist/RATools-for-PDF-NoUpdate_v0.3.3/RATools-for-PDF-NoUpdate.exe
```

说明：

- 请直接分发整个对应目录，不要只单独拷贝 `.exe`
- 输出目录会自动附带 `_v版本号` 后缀，便于区分不同发布版本；`exe` 文件名保持稳定不带版本号
- `settings.ini` 会在程序运行后自动生成到可执行文件所在目录
- 当前打包方案针对 Windows 桌面环境设计，默认不显示控制台窗口
- `RATools-for-PDF-NoUpdate` 会在打包时排除 `update_checker` 模块，用于对联网更敏感的分发场景

### 3. 为什么推荐 `onedir`

虽然单文件 `onefile` 分发更方便，但它通常需要在启动时先解包，启动速度反而更慢。对于本项目这种包含 `PySide6`、`PyMuPDF` 和 `qpdf` 资源的桌面工具，`onedir` 模式通常更合适。

### 4. 为什么不用直接 `windowed`

当前默认脚本会先使用更稳定的 console bootloader 构建，再将最终 exe 修正为 GUI 子系统。这样可以避免 `windowed` bootloader 在部分环境下出现的启动兼容性问题，同时保持桌面程序不显示控制台窗口。

### 5. Nuitka 说明

仓库中仍保留 `build_nuitka.bat` 作为可选方案。如果你更看重正式发布时的启动效率，可以再尝试 `Nuitka`；如果你更看重打包速度和成功率，优先使用 `PyInstaller onedir`。

`build_nuitka.bat` 当前也会将输出目录自动命名为 `main_v版本号.dist`，便于区分不同构建版本；`main.exe` 文件名保持不变。

## qpdf 说明

`PDF版本转换` 与 `启用线性化 (快速网页浏览)` 当前默认使用 `qpdf` 执行，以尽量保留 PDF 内部目录、书签和链接结构。

程序会在不同平台按如下方式查找 qpdf：

- Windows：优先 `plugins/qpdf/qpdf.exe`
- 其次：环境变量 `QPDF_PATH`
- 再次：常见系统安装路径或 `PATH` 中的 `qpdf.exe`
- macOS / Linux：直接调用系统中的 `qpdf`

说明：

- 当前仓库可附带 `plugins/qpdf/` 中的 `qpdf.exe` 与配套 DLL 作为 Windows 发布包内置依赖，避免终端用户额外安装
- 冻结打包后，`plugins/qpdf/` 会随发布目录一并分发
- 若程序未找到可用的 qpdf，可导致 `PDF版本转换` 与 `启用线性化` 功能失败
- 重新分发包含 qpdf 二进制的构建产物时，应同时保留 `LICENSE` 与 `THIRD_PARTY_NOTICES.md`

## Third-Party Notices

本项目使用并可能在分发产物中包含第三方组件。各第三方组件仍分别受其自身许可证约束，详见 `THIRD_PARTY_NOTICES.md`。

## 配置说明

项目使用根目录下的 `settings.ini` 保存本地选项，包括：

- 默认输出目录
- 是否自动打开输出目录
- 是否覆盖原文件
- 各模块规则勾选状态

说明：

- 规则勾选状态会持久化保存
- 默认输出目录也会持久化保存，用于处理输出选择和日志导出的默认位置
- 预设按钮不会完全跟随上次会话恢复，默认按当前界面逻辑重新进入自定义/预设状态

## 输出与日志

- 默认输出目录为用户选择目录下的 `RATools_Output`
- 导入书签 / 链接时，会在首个 PDF 所在目录下生成 `RATools_导入完成`，并保留源目录相对层级
- 日志可在界面中查看，并支持导出为 `CSV Summary` 或 `TXT`
- 日志导出会优先使用最近一次输出目录或全局设置中的默认输出目录作为初始保存位置

## 注意事项

- 仅会自动收集并处理 `.pdf` 文件
- 若启用“覆盖原始文件”，程序会先弹出确认，这是不可逆操作
- 批量处理执行期间，待处理队列不允许再清空或移除，以避免任务状态错乱
- “跳过当前文件”仅在批处理执行期间可用，用于不中断整批任务地跳过单个文件
- 加密 PDF 会被直接跳过，并记录为处理失败
- 部分清理、链接修复和内容重写操作依赖 PDF 本身结构，遇到异常文档时可能出现个别规则无法完全生效
- 导入书签和导入链接时，需要保证目标目录中存在与 PDF 文件名对应的数据文件

## License

本项目源码采用 `GNU AGPL v3` 许可证，详见 `LICENSE`。

第三方组件的授权、上游项目地址及源码获取说明见 `THIRD_PARTY_NOTICES.md`。
