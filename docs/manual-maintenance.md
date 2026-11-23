# 使用手册维护说明

## 内容入口与范围

正式中文使用手册源文件位于 [manual/index.md](manual/index.md)，共 21 页（含首页和功能索引）。项目 README 和本目录 README 均提供入口。组织方式参考 RATools-for-Word 的安装、第一次使用、逐项功能、保存规则、工作流程和 FAQ 分层。

手册可在仓库或 Markdown 阅读器中直接阅读，也通过独立的 `website/` Astro Starlight 工程生成带中文搜索的静态网站。正文仍只在 `docs/manual/` 维护；站点同步、构建和本地预览见[文档网站说明](docs-site.md)。网站建设不修改 PDF 应用功能，也不构成发布包构建或网站已部署。

核对日期：2026-11-23；源码版本标识 v0.8.2。版本标识不是对所有发布附件与当前工作区一致的保证。

## 功能覆盖依据

| 依据 | 手册覆盖 |
| --- | --- |
| `ratools_pdf/config/rules_catalog.py` | 42 条规则、7 个模块、中国 16 项/美国 13 项预设 |
| `ratools_pdf/ui/main_window.py` | 当前配置菜单、模块全选/清空、压缩档位、检测入口、隐藏工作台、队列与底部控制 |
| `ratools_pdf/config/compression.py` | DPI 72–600/默认 300，JPEG 质量 1–95/默认 85 |
| `ratools_pdf/controllers/main_controller.py`、`precheck_controller.py`、`tree_actions_controller.py` | 队列范围、签名与改名提示、输出目录、建议应用、右键操作、停止/跳过 |
| `ratools_pdf/pdf/processor.py`、`hyperlink_styles.py`、`precheck.py` | 标题、打开设置、清理范围、路径转换、新窗口实际动作与只读检测边界 |
| `ratools_pdf/controllers/io_actions.py`、`io_controller.py`、`workers.py` | CSV/JSON 文件匹配、相对目录、导入位置与发布处理 |
| `ratools_pdf/pdf/bookmarks_links.py`、`ui/dialogs/io_wizard.py` | 书签 CSV 字段、页码基准、链接范围、覆盖/增量及重叠判定 |
| `ratools_pdf/ui/dialogs/security_center_dialog.py` | AES、密码/权限、默认签名选项、解密命名和独立输出 |
| `ratools_pdf/ui/dialogs/settings_dialog.py`、`processing_plan.py` | 主题、输出选项、并行数、0–86400 秒时限和预览状态 |
| `ratools_pdf/ui/dialogs/font_embedding.py` | 外部字体手动交接和保存后复检操作 |
| `ratools_pdf/controllers/log_controller.py`、`ui/dialogs/log_dialog.py`、`about_dialog.py` | 日志筛选、完整记录导出、预检 CSV、审计入口、版本与更新 |
| 现有 `processing-verification.md`、`processing-control.md`、`profiles.md`、`audit-evidence.md`、`capabilities.md` | 处理事实、输出保全、配置与审计格式及验收范围；当前入口名称以源码为准 |

## 写作和审查约定

1. 新增或更改可见规则时，同时更新对应章节和[功能索引](manual/feature-index.md)。对照 `MODULES` 逐项核对，不把引擎元数据中的隐藏能力视为可见勾选框。
2. 对每个功能写清入口、对象范围、是否改原件、结果位置及复核方法。主队列、选中行和独立工具队列必须分开说明。
3. 检查默认选项和实际执行代码，不能只复述 tooltip。当前特别容易误解的是绝对路径仅保留文件名、URI 清理范围、新窗口动作范围、重置缩放同时回到首页，以及覆盖模式不实际改名。
4. 压缩四条底层规则映射到档位；JSON 导入只接受界面可准确表达的组合。新增参数要同步取值范围与默认值。
5. “执行”“实际变化”“独立验证”“已发布”分别描述。未知、未验证、手动和隐藏能力不能写成已通过或自动可用。
6. 修改后检查本手册和被修改入口的相对链接、显式锚点、42 条规则覆盖及预设数量，并做只读内容审查。文档检查不等于软件回归、真实 PDF 试用或法规资格验收。
7. 现有 `assets/screenshots/main-window.png` 是较早布局截图，正文已注明。替换图片时使用真实应用界面和无敏感信息的演示输入；不能把合成界面当作实际截图。

本轮现有入口文件的改前备份位于项目根目录 `backups/user-manual-20261123/`。保留原有开发说明作为技术参考，用户操作以新手册的当前入口及范围说明为主。

## 手册正文交付检查记录

2026-11-23 在本机以 PowerShell 对交付 Markdown 和已读取的规则目录执行只读检查，最终退出码为 0：

- 手册 21 页，功能索引覆盖规则目录 42/42 条标题。
- 手册、维护说明和三个文档入口中的 252 个本地链接目标存在，92 处带显式锚点的引用有效。
- 代码围栏配对，无重复显式锚点或替换字符；预设定义数量为中国 16 项、美国 13 项。
- 三个既有入口文件的 `git diff --check` 通过；改前备份三个文件均已保留。

只读内容审查重点覆盖输入队列与选中行、独立工具输出位置、书签重建与链接增量、安全中心默认签名选项、批处理终态和审计范围。审查发现并修正了覆盖模式不会实际重命名原文件的说明，以及“无书签数据行”与“空数据目录”可能混淆的表述。

上述正文交付阶段没有启动应用、处理真实 PDF、运行软件回归或构建/部署文档站；以上检查仅证明所述文档结构和静态覆盖结果，不构成运行行为或真实业务资格验收。后续网站构建结果单独记录于文档网站说明。
