# 字体 provider 与保存后复检（PD-05，契约 v1）

当前可用的是 **Acrobat/默认 PDF 程序手动交接，以及 RATools 只读字体复检**。自动嵌入字体仍不可用。检测到 Acrobat.exe 只用于选择打开程序，不证明 Acrobat Pro、COM、Preflight profile 或 fixup 执行能力已通过验证。

既有产品决定见 [Acrobat 字体嵌入自动化探索记录](acrobat-font-embedding-automation-notes.md)：COM/OLE 脚本上下文执行 Preflight fixup 不可靠，因此自动修复已退役。PD-05 保留这一决定，不重新启用旧 `embed_nonstandard_fonts` 配置，不添加字体引擎、自动 COM provider 或整页栅格化兜底。

## 在界面中使用

1. 在左侧队列选中目标 PDF，点击“嵌入缺失字体”。
2. 对话框展示平台、provider 契约版本、Acrobat 版本未知及自动修复不可用原因。
3. 点击“打开 Acrobat”。没有定位到 Acrobat 时，系统会尝试默认 PDF 程序；该程序可能没有字体修复工具。打开请求提交后状态保持待操作，不会显示字体修复成功。
4. 在 Acrobat 手动进入“所有工具 > 印刷制作 > 印前检查”，执行“嵌入缺失的字体”，并保存。需要保留原件时，应先自行另存为副本，并对选中的目标文件执行复检。
5. 回到对话框点击“保存后复检字体”。RATools 在后台重新读取选中的磁盘文件，分别显示未嵌入、替代风险、嵌入状态未知三个计数。只有每个文件可检查且三类计数均为 0，才显示本次字体复检通过。没有字体的 PDF 也可以通过检查；这不表示发生过字体修复。

复检期间关闭按钮禁用，关闭窗口/按 Esc 不会销毁工作线程；结束后恢复按钮。字体模块不修改 PDF、不发布输出、不生成持久化文档正文或路径日志，因此也不会覆盖原件或已有输出。用户在外部 Acrobat 中的保存行为由用户控制。

## Python 接口

入口为 `ratools_pdf.services.font_providers`：

| 接口 | 返回值与语义 |
| --- | --- |
| `AcrobatManualProvider().capability()` / `FontWorkflow.capability()` | `FontCapability`：`platform`、`provider`、适配契约 `version`、外部软件 `external_version`、`availability`、`supported_operations`、逐操作 `unavailable_reasons`、`acrobat_detected`、`external_qualification`。不输出程序绝对路径。 |
| `FontWorkflow.inspect(path)` | `FontInspection`：当前磁盘 PDF 的三类计数、只读 findings、状态和理由。 |
| `FontWorkflow.recheck(path)` | 每次重新调用与预检相同的公开字体检查器，不缓存上次结果；结果格式同 inspect。 |
| `FontWorkflow.open_manual(path)` | `FontOperationResult`：`manual_required` 表示打开请求已交接，`opened=True`；始终不等于字体已修复或应用已实际打开。 |
| `FontWorkflow.repair(path)` | 当前生产适配器返回 `blocked` 与已退役原因。若将来出现经过资格验证的受支持 backend，只有其返回 `reported_success` 且独立 fresh recheck 可用、三类风险均为 0，才可能得到 `verified`。当前仓库没有这样的自动 backend。 |

PDF 公开只读入口是 `ratools_pdf.pdf.inspect.collect_font_precheck_for_path` 与 `font_precheck_has_embedding_risk`；二者复用既有 `pdf.precheck` 实现。字体风险口径没有改变：`has_unembedded`、`substitution_risk`、`has_unknown_embedding_status` 任一为真均阻止通过。任何页面字体枚举失败会使文件级字体检查不可用，不会跳过该页后把不完整结果当作零风险。

`availability` 为 `manual_only`（探测到 Acrobat）、`default_app_only`（仅可尝试默认程序）或 `unsupported`（平台不支持交接）。这些值表示打开方式，不证明外部程序能执行修复。不支持交接的平台仍可运行本地字体检查。

检查结果为 `verified`、`risks_remain`、`unavailable` 或 `timeout`。无法检查时三个计数为 `None`，不会以零计数掩盖失败。操作结果区分 `blocked`、`unavailable`、`manual_required`、`failed`、`timeout`、`recheck_failed` 和 `verified`。异常只展示类别与通用原因，不展示异常中可能包含的文档绝对路径。底层抛出的超时被分类为 `timeout`；当前同步本地检查器和系统打开调用没有新增强制终止时限，也没有正在等待的自动 COM 作业。

`FontProvider` protocol 用于表达尝试与能力边界。假自动成功/失败 provider 仅定义于测试中。把 provider 声称成功直接升级为修复成功会违反此契约。

## 验证范围与外部缺口

软件验证覆盖缺 Acrobat、不支持平台、探测异常、打开失败/超时、provider 异常/超时/假成功、不可读 PDF、真实生成可解析 PDF、同一路径保存后重新读取、原始字节不变、后台复检状态恢复与关闭保护。现有预检直接回归亦运行。

真实 Acrobat 手动修复、本机字体文件/CJK 字体、字体授权、指定 profile、缺字/视觉正确性、操作系统打开关联和外部版本仍为 **pending 外部验收**。检查为零只反映既有结构性字体预检能力，不构成视觉、法规或完整字体资格认证。自动修复为 **blocked**，需要先证明存在受支持且可靠的 backend，不能因安装 Acrobat 就解除。
