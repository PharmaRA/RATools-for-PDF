# 当前能力、入口与验收边界

核对日期：2026-11-10。本次 PD10 构建基线：`55ceb0c` 加明确记录的 dirty PD10/用户主题源文件。先从桌面主窗口开始操作；本页开发者 CLI 限于已核对的证据、测试与构建入口。

“稳定”在下表仅表示已有实现和软件回归覆盖的机制，不构成法规、官方认证或全部 PDF 兼容承诺。“实验/隐藏”表示实现存在但当前不向普通桌面用户开放；“手动”需要外部程序与用户操作；“blocked”表示当前不可执行。软件自动化、当前 Windows 成品与真实业务/外部阅读器资格分别列示。PD10 运行源码与实际双变体统一 smoke；范围仅为隔离 GUI 初始化/关闭、合成核心 PDF、真实 qpdf、图片压缩和更新边界。下表其它功能不因该 smoke 升级为逐项成品资格，历史下载包亦不自动包含本轮源码。

## 能力矩阵

| 能力 | 当前 UI / CLI 入口 | 状态 | 软件自动化证据 | 当前成品证据 | 真实资格与已知限制 |
| --- | --- | --- | --- | --- | --- |
| 导入、预检与建议 | 主窗口拖入文件/文件夹 →“预检”“应用建议”“处理建议文件”；文件树详情 | 稳定：只读预检 | 既有 PDF/队列/预检回归，PD02–08 继续覆盖 | 本轮未逐项验收；PD10 smoke 未逐项覆盖 | 辅助结构检查；需要打开密码时先在安全中心解锁，不等同完整合规审阅 |
| PDF 规则批处理 | 各规则模块 →“开始处理”；含页面、视图、书签、链接、清理、压缩与版本/线性化 | 稳定：软件机制；逐规则能力有边界 | 既有真实生成 PDF、qpdf 与 worker 回归；PD07 保留旧管线行为 | 本轮未逐项验收；PD10 smoke 未逐项覆盖 | 异常/复杂 PDF、真实业务及阅读器资格 pending；qpdf 相关规则需要可用引擎 |
| 计划预览与输出保护（PD02） | “开始处理”→“处理计划预览”→确认；取消停止只读预览 | 稳定 | 计划/指纹、真实合成 PDF 保存重开、结构变化、取消/跳过保全回归 | 本轮未验收；PD10 smoke 未逐项覆盖 | 输入、规则或配置变化使计划失效；检查允许后原子发布，结构通过不证明视觉、所有规则或签名有效 |
| 最近批次审计（PD03） | “日志”→“导出最近批次审计包”；源码 CLI 见下文 | 稳定 | 实际 worker/日志按钮、四文件完整性、篡改/缺失/CSV 公式与模态重入回归 | 本轮未验收；PD10 smoke 未逐项覆盖 | 只导出当前窗口内存中最近已结束批次；无历史选择器。包不含 PDF/密码；哈希不认证来源或签名，也不重新核验磁盘 PDF |
| 配置导入/导出与常用（PD04） | 快速预设栏“导入配置 / 导出配置”“保存为常用 / 我的常用” | 稳定：structural | 实际 Qt/QSettings、JSON 往返、差异取消、严格拒绝不完整/互斥组合、旧常用迁移及证据摘要 | 本轮未验收；PD10 smoke 未逐项覆盖 | 只携带规则、参数、ID/版本等；不含输出、覆盖、并行、超时、主题。中国/美国 eCTD 预设非官方认证 |
| 字体风险与手动交接（PD05） | “页面与字体标准化”→选中 PDF →“嵌入缺失字体”→外部手动保存→“保存后复检字体” | 稳定：本地只读检查；手动：外部修复 | 三类风险、不可读/枚举失败、打开失败、fresh recheck 和后台关闭保护回归 | 本轮未验收；PD10 smoke 未逐项覆盖 | 三类计数均为零才表示本次结构检查通过，不证明修复发生、CJK/缺字/视觉正确或字体授权。真实 Acrobat/字体/profile pending |
| 自动字体嵌入 | 无可用自动 UI；生产 provider 的 repair 返回 blocked | blocked | 验证安装/探测 Acrobat 仍不解锁、假成功不得跳过独立复检 | 不提供自动修复成品能力 | 旧 COM/OLE fixup 已退役；需先证明可靠受支持 backend，不以检测到 Acrobat 替代资格 |
| 超时、取消与资源控制（PD06） | “全局设置 > 处理性能 > 每文件超时”；处理时停止、串行跳过或并行终止选中项，结束后重试失败项 | 稳定：Windows 软件机制 | 受控真实父子树/owner 退出、stop/skip/deadline、发布锁、低资源与设置持久化回归 | 本轮未验收；PD10 smoke 未逐项覆盖 | 0 关闭，1–86400 秒；单次 OS/fitz 调用不能保证立即中断。Windows 进程树管理需 Windows 10+；外部受限 Job、网络盘/移动设备和生产大 PDF 峰值 pending |
| 规则执行事实（PD07） | 日志修改摘要及审计 JSON 的 rule_results/rule_events；开发 API 见[规则结果](rule-results.md) | 稳定：有限观测/检查 | mutation 观测、smart 跳过、复合规则归因、部分失败及子进程→worker→批次→审计回归 | 本轮未验收；PD10 smoke 未逐项覆盖 | 选中、执行、变化、验证分开。后置检查仅覆盖标题、A4/Letter、删除 URI、普通批注和附件的对应规则；其余及复合条件保留 unknown/null、not_verified |
| 安全中心：密码与权限 | 主窗口顶部“安全中心”→“PDF 加密与权限保护”或“PDF 解密与权限脱壳” | 可用：独立工作流，非主队列规则 | 既有 qpdf 加解密/输出保全与线程生命周期回归；PD06 继续覆盖直接依赖 | 本轮未逐项验收；PD10 smoke 未逐项覆盖 | 支持已知打开/权限密码；未知密码不破解。免密仅适用于可打开的权限受限文件。解锁输出再加入主队列；不据此声称独立安全资格 |
| 签名外观与表单限制 | 安全中心解密页“解除数字签名限制”“移除表单字典” | 可用：需人工风险判断 | 既有结构/外观相关回归不构成证书验证 | 本轮未验收；PD10 smoke 未逐项覆盖 | 外观固化会改变文档，不保留密码学签名有效性、身份或信任链保证；保留原件，signature_validity 仍 not_verified |
| 文档检测、书签/链接 IO | “文档检测”两只读按钮；“书签管理 / 超链接处理”对应数据导入/导出按钮 | 稳定：现有范围 | 既有检测、匹配、相对目录、导入输出保全与线程回归 | 本轮未逐项验收；PD10 smoke 未逐项覆盖 | 失效引用只匹配 Word 固定占位文案；数据匹配不等同目标链接业务正确性 |
| 页面装配、图层套印、深度诊断 | 普通界面无开放入口；保留后端与对话框 | 实验/隐藏 | 部分既有后端/输出安全回归；不能据此升级为桌面资格 | 无本轮可见入口验收 | 顶部三按钮隐藏且未加入导航；“文档检测”未挂载已创建的诊断按钮。没有用户开启开关；不指导调用隐藏方法绕过入口 |
| 分层测试与样本资格（PD01/PD08） | 开发者安全 runner 的 unit / integration / qualification；见下文 | 稳定：验证机制；外部资格 pending | PD10 最终普通回归 575/575（unit 11 + integration 564）；qualification not_selected | 当前双成品仅完成上述五项合成 smoke；真实样本矩阵未执行 | 历史资格报告仍是独立快照；本轮未运行真实样本/外部阅读器与批准流程，不能算真实资格通过 |
| Windows 构建与更新变体 | 安全 PyInstaller；源码/EXE `--verify-release NEWDIR`；桌面“关于”，NoUpdate 无更新入口 | 已实现：锁定、可追溯、unsigned | wheel 哈希/闭包/RECORD/实际导入与 Analysis 归属闸门，来源/SBOM/许可证及全文件哈希 | 双变体实际 hidden smoke：GUI、合成核心 PDF、qpdf、图片压缩、更新边界；实际路径见[构建说明](build-verification.md) | NoUpdate 无按钮/import/可导入模块/worker，观察到 0 次 Python socket 请求；普通版不测在线更新。签名/许可法律审查/跨平台/真实业务资格 pending，本地构建不称位级可复现 |

上表当前软件数字来自 **PD10 最终安全回归：575/575，180.985 秒**；完整日志和各层机器摘要见[构建说明](build-verification.md)。qualification 本轮未选择，不与软件测试数相加。PD08 等更早阶段数字是各自提交时的历史快照，不作为当前额外通过数。

## 按桌面任务使用

1. **安全地批量处理**：导入 → 预检 → 选规则/配置 → 开始处理 → 审阅计划 → 确认 → 查看结果。先保留原件；详见[计划与验证](processing-verification.md)。
2. **复用处理方案**：主窗口导出 JSON → 另一台机器导入 → 阅读差异 → 确认；详见[配置说明](profiles.md)。
3. **复核并归档**：结束后在“日志”导出最近批次审计包；核验完整性后另行管理归档与来源认证，详见[审计包说明](audit-evidence.md)。
4. **先处理受密码保护的 PDF**：顶部安全中心使用已知密码另存解锁副本，再加入主队列。签名外观和密码学有效性分开检查，详见 [FAQ](../FAQ.md)。
5. **处理字体风险**：选中文件 → 手动交接外部程序 → 保存 → 返回复检；详见[字体说明](font-providers.md)。

## 开发者验证入口

以下命令在具备项目依赖的项目根目录执行。它们是实际存在的源码工具入口，不表示发布 EXE 提供同名通用 CLI：

```powershell
# 审计包完整性（参数替换为实际四文件包目录）
python -B -m ratools_pdf.services.audit_evidence verify ".\RATools_Audit_20261109_120000_example"

# 安全普通回归，以及独立层；不要直接运行 unittest discover
python -X utf8 -X faulthandler -B scripts/run_tests_safely.py
python -X utf8 -X faulthandler -B scripts/run_tests_safely.py --layer unit
python -X utf8 -X faulthandler -B scripts/run_tests_safely.py --layer integration

# 显式资格；输出目录必须尚不存在，下一次更换名称
python -X utf8 -X faulthandler -B scripts/run_tests_safely.py --layer qualification --qualification-manifest docs/qualification/pd01-manifest-v1.json --qualification-output backups/pd09-example-qualification-001

# 精确锁定运行/构建闭包的受保护入口
python -B scripts/build_pyinstaller_safely.py --install-dependency
```

测试层口径、样本版本/人工批准与 CI 选择见[测试说明](testing.md)。资格缺材料返回非零是有效待办，完整边界见[资格矩阵](qualification/README.md)。构建 guard、精确锁、代理、变体、实际 EXE 统一检查和证据位置见[构建说明](build-verification.md)。

## 入口核对依据

本页只对照实际相关路由与直接交付说明：`main_window.py` 的顶部导航、模块按钮、配置栏与开始/日志按钮，`main_controller.py` 的预览与字体接线，`security_center_dialog.py` 的双选项卡，`settings_dialog.py` 的时限控件，以及 `log_dialog.py` 的审计按钮。

`config/features.py` 当前只有更新检查开关；页面装配、套印和顶部诊断使用固定隐藏设置，并不存在面向用户的实验功能开关。各项证据的详细语义以本页链接的 PD02–08 说明为准；原始审查/计划保留在[文档导航](README.md)，历史变更见 [CHANGELOG](../CHANGELOG.md)。
