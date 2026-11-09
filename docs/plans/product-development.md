# 功能与开发质量实施计划

## 目标与授权

用户于 2026-11-09 接受十项建议，要求先落文档，再按顺序实施，每完成一项独立 commit。本计划记录范围、验收、提交和缺口；不替代[既有 18 项质量修复记录](quality-improvement.md)。

基线：`cc6150de1ee9364afa8a99ba0943e880d403204c`。本轮开始实测 `python -X utf8 -X faulthandler -B scripts/run_tests_safely.py`，401/401 通过，65.044 秒；现存 fitz 弃用提示保留记录。历史 EXE 图片压缩自检不等于完整发布验收。

## Global Constraints

1. 顺序为 PD-01 至 PD-10；每项先有行为验收与失败复现，再实现、验证、只读子代理 review，最后由主代理独立提交。文档基线可单独提交。不得把提交动作前移到 review 之前。
2. 保留用户未提交的 `ratools_pdf/ui/theme.py` 与 `.playwright-mcp/`；仅暂存本任务明确列出的文件。提交使用 Fiveonine 身份，禁止输出邮箱、凭据、密码和业务隐私。
3. 重要现有文件修改前备份到根目录 `backups/product-roadmap-20261109/`；禁止物理删除，临时产物仅回收，回收失败保留并报告。任何链接/重解析点/硬链接先审查。
4. 每次动作前说明目标与最小读取范围；读取限本项交付物和直接依赖。不得全仓库扫描、顺手重构、修改主题或读取无关旧档案。
5. 本轮在用户当前工作区实施；逐项精确提交，保持其原有修改。不得自动远程推送、发版、购买证书或读取密钥。
6. 安全测试入口为 `python -X utf8 -X faulthandler -B scripts/run_tests_safely.py [test_*.py]`。每项执行专项；生产变更提交前跑一次全量。真实外部系统未运行时标记 pending/blocked，不能将合成 PDF、mock、历史测试升级为真实资格证据。
7. 软件实现完成与真实资格完成分开记录。缺少非敏感业务样本、Acrobat、证书或外部设备时，先交付可执行机制和明确缺口，继续其余已授权开发；最终不得声称全部业务验收完成。
8. 输出默认保留原件；不持久化文档正文、密码或个人身份。证据默认使用样本 ID、相对路径、哈希和结构指标，绝对路径与文档元数据必须最小化。
9. 后续项消费前项明确版本的接口；PD-07 将 PD-02/03 的实际执行事件统一，不能先编造规则级 changed/verified 值。无法判定时明确 unknown/not_verified。
10. 既有成功输出原子发布与失败保全、队列互斥、预检失效和线程退出保障必须保留。

## 顺序与状态

| ID | 交付 | 软件状态 | 真实环境验收 | 提交 |
|---|---|---|---|---|
| PD-01 | PDF 资格矩阵与证据运行器 | 已完成 | 外部来源、成品和阅读器 pending | 8ce2d3d |
| PD-02 | 处理计划预览与处理后验证 | 已完成 | 依赖 PD-01 真实样本覆盖 | aa817d0 |
| PD-03 | 可归档审计证据包 | 已完成 | 内部质量系统采纳待业务确认 | 30fa400 |
| PD-04 | 可版本化配置与导入导出 | 已完成 | 机构规则有效性待专家确认 | 3154a17 |
| PD-05 | 字体 provider 能力与复检 | 手动交接/复检已完成 | 自动修复 blocked；Acrobat/CJK pending | 9428b39 |
| PD-06 | 超时、取消、进程树与资源限额 | 已完成 | 网络盘/设备故障待实测 | 7e32bb3 |
| PD-07 | 统一规则执行契约与结果 | 已完成 | 复杂规则观测/独立验证仍有限 | 0d1d10f |
| PD-08 | 分层测试与版本化样本回归 | 已完成 | 外部样本/人工批准仍 pending | 2cfff13 |
| PD-09 | 能力矩阵与文档一致性 | 已完成 | 只描述已有证据 | 5f96a04 |
| PD-10 | 可追溯发布与双变体验收 | 已完成 | unsigned；外部资格仍 pending | 见 PD-10 记录 |

## Task 1: PD-01 PDF 资格矩阵与证据运行器

交付：`docs/qualification/README.md`、版本化矩阵/样本清单、只读验证服务和显式 CLI，以及针对性测试。先阅读既有 `pdf/inspect.py`、成品压缩验证入口及安全文件工具的直接依赖，再确定最小复用范围。

- 矩阵区分自动化回归、Windows 成品、真实业务样本、外部阅读器四种证据层次；状态为 passed/failed/pending/blocked，不允许缺失样本算通过。
- 样本条目包含稳定 ID、来源工具/版本、来源类型（真实来源或合成）、授权/非敏感声明、SHA-256、特征、处理选项和预期指标；样本路径限定在显式样本根目录，禁止逃逸和链接绕过。
- 来源覆盖要求：Acrobat、Word、LibreOffice、扫描仪、PDF 打印驱动；特征覆盖 CJK、透明图、表单、签名、增量保存、命名/远程目标、加密、损坏、超大文件。不存在的样本保持 pending，禁止伪造来源。
- 运行器校验输入哈希，只读重新打开/提取结构/渲染指定页，并可执行 qpdf 检查；保留标准化报告、引擎版本与各检查结果。原件哈希前后相同，输出写入新目录，拒绝覆盖旧证据。
- Acrobat Reader/浏览器手工复核提供具备样本与输出哈希、工具版本、检查项和结果的记录模板。没有记录时保持 pending。
- 初始软件验收使用明确标识的合成 PDF，验证真结构、真渲染、qpdf、哈希失配、损坏文件、缺少样本、路径逃逸和报告不可覆盖。
- 完成边界：可以提交机制及其测试；真实来源 PDF 与阅读器资格独立开放，不能以机制通过宣布这些矩阵行通过。

## Task 2: PD-02 处理计划预览与处理后验证

交付：独立计划/结构快照模型，主批次入口的计划预览，实际输出的验证结果与 UI 状态。

- 按文件显示用户选中规则、预检命中、计划执行/跳过及原因、签名/表单/附件/几何变化提示；不支持预检的规则明确为用户选择后执行。
- 计划绑定输入指纹、选项和配置版本；执行前再检查，文件变化则重新预检/拒绝过期计划。
- 处理前后记录 SHA-256、字节数、页数、页面尺寸/旋转、书签/链接/批注/附件/表单/签名、PDF 版本；未知值不可写成 0。
- 保存重开验证与业务预期分别记录；检查自身失败必须可见，不得仍显示“验证通过”。几何、注释等预期变化不能简单当作损坏。
- 运行记录先表示文件级计划/处理/验证事实，规则级实际 changed 由 PD-07 补齐。
- 验收包含过期输入、失败输出保全、计划取消不写 PDF、正常处理结构对照、未知检查、UI 忙碌互斥和批次结果归属。

## Task 3: PD-03 可归档审计证据包

交付：从日志/批次结果可选导出证据包，含 JSON 主清单与 CSV 汇总，引用 PD-02 实际结果。

- 包含输入/输出相对清单及哈希、应用和引擎版本、配置版本、稳定规则 ID/参数、开始/结束时间、处理和验证结果、警告/失败/跳过原因。
- 预检快照与处理后验证使用明确版本 schema；不得从中文日志猜测事实，不保存正文、密码、凭据或用户系统路径。
- 生成文件完整性清单；新目录写入，错误不覆盖已有包。CSV 防公式注入，Unicode 与多行原因可往返读取。
- 部分失败和取消批次也可导出；缺失证据明确说明，不能静默补成功。哈希清单提供可调用核验入口。
- 验收：真实小批次、失败/取消、同名不同目录、多行/公式字段、包内容篡改与版本字段完整性。

## Task 4: PD-04 可版本化的配置

交付：配置 schema、内置配置版本、导入/导出/差异预览，接入“我的常用”和执行证据。

- 字段至少为 profile_id、profile_version、agency_or_region、enabled_rules、rule_parameters、validation_level、schema_version。
- 校验规则 ID、参数类型/范围、互斥组合和 schema 版本；未知必需字段/未来版本明确拒绝，禁止静默丢规则。
- 内置中美预设保持现有行为，记录来源/适用范围，不能暗示官方认证；旧 settings.ini 常用规则可迁移，不读取无关配置。
- UI 在应用导入前展示增减规则和参数变化；取消无副作用；保存和导入不包含输出目录或其他个人配置。
- 验收：内置行为一致、旧常用迁移、往返/差异、非法与未来版本、取消不写设置、批次/审计包记录精确配置。

## Task 5: PD-05 字体 provider 能力与复检

交付：统一 provider 接口/能力模型、现有 Acrobat 手动交接适配、UI 能力状态与修复后复检。实际源码和字体探索记录已明确停用不可靠的 COM/OLE 自动修复；保留该产品决定，自动修复标为 blocked，不因安装 Acrobat 推定支持。

- 返回平台、provider/版本、可用性、支持操作与不可用原因；不通过安装存在就推定 COM/Profile 可调用。
- 统一 inspect/capability/repair/recheck 结果，provider 返回成功后仍须重新预检三类字体风险（未嵌入、替代、未知）。
- provider 不可用、超时、修复失败、复检失败分别处理，保留原件/既有输出，禁止自动整页栅格化作为成功兜底。
- 只适配已有 Acrobat 手动交接；自动 repair 返回明确不可用原因，不重新启用未验证 COM fixup，不承诺自行修复所有 CJK。统一自动结果复检契约可用测试 provider 验证，测试 provider 只存在于测试工具中。
- 验收：缺依赖/不支持平台、provider 假成功仍有风险、真实可解析 PDF 复检、异常回收。真实 Acrobat/字体文件单列外部验收。

## Task 6: PD-06 取消、超时和资源约束

交付：文件超时设置、Windows qpdf 子进程树管理、可取消批次状态、保守的并行资源限制。

- 用户可设置每文件超时，支持关闭；计时基于单调时钟。超时/用户取消/引擎失败分开记录并能重试。
- qpdf 进程在启动时纳入受控进程树（优先 Windows Job Object）；取消、Python worker 异常和退出时后代进程收尾，不能按进程名全局杀进程。
- 回收失败保留 staging 并告知；不发布部分输出，既有输出保全。窗口退出仍受线程生命周期守卫保护。
- 并行数量作为用户上限，结合可用内存和输入尺寸保守下调；资源信息不可用时安全回退。低磁盘空间/网络 I/O 错误有可理解反馈。
- 验收：真实受控父/子进程终止、超时、取消竞态、启动失败、成功输出、重试和设置边界；网络设备断连只能依据真实证据标通过。

## Task 7: PD-07 统一规则执行契约

交付：在现有规则目录与管线内统一 inspect → plan → execute → verify → result，不进行无关整体重构。

- 每个规则结果包含 rule_id、matched、planned、changed、verified、skipped_reason、warning；无法观测的值用 null/明确状态，不把“执行过”转换成“已更改/已验证”。
- 预检别名映射、UI 文案与执行选择引用同一稳定规则元数据；保留原有规则顺序、互斥组合、smart 模式和公开入口兼容。
- 各执行阶段报告实际规则事件；独立后置检查支持的规则才设置 verified=true，其他明确 not_verified。
- PD-02 UI、PD-03 审计包和 PD-04 配置消费同一结果，错误/部分成功无歧义。
- 验收：全部现有规则被目录覆盖、复合规则、未命中/不支持预检/无实际变化/失败、事件完整性，以及真实 PDF 行为回归。

## Task 8: PD-08 分层测试与样本回归

交付：unit/integration/qualification 三层可执行入口、明确测试分类清单、样本版本/预期指标及变更说明格式。

- 保留安全全量入口，默认完整运行；按层运行不得把未选测试算通过；分类缺失/重复归属应报错。
- unit 侧无外部进程和 GUI；integration 覆盖 qpdf/PyMuPDF/Qt/发布；qualification 消费 PD-01 明确样本清单及环境条件。
- 资格样本与预期指标版本化；新增/变更样本记录原因、来源授权和人工批准状态，不擅自生成批准记录。
- 汇总展示各层通过/失败/跳过/缺材料，跨层数量不得重复合计。CI 普通回归与显式资格任务分开。
- 验收：层筛选真实运行、分类完整性、失败退出码、缺样本非通过、样本哈希与指标变更可追踪。

## Task 9: PD-09 能力状态与用户文档

交付：`docs/capabilities.md`、README/FAQ/CHANGELOG 与文档首页的一致更新。

- 矩阵列出能力、当前 UI/CLI 入口、稳定/实验/隐藏状态、自动化/成品/真实资格证据、已知限制。
- 对齐安全中心密码与权限说明、页面装配/深度诊断入口、字体 provider、配置/证据包/超时及构建路径。
- 操作步骤从桌面入口开始；底层参数仅出现在开发文档。避免将保留签名外观描述为数字签名仍有效。
- 验收：逐项对照实际路由与交付证据、局部链接检查、无矛盾能力声明。纯文档不新增镜像测试。

## Task 10: PD-10 可追溯发布

交付：依赖版本与哈希锁定流程、构建来源元数据、SBOM/许可证/成品哈希、双变体统一 smoke 验收和可选代码签名入口。

- 记录源码 commit、工作区是否有未提交修改、实际应用/构建依赖版本，不能把本地环境描述为完全隔离可复现；若包含用户主题修改必须可追踪。
- 默认产物新目录，所有构建与自检保持安全清理保护；无需购买或读取签名凭据。
- 锁定安装与校验失败必须中止；SBOM/许可证与实际打包组件对应，不凭固定名单猜测。
- 双变体至少验证启动、真实核心 PDF 管线、qpdf、图片压缩；NoUpdate 无更新入口/请求，产生机器可读报告。
- 代码签名通过显式证书选择/环境引用接入，无证书显示 unsigned/pending；正式签名、外部时间戳和远程发布不隐式执行。
- 发布证据包含能力矩阵、已知限制、构建元数据、哈希、分层测试和资格报告；最终只声明实际运行的检查。

## 执行记录

### 文档基线（2026-11-09）

- 工作区：`C:\Repositories\RATools-for-PDF`，初始分支 main，HEAD `cc6150d`。
- 已有用户改动：`ratools_pdf/ui/theme.py` 与 `.playwright-mcp/`；不得纳入本计划提交。
- 当前基线安全全量 401/401 通过。未重建历史 EXE 或重复业务资格声明。
- Ruling：沿用当前工作区和精确提交；用户已授权此项目连续实现和逐项 commit，避免另建目录丢失本地运行上下文。成本是必须严格守住暂存范围。
- Ruling：PD-01/05/10 外部资格材料缺口不伪造、不默认通过；记录开放项并继续软件交付，最终明确软件/真实验收状态。
- Ruling：PD-02 先定义可序列化的文件级事实，PD-07 再统一规则级实际事件，避免提前把 planned 当 changed。

后续每项记录 RED/GREEN 命令、全量结果、review 结论、提交哈希、用户可用入口和尚未完成的资格验收。

### PD-01：资格矩阵与只读证据运行器（2026-11-09）

- 交付：版本化清单、只读验证服务、CLI、公开合成 CJK PDF、手工阅读器模板。入口见 [资格验收说明](../qualification/README.md)。
- RED：服务不存在的明确断言失败；补充边界暴露缺少预期指标/引擎版本仍通过，修复后专项 17/17（5.140 秒）。
- GREEN：安全全量 418/418（65.569 秒），日志 `backups/product-roadmap-20261109/PD01/full-suite.log`；现存 fitz 弃用提示保留记录。
- 实际 CLI 使用随附 manifest 和 PDF，qpdf 12.3.2 clean，输入前后哈希一致。样本 1 passed / 5 pending；矩阵 1 passed / 8 pending，退出码 2 正确表示未完成全部资格。
- 证据：`backups/product-roadmap-20261109/PD01/evidence-20261109-v1/qualification-report.json`。
- Review：pd01_review，Spec PASS / Quality PASS，无阻断；只读核对路径安全/qpdf 直接依赖，无重复全量。
- 提交主题：`feat: add versioned read-only PDF qualification evidence`；具体哈希见 Git 同主题记录及下一项更新。
- 边界：未提供真实 Acrobat/Word/LibreOffice/扫描仪/打印驱动样本；Windows 成品及外部阅读器未执行。签名仅检查存在性，渲染哈希不证明跨阅读器一致。软件机制完成，外部资格保持开放。
### PD-02：计划预览、输出验证与批次事实（2026-11-09）

- 交付：后台只读计划预览、取消/忙碌守卫、输入与选项配置绑定、发布前结构快照/保存重开验证、批次及重试事实留存。[操作说明](../processing-verification.md)。
- RED：初始 12 项验证与 8 项 UI 验收失败，真实复现引擎报告成功但坏 PDF 覆盖旧输出。review 再复现验证期间停止/跳过仍发布、动态批注合法清理被误判。
- 修复轮：原两 P1 已解决；第二轮补动态批注标准空格/xref 去重和验证失败后 skip 请求不泄漏至下一文件。skip RED 测试初次有夹具 NameError，缺陷复现依据为 reviewer 内存探针；修正夹具后的两文件 GREEN 确认修复，不伪称初始日志有效复现。
- 最终专项 22/22、UI 8/8、worker 21/21、queue 8/8、busy 13/13、dynamic 10/10。安全全量 448/448（80.408 秒），`backups/product-roadmap-20261109/PD02/full-suite-round2-final.log`。首轮 438/1 错误是窗口生命周期测试使用无效 PDF 字节，改为真 PDF 后通过；失败日志保留。
- Review：pd02_review 原两 P1 和后续两项均 ADDRESSED，最终 Spec PASS / Quality PASS；报告证据措辞已修正。
- 提交主题：`feat: preview processing plans and verify outputs before publication`。
- 接口：`processing_evidence`、`processing_evidence_history[batch_id]` 保存实际结构化事件及文件前后快照；PD-03 直接消费，不解析本地化日志。重试单独 batch_id。
- 边界：规则级实际 changed/verified 尚为 unknown/not_verified，PD-07 补齐；签名有效性/真实业务/外部阅读器未验证。生产公开 `(success, message)` 契约保留。
### PD-03：可归档审计证据包（2026-11-09）

- 交付：日志窗口“导出最近批次审计包”、版本化 `evidence.json`/`summary.csv`/`integrity.json`/`COMPLETE.json`、核验 CLI 和[使用说明](../audit-evidence.md)。消费 PD-02 结构化事实，不解析本地化日志。
- RED：初始 17 项、19 个失败子场景；稳定原因码补充另有 5 个子场景 RED。最终专项 20/20（8.049 秒）。
- 全量：468/468、84.470 秒、退出码 0。可核对摘录见 `backups/product-roadmap-20261109/PD03/full-final-evidence.log`；该文件明确是执行会话结果尾部转录摘录，不是完整原始重定向日志，未重跑测试。
- Review：pd03_review 首轮及最终增量均 Spec PASS / Quality PASS，无阻断。runtime 缺失提示已按实际条件修正；`reason_codes` 从结构化 outcome 映射，非成功项标记 `engine_diagnostic:not_recorded`。
- 提交主题：`feat: export verifiable audit evidence packages`。
- 接口：导出最近终态批次；新批次留存应用/PyMuPDF 版本，旧批次、qpdf/build 缺失项保持 unknown；输出相对路径、字段白名单和哈希完整性核验，哈希不提供签名认证。
- 边界：当前无历史批次选择器；不包含 PDF 正文；规则实际变化和签名有效性 not_verified；内部质量系统采纳、真实业务/Acrobat/监管资格仍 pending。
### PD-04：版本化配置与桌面导入导出（2026-11-09）

- 交付：严格 profile schema、内置配置版本、常用旧格式迁移、桌面导入/导出和差异确认、精确规则/参数/身份及计划/审计摘要绑定。[使用说明](../profiles.md)。
- RED：独立模型 6 项、桌面接口 4 项、计划/审计绑定 2 项分别因能力缺失失败。最终 profiles 13/13（3.796 秒）、plan_ui 8/8、audit 20/20。
- 全量：481/481、110.946 秒、退出码 0；真实完整日志 `backups/product-roadmap-20261109/PD04/full-suite.log`。现存 fitz 弃用提示保留。
- Review：pd04_review Spec PASS / Quality PASS；核对 schema、压缩/别名精确映射、取消零修改、规则/参数变化清身份、内置集合与冻结摘要绑定，无阻断。
- 提交主题：`feat: version processing profiles and add desktop import export`。
- 边界：只支持 structural，不能表示的压缩子组合明确拒绝；custom 版本不自动递增但摘要区分内容；重启普通选择为 custom，应用常用后恢复身份。内置机构预设并非官方认证。
### PD-05：字体能力、手动交接与 fresh 复检（2026-11-09）

- 交付：统一 provider/capability/inspect/repair/recheck 契约、Acrobat 手动适配、后台三风险复检与关闭保护。[使用说明](../font-providers.md)。
- 计划校正依据：现有 `acrobat-font-embedding-automation-notes.md` 与 controller 明确 COM/OLE 自动 fixup 已退役。本项保留手动模式，不复活不可靠自动链路。自动 repair 返回 blocked，安装存在不证明 Pro/Profile 可用。
- RED：初始 20 项缺API/入口失败；探测异常边界增加后 24 中 1 失败；review 再实证字体枚举异常被跳过造成假通过，新增真实 PDF 边界回归 RED25中1，最小异常传播修复。
- GREEN：字体专项25/25、直接预检42/42；最终安全全量506/506（128.643秒，exit0），日志 `backups/product-roadmap-20261109/PD05/full-suite.log`。此前505结果另存，不混作修复后证据。
- Review：pd05_review 初轮1项P1，修复后 ADDRESSED，Spec PASS / Quality PASS；检查不可用返回未知计数并关闭文件，不泄露底层私人路径。
- 提交主题：`feat: expose font provider capabilities and verify manual repairs`。
- 边界：手动交接与软件复检完成；自动修复blocked、真实Acrobat/CJK/字体/视觉/监管资格pending。timeout仅分类底层异常，未增加本地解析器硬终止。复检通过仅代表既有三类结构风险未发现。
### PD-06：超时、取消、进程树与资源限额（2026-11-09）

- 交付：每文件超时设置、串并行 deadline、停止/跳过协作取消、Windows Job Object 原子进程归属、资源/低磁盘检查及审计 `file_timeout` 原因。[操作说明](../processing-control.md)。
- RED：初始 13 项中 11 项失败，复现后代残留、时限越过验证仍发布和设置/资源入口缺失；review 还复现了创建与 Job 分配窗口及快照循环未接取消谓词。修复轮 RED 21/4，覆盖 owner 刚返回退出与三阶段取消。
- GREEN：processing 51/51、workers 21/21、qpdf 25/25、busy 13/13；最终安全全量 527/527（156.676 秒，exit0），`backups/product-roadmap-20261109/PD06/full-suite.log`。
- Review：pd06_review fix round 1 两项 P2 ADDRESSED，Spec PASS / Quality PASS。原子 `STARTUPINFOEX` JobList/HandleList 创建后 owner 立即退出无残留；快照、计划哈希、输出快照接收 stop/skip/deadline，单次 OS/fitz 调用不能硬中断的边界已披露。
- 提交主题：`feat: control processing deadlines and owned child processes`。
- 边界：Job 归属与受控故障已实测；网络盘、移动设备、外部受限 Job 宿主和非 Windows owner 崩溃仍 pending。资源上限为保守估计，不等同任意 PDF 内存峰值保证。
### PD-07：统一规则计划、执行与验证结果（2026-11-09）

- 交付：同源规则元数据/预检别名、详细管线API、真实规则修改观察、有限独立后置检查、worker/批次/审计/日志消费，旧tuple入口保持兼容。[字段与能力说明](../rule-results.md)。
- RED：核心8项、消费者2项、打开前失败与纸型优先2项分别复现；review发现复合去色误归属和执行期skip原因被计划空值覆盖，4条回归RED后最小修复。
- GREEN：专项16/16（5.879秒）、audit20/20；最终全量543/543（156.270秒），日志 `backups/product-roadmap-20261109/PD07/full-suite.log`。初轮539/1旧unknown断言失败日志另存。
- Review：pd07_review原P1/P2均ADDRESSED，Spec PASS / Quality PASS。URI-only真实PDF、invalid枚举边界模拟均明确标注，合并保留superseded/pipeline_aborted与冻结smart跳过原因。
- 提交主题：`feat: record structured rule execution and verification results`。
- 边界：标题/纸型/基本URI/批注/附件支持有限独立后置检查；复杂书签/样式/qpdf等无直接观测时changed=null、verified未知，不由选中规则或中文日志推断。实际操作发生不等于最终published；原子发布与取消守卫保留。
### PD-08：分层测试与版本化样本回归（2026-11-10）

- 交付：安全入口分层选择、显式模块分类、准确唯一TestCase计数、样本历史/指标版本/批准记录校验、独立CI资格入口。[使用说明](../testing.md)。
- RED：旧CLI把--layer当pattern，0tests却exit0；版本历史被忽略；setup ErrorHolder摘要崩溃；review复现approval覆盖原路径failed。分别有行为回归，未选模块不导入，零选择/分类错误明确拒绝。
- GREEN：runner8/8、资格机制22/22；unit11/11，integration545/545（201.669秒）；默认全量556/556（182.849秒，exit0），`backups/product-roadmap-20261109/PD08/full-suite.log`。不同执行不重复累计为独立测试总数。
- 资格入口实际exit2：6个样本pending、9行矩阵pending，初始人工批准均pending。来源/批准未伪造，样本与矩阵分列计数。
- Review：pd08_review状态聚合P1 ADDRESSED，SpecQuality PASS；原failed/blocked保留，approval不得升级失败。ErrorHolder单列suite_errors、未执行pending。
- 提交主题：`test: add explicit layers and versioned sample qualification`。
- 边界：CI只本地配置检查，未远程运行；外部样本、阅读器、成品资格与人工批准未完成。
### PD-09：能力矩阵与用户文档一致性（2026-11-10）

- 交付：[能力矩阵](../capabilities.md)、README/FAQ/Unreleased与文档导航，及三份旧操作说明的最小一致性修正。真实入口/稳定状态/软件证据/成品证据/外部资格分列。
- 验收：63处本地链接、0失效；六项当前能力表述检查通过，diff --check通过。不存在的旧SECURITY.md导航已移除；未新建政策或改写旧版本历史。
- Review：pd09_review Spec PASS / Quality PASS，无阻断；核对隐藏工作台无开关、已知密码安全中心、自动字体blocked、有限规则事实、签名外观非密码学有效性及构建边界。
- 提交主题：`docs: align capability matrix and desktop workflows`。
- 范围：纯文档，未重跑代码测试/构建；556历史测试数注明引用。PD10成品证据将在实际构建后刷新，外部资格不自动通过。
### 最终集成审查修复：串行发布前跳过（2026-11-10）

- 整体交付只读审查确认1项P1：最终输入指纹或路径检查阶段，串行UI真实request_skip_current因资格提前清理而被忽略，仍会发布。
- 修复：资格持续至发布决策，成功在发布锁内清理，失败/停止/跳过/超时在终态事件前清理；不把请求泄漏到下一文件。
- RED：控制专项23项中新增2项失败；GREEN23/23（30.705秒）、验证22/22、worker21/21。联合冻结全量570/570（186.759秒）包含此修复，随后仅发布构建工具仍有调整，不能据此声明PD10最终全部通过。
- Review：final_integration_review限定复审Spec PASS / Quality PASS，原P1关闭，无新增阻断；源代码其余跨项契约已核对，PD10独立审查与真实成品仍待最后证据。
- 提交主题：`fix: honor serial skip requests until output publication`。
### PD-10：可追溯发布与双变体成品验收（2026-11-10）

- 交付：14包Windows CPython3.11 AMD64精确wheel锁、哈希/RECORD/实际导入来源校验、受保护fresh-child构建、完整Analysis归属、SBOM/许可证/成品哈希、显式签名入口、统一源码/EXE自检。[构建与成品证据说明](../build-verification.md)。
- 最终专项17/17；完整安全回归575/575（180.985秒），unit11+integration564，qualification未执行。日志 `backups/product-roadmap-20261109/PD10/full-release-final.log`，此前562/570结果保留为独立历史快照。
- 最终严格双构建exit0：`dist/f313fcee487b49a480d389fc1fafc5ca`；14包校验、主/分析子进程清理保护通过。两个EXE均隐藏运行--verify-release，exit0、ok=true，GUI/core PDF/qpdf/images/update boundary五项passed。
- 每变体885个实际Analysis输入完整归属、11个SBOM组件、24个许可证文件；加入验收证据后296文件哈希核验通过，EXE字节未变。最终机器摘要 `backups/product-roadmap-20261109/PD10/release-evidence-summary.json`。
- 源码与产物：构建使用55ceb0c加当时PD10和用户主题未提交字节，逐源文件哈希记录且与源码自检一致；提交后的程序源码不再变更。用户主题未纳入提交，但其字节确实包含于成品并记录来源。
- Review：pd10_review全部已确认完整性/来源/签名状态问题关闭，最终代码增量及成品证据/文档SpecQuality PASS；final_integration_review跨项唯一串行skip问题已独立55ceb0c修复并复审PASS。
- 验收过程中真实发现并修复：RECORD缺失放行、ambient构建导入、未知Analysis静默遗漏、签名状态/目标链接门禁、隐藏子进程日志、Pythonwin锁定来源、生成归档共享引用与精确常量类型比较、第二变体Analysis编号。失败日志和旧候选产物保留，不冒充最终成功。
- 提交主题：`build: produce traceable locked releases with verified variants`。
- 边界：本次全部unsigned；没有正式签名、时间戳、远程推送/发布。Smoke为合成软件检查，不替代全GUI/逐规则或真实业务/阅读器/Acrobat/监管资格；Python socket观察不代表所有OS网络路径。锁仅当前平台，不宣称位级可复现。

## 本轮交付结论（2026-11-10）

十项计划均有软件机制或文档交付、对应验收与独立只读review；各项独立提交。PD-05按当前产品决定完成手动交接/复检，自动修复仍blocked；真实样本与人工批准、Acrobat/字体/阅读器、设备与监管资格、正式代码签名继续作为明确开放项，不能把本轮软件完成理解为全部业务资格完成。

最终可运行目录：标准版 `dist/f313fcee487b49a480d389fc1fafc5ca/RATools-for-PDF/`；无更新版 `dist/f313fcee487b49a480d389fc1fafc5ca/RATools-for-PDF-NoUpdate/`。分发须保留完整目录及_internal/release，不单独复制EXE。

实现期间保留用户 `ratools_pdf/ui/theme.py` 和 `.playwright-mcp/`；备份、失败/成功日志及资格缺口证据保留在 `backups/product-roadmap-20261109/`。未远程推送。