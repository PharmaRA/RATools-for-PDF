# 质量改进执行计划

目标：先记录发现，再按优先级逐项修复/优化；每项验收和只读 review 通过后单独 commit。

审查来源：[2026-11-08 代码审查](../reviews/2026-11-08-code-review.md)。基线 `ad95c8c`。原始编号 R-01 至 R-18 不变，下表按执行顺序排列。

## 执行约束

1. 保留用户未提交的 theme.py 和 .playwright-mcp/，精确指定暂存文件。
2. 重要文件修改前复制到项目根 backups/，不提交备份。
3. 不物理删除任何文件；临时产物及测试清理进入回收站；处理链接前先审查。
4. 每项先有真实失败复现，再做最小修复；通过对应验收和只读子代理 review 后提交。
5. 新发现或无法复现项记录证据，不强行改动。未执行的真实集成/发布验收明确保留为未完成。
6. 提交使用 Fiveonine 身份。每个独立修复提交包含相关测试和本计划状态更新。

## 状态

| 顺序 | 发现 | 优先级 | 修复范围 | 状态 | 验收/提交 |
|---|---|---|---|---|---|
| 01 | R-01 | P1 | 临时输出所有权、成功提交、失败回收 | 已完成 | aa1ee0e；14/14；全量 252/252；review 通过 |
| 02 | R-08 | P1 | 普通批次/IO/拆分/解密输出冲突 | 已完成 | 35ff9d0；14/14；全量 266/266；review 通过 |
| 03 | R-02 | P1 | 链接改色内容流 | 已完成 | e82c5af；11/11；全量 277/277；review 通过 |
| 04 | R-03 | P1 | 旋转页纸张适配 | 已完成 | ac3bc4e；9/9；全量 286/286；review 通过 |
| 05 | R-04 | P1 | 缩放后导航恢复 | 已完成 | 4b584ea；10/10；全量 296/296；review 通过 |
| 06 | R-05 | P1 | 图像透明信息 | 已完成 | 17350c7；6/6；原压缩 11/11；全量 302/302；review 通过 |
| 07 | R-06 | P1 | 动态内容精确清理 | 已完成 | 785f5b3；10/10；全量 312/312；review 通过 |
| 08 | R-07 | P1 | 签名烘焙失败保护 | 已完成 | 3684dbd；11/11；原签名 5/5；全量 323/323；review 通过 |
| 09 | R-09 | P1 | 队列去重与任务身份 | 已完成 | e02b45c；8/8；全量 331/331；review 通过 |
| 10 | R-10 | P1 | 并行异常收尾，串行 IPC 验证 | 已完成 | d3af2c0；7/7；直接回归 22/22；全量 338/338；review 通过 |
| 11 | R-11 | P1 | IO 忙碌互斥 | 已完成 | 2cec9ab；13/13；全量 351/351；二轮 review 通过 |
| 12 | R-12 | P1 | 窗口关闭和应用退出收尾 | 已完成 | 75e57df；5/5；全量 356/356；review 通过 |
| 13 | R-13 | P2 | 命名书签 CSV 往返 | 已完成 | 60957f8；9/9；既有往返 7/7；全量 365/365；二轮 review 通过 |
| 14 | R-14 | P2 | 过期预检拦截 | 已完成 | 7a9903e；6/6；全量 371/371；review 通过 |
| 15 | R-15 | P2 | 诊断任务结果归属 | 已完成 | a565e7b；9/9；旧诊断 5/5；全量 380/380；review 通过 |
| 16 | R-16 | P2 | 解密准备异常反馈 | 已完成 | 7/7；全量 387/387；review 通过 |
| 17 | R-17 | P2 | Pillow 打包及能力反馈 | 待实施 | — |
| 18 | R-18 | P2 | 按真实尺寸计算压缩 DPI | 待实施 | — |

## 验证与推进记录

### 文档阶段

- 已记录全部 18 项发现、原始证据、修复方向和验收标准。
- 全量测试需通过不执行物理删除的安全运行方式，不能直接运行现有自动清理路径。
- 完成定义：18 项均有当前代码证据、针对性回归、必要集成验收、只读 review 和独立提交；未通过或未执行的要求不得标为完成。

### R-01：输出保全（2026-11-08）

- RED：初始 7 个针对性测试中 6 个按预期失败，实际捕获旧输出被清理、直接写正式结果、别名未拦截。
- 修复：每任务独立同盘 staging；子进程仅写 staging/result.pdf；成功后原子发布；失败/停止/跳过仅回收本轮目录。回收失败保留目录并在结果中提示。写入和回收前检查符号链接、重解析点和硬链接。
- GREEN：`python -X utf8 -B scripts/run_tests_safely.py test_worker_output_safety.py`：14/14 通过。含真实串行/并行 PDF 子进程、真实 qpdf 中间文件清理、成功与失败覆盖、启动失败、回收失败。
- 全量：`python -X utf8 -B scripts/run_tests_safely.py`：252/252 通过，37.105 秒。日志 `backups/verification/R01-full.log`（本机备份，不提交）。仅有上游 fitz API 弃用提示。
- Review：`controllers_review` 首轮确认主体通过，指出安全 runner 的 spawn 子进程未安装保护；已修复为测试脚本 `__mp_main__` 引导时安装，并以子进程审计探针验证。第二轮明确通过，无阻断。
- 提交主题：`fix: preserve published outputs when processing fails`。
- 边界：未实测没有回收站的设备、网络盘或异常系统回收策略；回收异常保留文件，不回退物理删除。其他 17 项仍待修复。

### 安全测试入口

Windows 本机使用 `python -X utf8 -B scripts/run_tests_safely.py [test_*.py]`。测试临时目录限定在 `backups/test-runs/`，清理进入系统回收站，Python 审计钩子拦截物理删除；Windows spawn 子进程安装相同保护。该运行器目前不支持非 Windows，不能据此声称其他平台通过。

### R-08：整批输出冲突（2026-11-08）

- RED：初始 9 个测试全部失败；真实 qpdf 拆分和解密的同名双输入均只留下一个输出。
- 修复：共享来源目录规划和写前校验。普通批次、IO、拆分、解密保留相对来源目录；跨盘时保留卷标识及来源路径。规范化后的重复最终路径、输出占用任一排队输入在首次写入前拒绝；明确启用原名解密只允许覆盖任务自己的输入。解密使用唯一 staging，避免固定临时名覆盖其他输入。
- 移除仅凭 basename 的 eCTD 拦截，改为比较最终路径，允许不同子目录保留同名结果。
- GREEN：`python -X utf8 -B scripts/run_tests_safely.py test_output_collisions.py`：14/14 通过。含真实双来源拆分/解密内容保全、IO 导出导回、正常 eCTD 子目录输出、原名解密及冲突前零引擎写入。
- Review：`controllers_review` 首轮指出拆分 `%d` 模板未审核实际分卷链接；补充已有分卷的限定枚举及逐个链接审核，新增测试 RED→GREEN，二轮明确通过。
- 全量：`python -X utf8 -B scripts/run_tests_safely.py`：266/266 通过，41.166 秒；日志 `backups/verification/R08-reviewed-full.log`。仅有上游 fitz API 弃用提示。
- 提交主题：`fix: prevent batch output collisions across processing tools`。
- 兼容性：跨盘 IO 数据及多目录拆分/解密采用保留来源层级的新路径；旧版平铺导出的数据需按新预览路径放置或重新导出。未在真实多卷/网络共享上写入，跨盘与大小写规则由 Windows 路径测试覆盖。

### R-02：链接改色正文保全（2026-11-08）

- RED：初始 7 个真实内存 PDF 测试全部失败，覆盖 Tj/TJ、继承颜色、括号转义、正文/注释中的 BT/ET/rg 和后续正文状态。
- 修复：识别完整 PDF 操作边界；字符串、数组、十六进制文字和注释不作为指令。颜色指令写在完整显示操作之前，随后恢复原填充颜色/颜色空间；不重写文字操作数，不重置字体、字间距或文本矩阵。先完成整页解析再写回，避免部分修改。
- GREEN：`python -X utf8 -B scripts/run_tests_safely.py test_hyperlink_content_safety.py`：11/11 通过；检查文字与坐标、重新打开后蓝色像素、后续正文颜色、跨流继承颜色，以及 process_document 实际保存后的文字和 URI 链接。
- Review：`pdf_core_review` 明确通过，并独立补验 CMYK、cs/sc、q/Q 颜色恢复及重复内容流引用；无阻断。
- 全量：`python -X utf8 -B scripts/run_tests_safely.py`：277/277 通过，41.314 秒；日志 `backups/verification/R02-reviewed-full.log`。仅有上游 fitz API 弃用提示。
- 提交主题：`fix: preserve PDF text and color state when styling links`。
- 安全边界：内联图像、跨流文本对象或源码文本块与渲染跨度无法可靠对应时，不猜测并改写原流；这类复杂页面的改色覆盖能力未作全面资格声明。文字安全与本项明确验收已通过。

### R-03：旋转页纸张适配（2026-11-09）

- RED：初始四组测试出现八个失败子场景，复现 90/270 度方向及裁切、批注/内联链接错位、跨内容流字体状态丢失。
- 修复：保留 Rotate，按照可见裁剪框在原始 PDF 坐标中缩放、居中和裁剪；为整个 Contents 添加独立前后流，避免修改共享内容流及逐流重置状态。按原始 PDF 几何统一调整批注、QuadPoints、表单与链接区域，保留动作和外观。
- Review：`pdf_core_review` 首轮发现共享 BS 字典被重复缩放；已在修改前复制字典，三引用回归 RED→GREEN，二轮明确通过。
- 全量初轮：284 项中三个页面适配用例发生回收 COM 错误。新增禁用循环 GC 的 Windows 独占打开测试，确认递归闭包保留底层 PDF 句柄；将递归助手移至顶层后 RED→GREEN，源文件可立即释放。原失败日志为 `backups/verification/R03-full.log`。
- GREEN：`python -X utf8 -B scripts/run_tests_safely.py test_page_resize_safety.py`：9/9 通过，含四种旋转角度的文字坐标和像素、批注/高亮/表单/URI 渲染、CropBox、内联链接、共享流/样式以及实际 A4 管线输出。
- 最终全量：`python -X utf8 -B scripts/run_tests_safely.py`：286/286 通过，43.221 秒；日志 `backups/verification/R03-reviewed-full.log`。仅有上游 fitz API 弃用提示。
- 提交主题：`fix: preserve rotated page content during paper normalization`。
- 边界：书签及跨页导航的目标坐标仍按 R-04 独立修复；本项完成页面可见内容和页面内对象的适配。

### R-04：缩放后的本地导航（2026-11-09）

- RED：初始六组测试出现七个失败子场景，覆盖书签变为 (40,1000)、不缩放源页跨页链接变为 (20,500)、旋转目标、共享命名目标以及打开动作。原两个场景正确结果均为 (40,200)。
- 修复：修改页面几何前，按目标页原始 PDF 坐标计算本地目标；复制目标数组后统一写回。覆盖 Outlines、页面链接（含内联）、Catalog/OpenAction、AA/Next、Names/Dests 名称树和旧 Dests 字典，支持 XYZ/FitH/FitV/FitBH/FitBV/FitR；保留 null、固定缩放、名称、书签层级和动作参数。远程 GoToR 目标不变。
- Review：`pdf_core_review` 首轮指出 OpenAction 动作字典被通用目标入口误认，直接/间接 GoToR 的新增回归 RED→GREEN；现按数组/动作类型分派，二轮明确通过。
- GREEN：`python -X utf8 -B scripts/run_tests_safely.py test_page_navigation_resize.py`：10/10 通过。包括两个页面顺序、共享数组仅缩放一次、远程共享数组不污染、循环动作终止、内联点击区域、原名命名目标及实际 Letter 输出后的书签层级/坐标/缩放。
- 最终全量：296/296 通过，44.728 秒；命令 `python -X utf8 -B scripts/run_tests_safely.py`，日志 `backups/verification/R04-reviewed-full.log`。仅有上游 fitz API 弃用提示。
- 提交主题：`fix: transform local navigation targets before resizing pages`。

### R-05：压缩透明图像保全（2026-11-09）

- RED：六个专项测试全部失败，半透明、颜色键、Matte 和显式遮罩的平均像素差约 84，原共享 SMask 被清空。
- 修复：从 PDF 解码像素读取颜色（应用 Decode、颜色键、Matte 解混），合并软透明度后降采样；颜色 JPEG 与无损 Flate SMask 分开写入，新遮罩独立 xref，体积比较包含新增遮罩。显式二值 Mask 保留原对象和分辨率，避免细透明条纹二次采样；无白底铺平。
- GREEN：`python -X utf8 -B scripts/run_tests_safely.py test_image_transparency.py`：6/6 通过，要求实际降采样且保存重开后平均通道差小于 3。`test_image_compression.py` 原 11 项通过；含主处理管线输出、共享遮罩保全。
- Review：`pdf_core_review` 明确通过，独立五项内存回归及反向 Decode、JPX SMaskInData=1/2 均通过，补充场景像素差 0 至 1.0，无阻断。
- 全量：302/302 通过，43.151 秒；命令 `python -X utf8 -B scripts/run_tests_safely.py`，日志 `backups/verification/R05-full.log`。仅有上游 fitz API 弃用提示。
- 提交主题：`fix: retain image transparency during PDF compression`。
- 边界：DPI 根据实际绘制尺寸计算属于 R-18，仍待实施；本项没有把该能力标记完成。

### R-06：动态入口精确清理与预检（2026-11-09）

- RED：六组测试全部未通过，复现附件丢失、动态批注及 OpenAction/AA 脚本残留、智能预检遗漏和无修改时仍假报变更。
- 修复：共用只读清理计划；仅移除 Names/JavaScript 分支，保留 Dests/EmbeddedFiles。处理 Catalog/OpenAction、Catalog/页面/批注/字段 AA、A/Next、书签动作和嵌套字段；移除 3D/RichMedia/Movie/Sound 批注及相应动态动作。复制动作链以保留安全后继顺序，避免修改共享原对象；循环引用有限终止，超深链明确报错。
- GREEN：`python -X utf8 -B scripts/run_tests_safely.py test_dynamic_content_safety.py`：10/10 通过，含预检输入字节/xref 不变、命名书签和附件保留、普通评论/表单值、安全动作顺序、共享循环脚本、智能 OpenAction 及富媒体实际输出。富媒体合成样例补齐合法静态外观流后无相关 MuPDF 警告。
- Review：`pdf_core_review` 明确通过，独立八项内存回归及嵌套共享动作顺序验证通过，无阻断。
- 全量：312/312 通过，43.695 秒；命令 `python -X utf8 -B scripts/run_tests_safely.py`，日志 `backups/verification/R06-full.log`。随后对增强的富媒体保存断言重跑本项 10/10 通过，生产实现不变。仅有上游 fitz API 弃用提示。
- 提交主题：`fix: remove dynamic PDF entries without losing navigation or attachments`。
- 范围：支持上述 PDF 结构入口；XFA/XML 包内脚本未实现解析清理，未删除整个 XFA 以避免丢失其表单布局。本项不作任意 PDF 主动内容的全面安全认证。

### R-07：签名烘焙失败保护（2026-11-09）

- RED：初始七项验收发现渲染/插图/控件移除失败仍保存、qpdf 吞异常继续解除限制，以及原位处理源文件锁未释放；原测试五个失败及四个错误（含清理错误）。
- 修复：独立字节文档避免 Widget/Annot 延迟释放锁住源文件；先成功写入外观，再删除控件并计数；任一控件失败中止整份文档。所有外观完成后保存到独立同盘目录，成功后原子替换。qpdf 重写及解密遇到烘焙异常停止，不再回退原文件继续；中间目录移入回收站。
- 备用渲染使用未旋转坐标裁剪并恢复原旋转角。空矩形/实际不可见外观与渲染错误区分，不因渲染失败判为成功。
- GREEN：`python -X utf8 -B scripts/run_tests_safely.py test_signature_failure_safety.py`：11/11；原 `test_signature_baker.py`：5/5。包含第二控件失败、部分保存失败、成功/失败原位处理、隐藏/零矩形、旋转备用渲染和两个 qpdf 入口。旋转像素对照使用与烘焙相同 DPI，避免重复重采样误差。
- Review：`pdf_core_review` 明确通过，独立验证渲染/插图/删除/保存/替换故障均不发布不安全结果，qpdf 不再继续，已分配工作目录进入回收路径。
- 全量：323/323 通过，46.999 秒；命令 `python -X utf8 -B scripts/run_tests_safely.py`，日志 `backups/verification/R07-full.log`。仅有上游 fitz API 弃用提示。
- 提交主题：`fix: stop signature processing when appearance baking fails`。
- 限制：字节工作副本需要额外约一个输入 PDF 大小的内存；本项未扩展既有签名探测覆盖范围，也未执行真实 Acrobat/签名证书资格验证。

### R-09：队列与运行任务身份（2026-11-09）

- RED：初始六项测试全部失败，复现文件与父目录重复、相对/绝对及大小写别名、外部列表修改影响运行队列、并行任务被覆盖和按路径跳过遗漏。
- 修复：逐项按绝对路径及 Windows 大小写规则去重，保留首次绝对路径；worker 持有队列快照。并行任务按独立索引登记，路径仅用于匹配跳过请求，完成/跳过/整批停止均逐个收尾。
- GREEN：初始六项 RED→GREEN，再增加不同目录同名文件和整批停止场景，共 8/8 通过。公共输出规划仍提前拒绝重复输出；低层调度回归直接调用调度器，验证调用方变化也不会静默遗失已启动任务。
- Review：`controllers_review` 明确通过，独立内存探针验证同路径及大小写别名在完成、跳过、停止三条路径均 join 全部进程并关闭各自连接，无阻断。
- 全量：`python -X utf8 -B scripts/run_tests_safely.py`：331/331 通过，53.174 秒；日志 `backups/verification/R09-full.log`。仅有上游 fitz API 弃用提示。
- 提交主题：`fix: deduplicate queued files and track parallel jobs independently`。
- 边界：调度异常后的统一收尾及大消息 IPC 顺序属于 R-10，仍待实施。

### R-10：异常进程收尾与管道接收（2026-11-09）

- RED：初始七项回归中六项失败，复现启动/通信异常遗留任务、缺少可重试终态及串行大消息 WinError 109；另一个测试夹具重复目录错误已改为独立临时目录。
- 修复：启动前保留失败任务身份；串行/并行 finally 收尾已登记进程、连接及 staging，再向 UI 发错误。完成任务立即移出注册表，后续异常不会重复污染成功记录。串行先完整 recv，再 join，允许发送大消息的子进程正常结束。
- GREEN：`test_worker_lifecycle.py` 7/7；既有输出保全 14/14、队列身份 8/8。覆盖第二项启动、poll/recv、发布异常、串行状态恢复；既有停止/跳过回归通过。真实 Windows spawn/Pipe 传输约 1.4 MB 结果，完整消息、退出码 0、实际 PDF 输出均通过。
- Review：`controllers_review` 明确通过，独立纯内存探针补验上述异常及停止时首次 join 抛错，全部已登记任务仍收尾；确认 recv→join 顺序，无阻断。
- 全量：`python -X utf8 -B scripts/run_tests_safely.py`：338/338 通过，53.024 秒；日志 `backups/verification/R10-full.log`。仅有上游 fitz API 弃用提示。
- 提交主题：`fix: reap processing jobs on scheduler failures and drain results before join`。
- 边界：大消息测试验证真实 IPC 机制，不代表已证明生产日志一定达到该大小；本项管理已登记的 Python 子进程，未实现或验证 qpdf 后代进程树终止。

### R-11：IO 与后台任务互斥（2026-11-09）

- RED：初始五组测试全部未通过，复现其他任务运行时仍创建 IO、IO 运行时队列可操作、结果先于线程退出时按钮提前启用及启动异常逃逸。
- 修复：主处理/IO/预检/检测共用忙碌守卫；IO 向导及执行入口均受保护，队列增删清空在 IO 期间受保护。视图保留任务忙碌状态，普通刷新不会提前启用入口；控制器保留线程引用至 QThread.finished 后再释放。主处理原有运行中追加行为保留，运行队列使用 R-09 快照。
- 准备与启动异常经统一错误反馈和状态恢复；IO 目录选择后重检忙碌。预检/检测的线程启动失败补充 RED→GREEN。
- Review：首轮指出主处理模态提示后缺少第二次守卫；新增真实方法回归复现后，在全部模态交互之后、共享状态赋值及创建 worker 之前重检，RED→GREEN。`controllers_review` 二轮明确通过。
- GREEN：`test_io_busy.py` 13/13；含 IO/主处理/预检/检测的受控真实 QThread 成功和失败、结果与退出分离、重复启动、队列保护、目录准备和线程启动失败、双向目录选择重入。主处理回归使用实际 QObject 接收槽，避免非 Qt 夹具的连接限制。
- 初轮全量 348/348；最终全量：`python -X utf8 -B scripts/run_tests_safely.py`：351/351 通过，55.205 秒；日志 `backups/verification/R11-reviewed-full.log`。仅有上游 fitz API 弃用提示。
- 提交主题：`fix: serialize IO and processing tasks until worker threads finish`。

### R-12：关闭与退出线程收尾（2026-11-09）

- RED：四个工作台的顶部关闭、Escape、系统关闭、accept/done 共 20 个子场景均在工作线程运行时隐藏窗口；主窗口统一生命周期接口尚不存在。
- 修复：对话框基类统一拦截所有关闭路径，检查全部子 QThread 的 wait(0)，包括被最新引用遮蔽的旧诊断线程。主窗口关闭前经控制器检查主处理、IO、预检、检测、更新及全部工作台子线程，未结束则保留窗口并提示等待。
- 应用 aboutToQuit 统一等待所有受管理线程结束后退出，不使用超时后销毁运行中线程的路径。任务自然完成写入和子进程收尾；正常点击关闭被拦截后需在任务结束后重新关闭。直接退出等待期间可能持续至当前任务结束，本项没有实现中途强制取消。
- GREEN：`test_window_lifecycle.py` 5/5，含上述 20 子场景、五类主控 worker、旧诊断线程、六线程退出等待及真实 Windows spawn PDF 任务；关闭时子进程保持运行，任务完成后子进程退出且有效 PDF 已发布，再允许关闭。测试绕过用户 QSettings，不读写真实设置。
- Review：`ui_review` 明确通过，独立内存复验四项（含 20 个关闭场景），并验证真实 QApplication.exec→app.quit 会等待线程且设置退出状态；未发现线程依赖 GUI 回调造成死锁。真实 PDF 回归由主代理安全运行器验收。
- 全量：`python -X utf8 -B scripts/run_tests_safely.py`：356/356 通过，53.763 秒；日志 `backups/verification/R12-full.log`。仅有上游 fitz API 弃用提示。
- 提交主题：`fix: guard window closure until background work has finished`。

### R-13：书签 CSV 命名目标与文件动作（2026-11-09）

- RED：初始六项测试中五项未通过，复现本地命名跳转变为空动作、远程目标丢失、不支持动作/CSV 类型被静默接受；旧三列格式原本通过。
- 修复：新增 NamedDest 字段；先创建层级，再写回标准 GoTo/GoToR 命名动作。按原始 PDF 动作读取远程文件、命名/XYZ 目标和窗口标志；窗口未指定/true/false 三态保留。输入解析失败报告 CSV 行号，未知动作明确拒绝，导出验证完成后才覆盖 CSV；文档通过上下文管理器关闭。
- PyMuPDF 边界：远程目标页码超出当前文档页数可让 get_toc 返回空列表，且远程命名目标/Launch 类型会误判。导出先保存原始文件动作，仅在未保存的内存文档去除动作后遍历标题层级，再合并原始数据；验证输入 PDF 字节不变。远程写回避开 set_toc 将命名字符串当 Point 的限制。
- Review 首轮确认 Launch 被高层 API 误报 GoToR；新增回归 RED→GREEN，现按原始 /A/S 保留 Launch 类型及窗口标志。
- GREEN：`test_bookmark_csv_targets.py` 9/9；原 `test_pdf_processor_roundtrip.py` 7/7。包括本地命名层级/解析目标、转义字符、远程命名与 XYZ、Launch、窗口三态、源字节保全、旧三列兼容及错误输入不覆盖已有产物。
- Review：`pdf_core_review` 二轮明确通过，独立内存 PDF 往返复验 Launch 的窗口三态，原阻断解除。
- 全量：`python -X utf8 -B scripts/run_tests_safely.py`：365/365 通过，56.537 秒；日志 `backups/verification/R13-full.log`。仅有上游 fitz API 弃用提示。
- 提交主题：`fix: preserve named bookmark targets and file actions in CSV round trips`。
- 兼容性：新导出多一个 NamedDest 列；旧三列及原扩展格式仍支持。不能表示的远程目标（例如 Fit 模式）明确报错，避免静默生成错误跳转。

### R-14：过期预检建议拦截（2026-11-09）

- RED：初始六项中五项未通过（处理成功/失败两个子场景），确认新增/移除队列项、处理后、部分预检失败及忙碌状态仍可执行旧建议；重新预检的正常路径原已通过。
- 修复：mark_stale 清空可执行建议文件集合及两个按钮的建议属性，禁用并隐藏建议入口。应用规则和建议文件处理均检查统一 idle 与 current；主处理成功/失败、预检失败均使结果失效。原始结果行保留用于查看，不能据其执行建议。
- GREEN：`test_precheck_staleness.py` 6/6，调用真实队列添加/移除及处理回调，验证规则不被勾选、不持久化设置、不启动处理；重新成功预检后恢复规则应用和建议文件处理。
- Review：`controllers_review` 明确通过，确认失效逻辑、直接执行入口保护、异常路径和重新预检恢复，无阻断。
- 全量：`python -X utf8 -B scripts/run_tests_safely.py`：371/371 通过，56.718 秒；日志 `backups/verification/R14-full.log`。仅有上游 fitz API 弃用提示。
- 提交主题：`fix: reject stale precheck suggestions before applying or processing`。

### R-15：诊断结果与导出 AST 归属（2026-11-09）

- RED：初始七项全部未通过，复现不同文件及 A→B→A 的旧结果覆盖、路径编辑后旧 AST 仍可导出、错误路径 payload 被接受、失败/启动失败和旧线程释放问题。
- 修复：路径变化与每次诊断递增代次并清空旧展示/AST；结果必须来自当前 worker、匹配代次及规范化后的显示路径。失败不发布部分 AST。各线程在真实 finished 时按 sender 释放，旧任务不会清空新任务引用；启动异常恢复重试。
- 导出保存对话框返回后再次核对代次，避免选择目录期间切换文件后写出错误内容；新增失败复现后修复。
- GREEN：`test_diagnosis_identity.py` 9/9，包括真实受控 QThread 的 B→A 完成与两线程销毁、当前引用保全、错误路径、失败重试、导出重入及路径编辑失效。运行中的旧线程继续受 R-12 的父子关系关闭保护管理。
- 全量前两轮在主题测试恢复全局样式表时发生原生 0xC0000005（`R15-full.log`、`R15-faulthandler-full.log`）；诊断+主题窄组合 18/18 通过。新增诊断测试夹具改为明确销毁自身对话框，并将 DeferredDelete 刷新限定到自身线程，排除测试对象延迟销毁和全局清理耦合；生产主题代码未修改，最终全量正在复验。
- 上述夹具改动未单独解决崩溃；完整隔离保留生产代码，排除新增诊断测试 371/371、只排除旧诊断窗口初始化用例 379/379 通过。已确认旧初始化用例自动启动真实诊断，却未等待线程和排队结果就离开临时输入目录；R-12 的关闭守卫会拒绝其提前关闭。现该用例等待真实完成、处理结果、断言 AST 后定向销毁窗口，专项 5/5 通过，完整套件复验中。未把 GC、销毁信号参数或主题代码认定为未经证实的根因。
- Review：`ui_review` 实现及新增夹具 review 通过，独立内存复验 9/9；旧诊断测试补齐 reviewer 建议的 isValid 检查，已销毁线程包装不会阻断对话框清理，专项再验 5/5。
- 最终全量：`python -X utf8 -X faulthandler -B scripts/run_tests_safely.py`：380/380 通过，60.648 秒；日志 `backups/verification/R15-lifecycle-full.log`。最后仅新增测试清理有效性检查后专项 5/5 通过，生产实现未变。临时隔离运行器已移入回收站；排查日志保留。
- 提交主题：`fix: bind diagnosis results and AST exports to the current request`。

### R-16：解密准备失败与界面恢复（2026-11-09）

- RED：初始六项全部未通过，复现目录权限/设备错误和普通文件占用路径时异常逃逸，真实线程失败后不恢复、无摘要退出和启动失败无兜底。
- 修复：输出规划和根目录创建统一纳入准备异常处理，逐文件报告失败并发一次批次摘要。UI 保存摘要到真实 QThread.finished 后再释放引用、恢复执行/关闭按钮；无摘要及启动异常均明确报告失败，支持重试。全部失败时不显示打开结果目录，提示用户查看具体失败原因。
- GREEN：`test_decrypt_failure_recovery.py` 7/7，覆盖模拟权限拒绝/断开设备、真实普通文件路径保全、真实受控错误线程两次重试、无摘要退出、启动异常及真实 qpdf 成功输出内容。
- Review：`ui_review` 明确通过，独立七项内存验证覆盖目录准备、结果先于线程结束、重复启动、无摘要、构造/启动异常及恢复，无阻断。
- 全量：`python -X utf8 -X faulthandler -B scripts/run_tests_safely.py`：387/387 通过，59.778 秒；日志 `backups/verification/R16-full.log`。
- 提交主题：`fix: report decryption preparation failures and restore dialog controls`。
