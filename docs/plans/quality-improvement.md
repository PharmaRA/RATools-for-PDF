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
| 06 | R-05 | P1 | 图像透明信息 | 已完成 | 6/6；原压缩 11/11；全量 302/302；review 通过 |
| 07 | R-06 | P1 | 动态内容精确清理 | 待实施 | — |
| 08 | R-07 | P1 | 签名烘焙失败保护 | 待实施 | — |
| 09 | R-09 | P1 | 队列去重与任务身份 | 待实施 | — |
| 10 | R-10 | P1 | 并行异常收尾，串行 IPC 验证 | 待实施 | — |
| 11 | R-11 | P1 | IO 忙碌互斥 | 待实施 | — |
| 12 | R-12 | P1 | 窗口关闭和应用退出收尾 | 待实施 | — |
| 13 | R-13 | P2 | 命名书签 CSV 往返 | 待实施 | — |
| 14 | R-14 | P2 | 过期预检拦截 | 待实施 | — |
| 15 | R-15 | P2 | 诊断任务结果归属 | 待实施 | — |
| 16 | R-16 | P2 | 解密准备异常反馈 | 待实施 | — |
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
