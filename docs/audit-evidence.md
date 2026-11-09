# 可归档审计证据包

在主窗口打开“处理日志记录”，点击“导出最近批次审计包”。选择保存位置后，程序会在其中创建独立的 `RATools_Audit_日期时间_随机ID` 新目录。导出范围是主窗口内存中**最近已结束的批次**，与日志表格的筛选、旧 CSV 行或显示文本无关。当前版本不提供历史批次选择器；重试会成为新的最近批次。

失败、部分失败、停止和预览取消的批次都可以导出。运行中的批次会拒绝导出。原生目录对话框打开前先复制证据，返回后再次检查批次身份；期间换了批次会要求重新导出。取消目录选择不会写入文件。导出只读取结构化证据，不重新打开输入/输出 PDF。

每个包包含：

| 文件 | 内容 |
| --- | --- |
| `evidence.json` | `pd03.audit.v1` 主清单，包含批次、配置、版本、规则计划、文件前后快照、终态和结构化事件 |
| `summary.csv` | UTF-8 BOM 文件汇总；CSV 列与 JSON 文件事实可以交叉检查 |
| `integrity.json` | `pd03.integrity.v1`，逐一记录 JSON/CSV 的 SHA-256 和字节数 |
| `COMPLETE.json` | `pd03.complete.v1`，最后写入，绑定批次身份与完整性清单 SHA-256 |

完整标识是**包写入完成**，不表示批次处理成功或规则已得到资格验证。创建目标时不允许目录已存在。写入失败时不删除部分目录，也不覆盖原包；失败目录可能只有一部分文件或不完整 JSON，核验必定拒绝。不要把失败目录重命名成成功包或手工补完整标识。清理失败目录请使用系统回收站。

## 核验入口

在安装了项目依赖的项目根目录运行：

```powershell
python -B -m ratools_pdf.services.audit_evidence verify ".\RATools_Audit_20261109_120000_example"
```

也可以在 Python 中调用：

```python
from ratools_pdf.services.audit_evidence import verify_evidence_package
result = verify_evidence_package(package_directory)
```

命令输出 `pd03.verify.v1` JSON。`valid: true`、退出码 `0` 表示当前包的完整性和受支持 schema 检查通过；`valid: false`、退出码 `1` 表示失败。核验会拒绝缺失、额外、被篡改或类型错误的载荷、不支持的 schema、越界相对路径、链接/重解析点/硬链接，以及 CSV 与 JSON 事实不一致的包。包内只允许上表四个普通文件。

SHA-256 清单和完整标识**不是数字签名或来源认证**。能修改全部文件并重新计算哈希的人可以重新制作一个自洽包。需要来源认证时应使用另行管理的签名/可信归档系统；本项没有实现签名。包不包含 PDF 载荷，因此核验包不等于重新核验磁盘上的 PDF，也不证明 PDF 原件仍未改变。

## 字段与事实边界

源证据采用 `pd02.batch.v1`、预检计划 `pd02.plan.v1`、前后结构快照 `pd02.snapshot.v1`、处理后验证 `pd02.verification.v1`。配置使用 `pd02.engine.v1`。未知版本或缺失必须 schema 的输入会拒绝导出。包保留配置摘要、稳定 `option_id`、参数和计划理由。

文件用 `file_id` 和输入根/输出根下的相对路径表示。同名文件在不同子目录时保持不同身份。输入哈希来自批次快照；已发布输出的哈希来自发布前验证快照，相对身份来自 Worker 实际输出位置。没有发布输出时 `output` 为 `null`，缺失事实在 `missing_evidence` 中明确列出。较早的内存批次若没有输出相对身份，会保留哈希并报告身份缺失，导出层不会猜测输出文件名。

`processing_runtime` 复制批次建立时实际留存的应用和 PyMuPDF 版本；升级前建立、没有版本留存的历史批次保留 `unknown`。qpdf 执行版本及应用构建提交未在 PD02 留存，明确为 `unknown`。`runtime` 是导出时当前应用/PyMuPDF 的观测值，并明确标为 `export_time_not_processing_time`；不能作为旧批次的执行版本。GUI 不会为探测 qpdf 版本启动慢外部命令。

处理和验证分开：`engine_success` 只说明处理引擎报告结果，`published` 是最终发布事实；`save_reopen`、`structural_expectations`、`rule_verification` 独立保留。当前 PD-07 已将实际 `rule_results` / `rule_events` 和有限独立检查传入批次与审计包，范围见[规则结果说明](rule-results.md)。历史批次或缺少观测/检查的规则仍保留 `unknown/null` / `not_verified`；签名有效性保持 `not_verified`，不能依据中文日志推断。

`reason_codes` 直接映射实际终态：如 `processing_failed → engine_failure`、`verification_failed → verification_failed`、`stale_plan → stale_plan`、`stopped → stop_requested`。不会把配置变化与输入变化都细化成已确认的输入变化；源证据没有更细分类时保留原有粒度。非成功文件以 `engine_diagnostic:not_recorded` 标出缺失的详细引擎诊断。`reasons`/`warnings` 来自实际计划中的错误代码、预检决定理由和结构风险。PD02 没有留存完整引擎警告或失败文本，包用 `engine_warnings_not_recorded_by_pd02` 明确限制，不能把不存在的警告记录解释为没有警告。导出不包含正文、元数据、密码、凭据、绝对路径、进程命令行或本地化日志正文。

## CSV 安全与往返

所有单元格在写入前检查。原值去掉前导空白后以 `= + - @` 开头，或原值本身以单引号开头时，统一增加一个单引号前缀。检查覆盖空格、制表、回车、换行和 Unicode 空白前缀，避免电子表格把字段当作公式。CSV 使用标准引号与 `newline=''`，Unicode 和多行字段不会被截断。

程序读取这种 CSV 后，可对每个单元格调用 `decode_csv_cell` 恢复原始字符串。它只移除本 schema 加入的一层安全前缀；直接用普通 CSV 读取器会看到安全前缀。`reasons`、`warnings` 和 `missing_evidence` 列用 LF 连接列表，多行字符串可往返，但列表边界应以 JSON 为准。

自动测试使用合成小 PDF 和真实 ProcessWorker，覆盖同名不同目录、失败、处理中停止、GUI 入口、模态重入、CSV 公式/换行、部分写入和核验异常。这是软件行为验证；真实业务样本、Acrobat、签名、监管专家和外部设备资格仍未完成。
