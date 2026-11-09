# PD07 规则执行结果

`ratools_pdf.config.rules_catalog.RULE_METADATA` 是稳定规则 ID 与显示文案的目录，覆盖 UI 规则，以及既有引擎入口 `remove_pdf_restrictions`、外部字体流程 `embed_nonstandard_fonts`。UI 仍通过 `MODULES` 展示原有勾选项。

`PRECHECK_OPTION_ALIASES` 也由此目录持有；`precheck` 保留原来的公开别名和选项解析逻辑。复合规则与基础规则共享预检身份，执行结果保留用户选中的 ID。例如 `cleanup_remove_external_uri_and_text_black` 不会被改名为 `cleanup_remove_external_uri`。

## 调用接口

既有入口保持兼容：

```python
success, message = PDFProcessor.process_document(
    input_path, output_path, options, processing_mode="smart",
    compression_settings=None,
)
```

需要结构化事实时使用相同参数的 `PDFProcessor.process_document_detailed(...)`：

```json
{
  "schema_version": "pd07.processing.v1",
  "success": true,
  "message": "用户可见结果",
  "rule_results": [],
  "rule_events": []
}
```

每个被选规则都会有一条 `rule_results`。结果按稳定 ID 排序，管线仍按原有步骤顺序执行。事件包含同一条规则的字段与 `phase`，顺序来自实际管线调用；阶段值为 `inspect`、`plan`、`execute`、`verify`、`error`、`result`。

## 字段含义

| 字段 | 类型 | 证据含义 |
| --- | --- | --- |
| `rule_id` | string | 稳定 ID，独立于中文标题 |
| `matched` | boolean/null | 本次预检命中情况；force 未预检、无法预检或预检失败为 null |
| `planned` | boolean | 解析后的选项是否进入文档处理计划 |
| `executed` | boolean | 规则所在步骤实际开始；失败步骤仍可为 true |
| `changed` | boolean/null | 实际 mutation 点报告的事实；无法观察时 null |
| `verified` | boolean/null | 保存后独立重开检查所支持的不变量；无独立检查为 null |
| `verification` | string | `passed`、`failed`、`not_verified` |
| `skipped_reason` | string/null | 预检未命中、外部流程、优先规则覆盖、管线中断等稳定原因码 |
| `warning` | string[] | 无法预检、变更不可观测、检查或执行错误等稳定代码；不包含文档内容 |
| `status` | string | `pending`、`running`、`executed`、`skipped`、`failed` |

选中、执行、变更、验证是四个不同事实。运行成功不会自动令 `changed` 或 `verified` 为 true。已跳过的规则不会产生 `execute` 事件。规则开始后发生异常，结果保留此前已观察到的变更，后续规则保留 `executed=false`、`changed=null`，原因是 `pipeline_aborted`。变更发生于失败的内存文档也可被记录；文件是否发布仍以 worker 的 `published` 为准。

既有互斥优先级保留。例如同时传 A4 与 Letter 时仍选择 A4；Letter 结果记录 `superseded_by_page_size_a4`，不会被报告为已执行。`filename_ectd_format` 属 worker 输出命名流程，文档管线记录 `output_naming_handled_by_worker`；字体嵌入记录 `external_workflow_required`。

## 当前可观测与后置检查范围

实际 mutation 观察来自标题写入、纸型适配返回的修改页数、URI/失效/非标准链接删除、复合去色内容流修改、批注/附件/动态内容删除。`fast_web_view` 补全空标题的副作用也记录于该规则。

书签组、链接样式组、默认视图的历史 mark、标签/元数据清理，以及 qpdf/保存优化当前不从“执行过”推导变更，缺乏独立 mutation 观察时 `changed=null` 并带 `mutation_not_observed`。这些规则的既有行为仍由原有管线执行。

独立重开输出后，仅支持以下检查：

- `title_from_filename`：标题等于输入文件名。
- `page_size_a4` / `page_size_letter`：每页尺寸在目标纸型容差内，并保留横竖方向。
- `cleanup_remove_external_uri`：没有 URI 页面链接。
- `cleanup_remove_annotations`：没有普通注释。
- `cleanup_remove_attachments`：附件计数为零。

复合去色规则的 URI 删除只是部分后置条件，因此不会据此将整个复合规则设为已验证。其他规则维持 `verified=null` / `not_verified`。后置读取异常仍为未知，并附错误类型码；既有保存重开、结构预期和签名有效性检查的职责不变。规则检查通过也不代表法规资格、视觉质量或签名有效性通过。

## PD02 / PD03 / PD04 消费

PD02 计划新增 `rule_results`，同时保留原有 `rules` 字段。冻结的 smart 计划依然使用原有 effective options 执行，worker 会把实际执行结果合并回冻结计划，保留未执行规则的预检未命中事实。子进程使用 detailed API；tuple 入口继续给旧调用者使用。

worker 终态事件提供 `rule_results`、`rule_events` 与保守聚合 `rule_actual_changed`。只要存在实测变更，聚合为 true；所有规则都已知无变更时为 false；其余为 `unknown`。PD06 的发布锁、停止/跳过/截止时间判断、指纹校验和原子发布顺序保留。

控制器将这些事实保存至 PD02 batch。既有日志修改列通过实际 `changed=true` 的稳定规则 ID 查询 catalog 文案，历史无 PD07 数据的事件保留原有日志兼容路径。PD03 JSON 附加同一规则结果和事件，并用实际结果更新每条规则的 `actual_changed` / `verified`；没有新增字段的旧 evidence 仍可导出。原有 schema version、CSV 字段与完整性校验兼容。

PD04 profile 的 `enabled_rules` 仍使用原 catalog 稳定 ID，配置摘要、参数、冻结计划校验不变。PD07 不扩展 profile 格式，也不将所有规则统一标为已验证。

## 验证入口

使用仓库安全测试入口，临时 PDF 仅含合成文本与 `example.invalid` 链接，清理移入回收站：

```powershell
python -X utf8 -X faulthandler -B scripts/run_tests_safely.py test_rule_results.py
python -X utf8 -X faulthandler -B scripts/run_tests_safely.py
```

专项覆盖真实 PDF 的有变化/无变化、smart 跳过、无法可靠预检、复合去色、部分执行失败、打开前失败、互斥纸型，以及子进程 → worker → PD02 host → PD03 的事实保留。它们属于软件行为验证；真实业务样本、Acrobat 及外部法规资格仍需独立验收。

双复合规则归因回归中，URI-only 使用真实可枚举 URI；invalid-only 在 `get_links` 枚举边界模拟 LINK_NONE，实际删除和文本去色仍处理真实 PDF。该对称用例仅验证分支归因，不扩张为 PyMuPDF 对所有失效链接的解析资格。
