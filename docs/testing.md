# 测试分层与样本回归

在项目根目录的 Windows PowerShell 中运行以下安全入口。它安装 offscreen Qt、限定临时根目录、回收站清理适配器和禁止物理删除的 Python audit hook；Windows multiprocessing spawn 同样安装清理保护。不要绕过入口直接运行 unittest。

```powershell
# 默认完整普通回归（所有 unit + integration；资格工作负载需显式运行）
python -X utf8 -X faulthandler -B scripts/run_tests_safely.py

# 保留既有单文件/通配符位置参数
python -X utf8 -X faulthandler -B scripts/run_tests_safely.py test_pdf_qualification.py

# 独立层选择；未选模块在发现阶段不导入
python -X utf8 -X faulthandler -B scripts/run_tests_safely.py --layer unit
python -X utf8 -X faulthandler -B scripts/run_tests_safely.py --layer integration

# 显式资格：输出目录必须全新，每次更换名称；输入样本只读
python -X utf8 -X faulthandler -B scripts/run_tests_safely.py --layer qualification --qualification-manifest docs/qualification/pd01-manifest-v1.json --qualification-output backups/pd08-qualification-001
```

`tests/test-layers-v1.json` 是 schema `1.0`、inventory `pd08-2026-11-09` 的显式模块分类清单。新增测试文件必须归入唯一层；缺失、重复、陈旧归属、重复 JSON 键、未知版本及零匹配选择在执行前报错（exit 2），不会静默通过。`--tests-dir` 与 `--classification` 可供隔离的 runner 行为夹具使用；它们不关闭清理保护。

unit 仅包含路径/版本/发布说明三组纯标准库逻辑。独立 unit 入口阻止外部进程启动与 Qt/PyMuPDF 导入。混合模块整体保守归入 integration，包括 qpdf、PyMuPDF、Qt、原子发布、资格机制的合成 PDF 回归，以及会启动受保护子进程的 runner 自测。文件名为 qualification 的 unittest 模块测试资格机制，并非外部资格证据。

每次输出一个 `LAYER_SUMMARY` JSON。普通回归按实际执行的 TestCase ID 唯一计数（一个父 test 的多个失败 subtest 仅算一个失败），显示 passed/failed/skipped/pending；未选层显示 `not_selected` 和零计数。预期失败记 skipped，意外成功记 failed。类/模块初始化失败记 `suite_errors` 并使层失败，未执行用例保留 pending；这些初始化错误不作为额外 TestCase 加总。资格层按声明的样本计数，blocked 样本归入 pending 缺材料计数；矩阵行单列 `matrix_statuses`，不再加入样本总数。`count_basis` 明确这两种计数口径，禁止将 unittest 数量与资格样本数量相加作为测试通过数。

普通回归成功 exit 0，测试失败或清理失败 exit 1，分类/选择无效 exit 2。资格通过 exit 0，资格失败 exit 1，缺样本、待批准、缺引擎或独立外部证据 exit 2。传入 `--skip-qpdf` 会保留 qpdf pending，不会提升资格状态。当前安全入口要求 Windows 回收站接口；CI 也使用 Windows。

PD01 清单 schema `1.0` 保持兼容；现有未声明版本历史的旧 v1 调用仍可用于机制回归。项目清单为每个样本显式声明 `sample_version`、`expected_metrics_version`，并带 `sample_history` schema `1.0`。当前声明必须与历史记录的 SHA256、预期指标快照、指标版本和来源授权一致；重复修订、复用指标版本但改动指标、缺历史或重复键将拒绝且不创建证据目录。历史中保留旧修订；新增或改动时复制 [待批准记录模板](qualification/sample-change-record-template-v1.json)，填写变更原因与来源授权，增加样本版本，指标变化另增指标版本，并更新当前指针。

模板和本次初始历史均为 `approval.status: pending`、`reference: null`。只有实际存在的独立人工记录才可填写 approved/rejected 和记录 ID；工具不会自动生成批准。样本检查通过后仍保留待批准状态。Acrobat、Word、LibreOffice、扫描仪、打印驱动的真实授权样本以及打包 Windows/外部阅读器矩阵证据仍缺，合成 PDF 与自动测试不会满足这些行。报告只保存相对样本标识、哈希与结构指标，不持久化正文。

CI `regression` 执行安全默认全量；`qualification` 仅在 Actions 手动 workflow_dispatch 勾选 `run_qualification` 时执行 PD01 声明并上传报告。缺材料导致非零是有效缺口信号，不能改成成功退出以使工作流变绿。该工作流未执行外部发布、推送或批准。
