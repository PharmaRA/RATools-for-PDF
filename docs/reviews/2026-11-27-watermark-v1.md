# 去水印首版交付审查与验证

日期：2026-11-27。范围仅为新增去水印后端、独立对话框、主窗口入口及直接依赖。

## 交付

- `ratools_pdf/pdf/watermarks.py`：不可变输入快照、严格页码解析、批注／简单文字／图片调用／路径与组合对象候选、页面独立改写、同源预览、安全另存。
- `ratools_pdf/ui/dialogs/watermark_dialog.py`：后台分析、逐项选择、逐页前后预览、选择变化后的确认失效、取消与线程关闭保护。
- 主窗口“内容合规与安全性 → 去除水印…”：可直接选择文件，预检或主处理忙碌时禁用入口；窗口关闭后释放对话框。
- 使用说明：[去除水印](../manual/watermarks.md)。不新增自动预设或批处理规则。

## 只读审查

审查了对象选择到内容改写、预览到保存、输入失效、共享对象隔离、取消到发布、临时文件归属和窗口关闭路径。审查发现并补齐：独占临时文件创建成功后才取得清理资格；最终输入校验后再次检查取消；逐个包含选择的页面要求当前预览；回收失败明确报告保留路径；关闭工具后释放输入快照。

修正均有专项回归覆盖。未对无关模块作重构或扩大扫描；此次审查不是全部 PDF 兼容性认证。

## 本机执行证据

执行环境为 Windows、CPython 3.15.2，使用现有安全 runner，测试临时内容经回收站清理。

```powershell
$pythonExe = 'C:/Users/Administrator/AppData/Local/Programs/Python/Python315/python.exe'
& $pythonExe -B -X utf8 -X faulthandler scripts/run_tests_safely.py 'test_watermark*.py' --layer integration
& $pythonExe -B -X utf8 -X faulthandler scripts/run_tests_safely.py 'test_hyperlink_content_safety.py' --layer integration
& $pythonExe -B -X utf8 -X faulthandler scripts/run_tests_safely.py 'test_window_lifecycle.py' --layer integration
```

| 范围 | 最终结果 | 日志（仓库根目录起） |
| --- | --- | --- |
| 全量普通回归（最终提交前） | exit 0，618/618，通过，无失败／跳过／待执行；unit 15 + integration 603 | `backups/watermark-full-regression-20261127-162027/full-regression-3.log` |
| 去水印后端及 Qt 操作链（生命周期修复后重验） | exit 0，27/27，通过，无跳过 | `backups/window-lifecycle-20261127-160000/verified-watermarks.log` |
| 复用内容流解析器的直接依赖回归 | exit 0，11/11，通过，无跳过 | `backups/watermark-v1-20261127-152923/tests-content-dependency.log` |
| 窗口生命周期组（修复后） | exit 0，5/5，通过，无跳过 | `backups/window-lifecycle-20261127-160000/verified-lifecycle.log` |

去水印用例包括实际主窗口按钮、分析／选择／预览／另存、批注与链接保留、重叠文字、字体状态、共享内容流／图片／组合对象／批注数组、旋转裁切、复杂指令和扫描图像跳过、页码、输入变化、已有输出与输出竞争、取消、临时文件碰撞、清理失败及后台线程关闭。

中间一次共享批注测试以低层改写后尚未重开的内存页面为基线而失败，改为以实际保存并重开的输入文件为基线后通过；未据此修改批注处理逻辑。中间日志 `tests-final.log` 保留，最终结果以上表为准，历史次数不累加。

生命周期组修复前错误切片（历史失败保留，不代表当前结果）：

```text
test_main_close_preserves_live_pdf_subprocess_until_output_is_published ... Fatal Python error: Aborted
tests/test_window_lifecycle.py, line 196
controllers/workers.py, line 520, _run_serial
```

初次异常后仍驻留的测试进程已结束，其独占临时目录 `backups/test-runs/ff93ee2099d14c00aec08ca0074ae491` 已移入回收站。

## 生命周期崩溃修复与重新验证

2026-11-27 后续复现发现：失败用例单独执行通过，按原顺序执行整组则再次崩溃；Qt 消息处理器捕获到崩溃前的 `QBasicTimer::stop: Failed. Possibly trying to stop from a different thread`。此前测试仅关闭窗口，未明确销毁持有信号连接和关闭回调引用的 Qt 对象，导致其生命周期延续到后续工作线程运行阶段。

修复限定在 `tests/test_window_lifecycle.py`：测试显式持有窗口、控制器和线程；每项测试结束时确认在 GUI 主线程执行，释放受控线程／必要时请求处理线程停止并等待退出，然后关闭窗口，再对每个对象执行 `deleteLater` 和定向 `DeferredDelete`，断言原生对象已经失效。原有关闭阻断、所有 worker 类型和真实 PDF 子进程发布断言不变，生产处理代码未修改。

证据目录为 `backups/window-lifecycle-20261127-160000/`：

- `repro-single.log`：原失败用例单项 1/1 通过；其他 4 项未选择。
- `repro-group.log`：原测试顺序再次复现原生 abort，含上述跨线程计时器警告；异常驻留进程结束后退出码 3。
- `fixed-diagnostic.log`：修复后同一诊断入口整组 5/5 通过，未再出现该跨线程计时器警告。
- `verified-lifecycle.log`：移除诊断入口影响，直接执行原安全 runner，5/5 通过，2.664 秒，退出码 0。
- `verified-watermarks.log`：直接执行原安全 runner，去水印专项 27/27 通过，7.125 秒，退出码 0。

修复后只读审查覆盖引用持有、线程停止与等待、关闭回调所需对象的存活顺序、定向销毁及失效断言。诊断临时脚本及崩溃复现的独占临时目录均移入回收站，保留备份和日志。此次重验不等于全项目回归。

## 提交前全量回归

2026-11-27 使用 Windows CPython 3.15.2 执行原始安全入口的全部 unit + integration，最终 **618/618 通过（unit 15、integration 603），零失败、零跳过、零待执行，188.932 秒，退出码 0**。Qt 清理完整完成，qualification 为 `not_selected`，不计入普通测试总数。

全量验证期间保留三轮证据，目录为 `backups/watermark-full-regression-20261127-162027/`：

1. `full-regression.log`：在去水印主窗口入口测试应用主题时发生访问冲突，退出码 `-1073741819`；主题＋去水印 16 项缩小组合未复现。检查发现新增去水印测试清理了全局延迟删除队列，改为只对自己创建的对话框／主窗口执行定向 `DeferredDelete` 并断言销毁完成。
2. `full-regression-2.log`：修正后全部 618 项执行完毕，Qt 清理正常；601 项通过，17 项因本机缺少构建测试依赖 `packaging` 报错，退出码 1。
3. `full-regression-3.log`：按项目发布锁文件获取 `packaging 26.3` wheel，核对 SHA-256 为 `d7193f7c8e4e93f444fde0262bf90af30e16fa0ad0ad44cb553c87339b23cd1c`，解包到本轮临时目录并通过进程级 `PYTHONPATH` 注入；未修改系统 Python。重跑原完整入口后全部通过。

最终实际执行方式：

```powershell
# packaging 26.3 已按锁文件校验并解包至本轮临时 site 目录。
$env:PYTHONPATH = (Join-Path (Get-Location) 'backups/watermark-full-regression-20261127-162027/test-dependencies/site')
& 'C:/Users/Administrator/AppData/Local/Programs/Python/Python315/python.exe' -B -X utf8 -X faulthandler scripts/run_tests_safely.py
# 执行后恢复原进程 PYTHONPATH，并将临时依赖和诊断脚本移入回收站。
```

新增清理修正经只读审查，测试覆盖与生产代码均未为通过回归而删减。三轮退出码分别留存在 `exit-code.txt`、`exit-code-2.txt` 和 `exit-code-3.txt`；失败历史不与最终统计累加。

## 验证边界

已执行全量普通回归；未执行真实业务样本、外部阅读器资格、远程测试或 EXE 重打包。现有成品不会因此自动包含该功能。扫描图像内的水印、复杂混排或无法隔离的对象不在首版支持范围内。

重要文件修改前备份位于 `backups/watermark-v1-20261127-152923/`；生命周期修复前备份位于 `backups/window-lifecycle-20261127-160000/`；全量验证修正前备份位于 `backups/watermark-full-regression-20261127-162027/`。已有 `.playwright-mcp/` 未参与本次修改。
