# RATools-for-PDF 全面重构路线图（2026-07-26）

> 工作分支：`worktree-refactor-architecture`（worktree 位于 `.claude/worktrees/refactor-architecture`，基于 `main@ff322f8` v0.7.3）。
> 基线状态：87 个单元测试全部通过（`python -m unittest discover -s tests -p "test_*.py"`）。

## 1. 背景

v0.7.0 已完成第一轮结构重构（Phase 1–3：扁平脚本 → `ratools_pdf` 包、拆分三大文件、删除根目录兼容 shim，见 `2026-07-04-project-structure-refactor-phase-1-2.md`）。那一轮是"按行搬家"：文件位置合理了，但**职责边界、依赖方向和重复代码没有解决**。本路线图是第二轮重构：在不改变用户可见行为的前提下，解决上帝类、循环依赖、字符串协议和零测试核心管线四大问题。

## 2. 现状问题清单（按层，含代码位置）

### 2.1 pdf 层（核心处理，最高风险区）

| 问题 | 位置 | 严重度 |
|---|---|---|
| `process_document` 单函数 588 行，内含 4 个闭包、三层嵌套，**零测试覆盖** | `processor.py:377-964` | 高 |
| `_processor_cls()` 人为循环依赖：5 个子模块经 PDFProcessor 门面回环调用自己，共 **84 处**（precheck 63、hyperlink_styles 10、page_layout 6、qpdf 4、bookmarks_links 1） | 各 pdf 子模块 | 高 |
| PDFProcessor 含 ~50 个纯转发 static method（~290 行无逻辑委托层） | `processor.py:85-374` | 中 |
| `build_precheck_report` 298 行 if 串，无检查器抽象；新增检查项要同时改 4–5 处 | `precheck.py:789-1086` | 中 |
| controllers 直接调用 pdf 层 **7 个下划线私有方法**（`_pdf_has_signature`、`_read_pdf_header_version`、`_collect_annotation_findings_for_path` 等） | `main_controller.py:579-610,1088`、`workers.py:536-539` | 中 |
| 5 个子模块头部 19 行 import 完全相同且大量未使用（qpdf.py import fitz 未用等） | 各模块 :1-19 | 低 |
| 死代码：`_run_font_embedding_workflow`（254-286）、`_overlay_black_text_in_rect`（754-783，每页循环内定义且无调用） | `processor.py` | 低 |
| 蓝色判定逻辑 3 处重复且阈值不一致（+40 于 0-255 域 vs +0.1 于 0-1 域） | `hyperlink_styles.py:27-31,51-55,222`、`processor.py:747-752` | 中 |
| qpdf 路径硬编码开发机路径 `D:\Program Files\qpdf 11.9.1\...` | `qpdf.py:27` | 低 |
| 43 处 `except Exception: pass` 静默吞异常；错误反馈四种形态并存（tuple/dict/异常/dataclass） | pdf 包全域 | 中 |

### 2.2 controllers 层

| 问题 | 位置 | 严重度 |
|---|---|---|
| `MainController` 上帝类：1717 行、约 70 方法、10 类职责（队列模型、树 UI、5 种 worker 编排、Acrobat 探测、PDF 元数据解析、日志导出、更新检查…） | `main_controller.py` | 高 |
| **eCTD 重命名逻辑两份实现**：预览用一份、实际处理用一份，发散即"预览名 ≠ 输出名" | `workers.py:79-85` vs `io_actions.py:21-27` | 高 |
| 中文状态字符串作为 worker↔controller↔log_export 三方协议（"处理完成"/"已跳过"等），log_export 还用正则反解日志文本 | `workers.py` emit、`main_controller.py:1340-1384`、`log_export.py:22-90` | 高 |
| worker `progress(int,…)` 发行索引，controller 靠"猜哪个列表"路由（4 类 worker 共用一个槽） | `main_controller.py:1330-1338` | 中 |
| 忙碌互斥守卫三连检查重复 6 处 | `main_controller.py:664,784,898,923,1125,1178` | 中 |
| 跨平台打开文件/目录三分支重复 4 处；common_base 计算重复 5 处；按钮复位块 processing_finished/error 大段相同 | 详见分析 | 中 |
| controller 越层改 view 内部：写 `is_applying_preset`/`active_preset_key`、调私有 `_set_preset_button_state`、monkey-patch `drop_zone.mousePressEvent`、共 214 处 `self.view.` | `main_controller.py:82,171-182` 等 | 高 |
| 右键菜单 QSS 硬编码亮色（**暗色主题下右键菜单不适配**，实际 bug）；树节点状态色 QColor 硬编码不随主题 | `main_controller.py:442-467,1341-1353` | 中 |
| 死代码 `_process_document_task`；`last_detection_results = []` 连写两遍 | `workers.py:23-24`、`main_controller.py:73-74` | 低 |

### 2.3 ui 层

| 问题 | 位置 | 严重度 |
|---|---|---|
| `MainWindow.__init__` 535 行（数据定义 + 六大区块 + 接线 + 设置加载全在一起） | `main_window.py:74-608` | 高 |
| 领域数据长在窗口类里：`MODULES_DATA`（40 条规则目录）、`PRESET_OPTIONS`（eCTD 预设） | `main_window.py:32-72,115-193` | 高 |
| 选项标题双份维护**已经漂移**："书签动作改为新窗口打开"（precheck.py:66）vs "书签动作：新窗口打开"（main_window.py:141） | 两处 | 中 |
| 跨文件动态属性状态机：controller `setProperty`（stopMode/hasFailedItems 等 6 个），view `refresh_selection_summary`（112 行）读它们决定按钮状态 | `main_controller.py` 多处 → `main_window.py:1042-1153` | 高 |
| dialogs.py 样板重复：底部按钮行 ×8、toggle 按钮组 ×5、卡片 QFrame ×7、repolish 三行 ×4、目录选择器 ×2 | `dialogs.py` 全域 | 中 |
| 两个完整对话框内联在 MainWindow 里（`show_signed_files_prompt` 99 行、`show_major_update_prompt`） | `main_window.py:630-772` | 中 |
| Win11 判定两套并存（`platform.release()=="11"` vs build≥22000）；DWM ctypes 逻辑三份 | `dialogs.py:37,119-144`、`theme.py:262-318`、`platform.py` | 低 |
| `AboutDialog.set_update_checking/set_update_result` 无条件引用仅在 `ENABLE_UPDATE_CHECK` 时才创建的按钮（潜在 AttributeError） | `dialogs.py:1111-1158` | 中 |
| theme.py 模板内残留一处硬编码 `#FFFFFF`；`LogDialog.log_blocks` 计算后从未读取 | `theme.py:477`、`dialogs.py:579` | 低 |
| QSettings 读写长在 MainWindow；`all_checkboxes` 混用 option id 与中文文案当 key（"覆盖原始文件 (不推荐)" 同时是显示文案和 dict key） | `main_window.py:90-113,590-596,774-860` | 中 |

### 2.4 入口与工程化

- `main_no_update.py` 依赖 import 顺序副作用（env 必须先于任何 ratools_pdf import 设置），是隐式契约。`ENABLE_UPDATE_CHECK` 在 6 处重复防御判断。
- 无 `pyproject.toml`、无 lint/format 配置、核心模块几乎无类型标注（上一轮 Phase 4 遗留项）。
- 测试缺口：`process_document`（588 行管线）、hyperlink_styles、page_layout、qpdf、font_embedding_providers、ProcessWorker 编排逻辑全部零测试；MainController 仅 1 个方法被测。

## 3. 重构原则

1. **行为不变**：不改功能、不改可执行文件名、不动双入口（`main.py`/`main_no_update.py`）与打包脚本的根目录位置。
2. **测试先行**：先给要动的高风险区补特征测试（characterization tests），红线是任一阶段结束时全量测试绿。
3. **每阶段独立可合并**：每个 Phase 是一个可独立 PR/合并的增量，随时可以停在任一阶段出口。
4. **依赖方向单一化**：目标分层 `config ← pdf ← services ← controllers ← ui(组装)`，pdf 层继续保持零 Qt 依赖。
5. **协议显式化**：状态用常量/枚举，跨层数据用 dataclass，不再用中文显示文案当协议或 dict key。

## 4. 阶段计划

### Phase 0：安全网（先补测试，不动生产代码）

**目标**：给后续最危险的 Phase 3 建立回归保护。

- 为 `process_document` 按规则组补特征测试（仿照现有 `test_pdf_processor_roundtrip.py` 用 fitz 运行时合成 PDF）：
  - 初始视图组（title_from_filename / open_page_first / page_layout_default / PageMode / 折叠书签）
  - 页面尺寸组（A4/Letter resize）
  - 书签规则组（承前缩放/新窗口/删外链/删失效）
  - 超链接规则组（相对路径转换/样式/边框）
  - 清理组（删 URI、删批注、删元数据、删附件、一键全删）
  - qpdf 分支（版本转换/线性化/解限制——qpdf.exe 在 `plugins/` 内可直接调用）
- 为 `build_precheck_report` 主要检查项补测试（目前只覆盖批注与失效引用两项）。
- 为 `resolve_processing_options` smart/force 模式补齐分支测试。

**出口**：新增 ~20-30 个测试，全绿。**估算 400-600 行测试代码。**

### Phase 1：零风险清理（机械性删除与顺带 bug 修复）

- 删死代码：`workers.py` `_process_document_task`、`processor.py` `_run_font_embedding_workflow`、`_overlay_black_text_in_rect`、`LogDialog.log_blocks`、`main_controller.py:73-74` 重复行。
- 清理 5 个 pdf 子模块的相同 19 行头部中未使用的 import。
- 删除 `qpdf.py:27` 硬编码开发机路径（保留 resource path → `QPDF_PATH` 环境变量 → PATH 三级查找）。
- 顺带修复分析中发现的真实缺陷：
  - `AboutDialog.set_update_checking/set_update_result` 的条件按钮 AttributeError 风险；
  - `theme.py:477` 硬编码 `#FFFFFF` → `$text_on_primary`；
  - 统一 Win11 判定为 `platform.py:is_win11()`（build≥22000）一处。
- 合并 `log_export.py` 内两份 `_log_time_to_seconds`。

**出口**：全量测试绿；diff 以删除为主。**估算净减 ~300 行。**

### Phase 2：共享基础设施与协议显式化

**目标**：先立地基，否则后面每个拆分都会复刻重复块。

1. `ratools_pdf/common/status.py`：处理状态常量（"处理完成"/"处理失败"/"已跳过"/"已停止"…字面值保持不变以兼容现有日志），状态→颜色映射从 `main_controller.py:1341-1353` 迁入并接 theme token。worker、controller、log_export 三方共用。
2. `ratools_pdf/config/rules_catalog.py`：合并 `MODULES_DATA` + `PRESET_OPTIONS` + `PRECHECK_OPTION_TITLES` 为单一规则目录（id、标题、描述、模块分组、预设归属、可预检标志一处定义），**消灭标题漂移**。ui 与 precheck 都从这里读。
3. `ratools_pdf/services/system_shell.py`：`open_with_default_app` / `reveal_in_file_manager` / `open_directory` + Acrobat 探测（迁 `main_controller.py:261-324,502-535,1707-1717`），消灭 4 处平台三分支重复。
4. **统一 eCTD 命名实现**：删掉 `workers.py:79-85` 内联版，公开 `io_actions` 的实现为唯一入口（顺带把 `io_actions.py` 更名/移到 `services/io_paths.py`，它不依赖 Qt）。
5. worker 进度协议改造：`progress` 信号直接携带 `file_path`（各 emit 点已有该值），删除 `update_progress` 的列表猜测路由。
6. `ratools_pdf/ui/win32.py`：合并三份 DWM/ctypes 逻辑（`dialogs.py:119-144`、`theme.py:262-335`、`platform.py`），DWM 属性 ID 命名常量化。
7. 右键菜单 QSS 迁入 theme 中央模板（修复暗色主题右键菜单 bug）。

**出口**：全量测试绿 + 手工冒烟（明暗主题右键菜单、拖入处理一批文件）。**估算 +400/-350 行。**

### Phase 3：pdf 层内科手术（风险最高，靠 Phase 0 保护）

1. **拆环**：删除 5 个子模块里的 `_processor_cls()` 回环（84 处），子模块之间直接函数 import；依赖方向变为 `processor → {precheck, qpdf, page_layout, bookmarks_links, hyperlink_styles} → (无)`。
2. **门面收缩**：PDFProcessor 删掉 ~50 个纯转发 static method，只保留 controllers 实际使用的入口（`process_document`、`build_precheck_report`、`resolve_processing_options`、书签/链接导入导出 ×4）；controllers 用到的 7 个私有方法提升为具名公开函数（如 `inspect.pdf_has_signature()`、`inspect.read_header_version()`，可归入新模块 `pdf/inspect.py`）。
3. **管线化 `process_document`**：按现有 10 步拆为独立 step 函数（每步签名统一 `step(doc, options, ctx) -> None`，ctx 收集 applied_changes/change_counts），主函数缩为 <100 行的顺序调度；闭包（`_to_point`、`_normalize_bookmark_dest`、`_is_span_blue`）提升为模块函数并与 hyperlink_styles 的重复实现合并（统一蓝色判定阈值到一处）。
4. **precheck 检查器注册表**：每个检查项一条注册记录（option_id、detect 函数、report_only 标志），`build_precheck_report` 缩为遍历注册表；标题从 `rules_catalog` 读取。
5. 结果对象化（**本阶段只做增量**）：`process_document` 返回值保持 `(bool, str)` 兼容，内部先引入 `ProcessResult` dataclass，controller 侧消费 structured 事件而非解析文本；log_export 的正则解析标记为 deprecated，导出改走结构化行（`process_log_rows` 已存在）。

**出口**：Phase 0 特征测试全绿；`ra_test/` 样例 PDF 手工处理对比输出一致。**估算 ±1500 行改动，pdf 层净减 ~400 行。**

### Phase 4：MainController 拆分

> **执行记录（2026-07-26）**：已拆出 precheck / detection / io / log / font_embedding /
> tree_actions / update 七个子控制器与 system_shell / pdf_inspector 两个服务，
> MainController 1717 → 785 行。**决策**：批处理生命周期与文件队列保留在
> MainController —— 二者共享 loaded_files/file_nodes/日志缓冲，且 update_progress
> 是全部 worker 的共享进度路由，强拆会制造大量 host 间接调用，可读性反而下降。
> MainController 的收敛职责即"文件队列 + 批处理协调 + 组合根"。
> 原计划中的 file_queue_controller / processing_controller 不再单独拆出。

前置件（已在 Phase 2 就位：status 常量、system_shell、file_path 进度协议）。

1. `controllers/task_guard.py`：集中持有各 worker 引用，`ensure_idle(action) -> bool`，消灭 6 处忙碌守卫重复。
2. 按职责拆出子 controller（MainController 退化为组合根 + 信号分发，~150 行）：

| 新模块 | 迁移内容 | 估算行数 |
|---|---|---|
| `file_queue_controller.py` | 队列增删清 + 树节点缓存 | ~280 |
| `processing_controller.py` | 批处理生命周期 + 进度 + 批次摘要（把 finished/error 的重复复位块合并为 `_reset_processing_ui()`） | ~430 |
| `precheck_controller.py` | 预检 + 建议应用 + CSV 导出 | ~190 |
| `detection_controller.py` | 只读检测（与 workers 的 KIND_LABELS 映射合一） | ~90 |
| `update_controller.py` | 更新检查（`ENABLE_UPDATE_CHECK` 判断收敛到此一处） | ~140 |
| `io_controller.py` | 书签/链接导入导出向导 | ~90 |
| `log_controller.py` | 日志对话框 + 导出 | ~130 |
| `tree_actions_controller.py` | 右键菜单/双击/定位（PDF 详情文本生成迁 `services/pdf_inspector.py` 纯函数） | ~120 |
| `font_embedding_controller.py` | Acrobat 手动嵌字 UI 流程（探测逻辑已在 system_shell） | ~50 |

3. controller→view 写操作收敛为 view 公开方法：`view.set_processing_state()`、`view.apply_options(ids)`、`view.get_parallel_worker_settings()` 等，废除对 `is_applying_preset`/`_set_preset_button_state` 的直捣与 `drop_zone.mousePressEvent` monkey-patch（换成 DropZoneLabel 自定义 clicked 信号）。
4. `setProperty` 跨文件状态机改为显式状态对象：controller 调 `view.update_footer_state(FooterState(...))`，view 端 `refresh_selection_summary` 拆为纯函数 + 薄 apply 层。

**出口**：全量测试绿 + 为新子 controller 补守卫类单测；手工冒烟全功能。**估算 ±1800 行。**

### Phase 5：UI 层拆分

1. `ui/dialogs/` 包化（`__init__.py` re-export 保持现有 import 兼容）：
   - `base.py`（FramelessDraggableDialog）、`message_box.py`、`io_wizard.py`、`log_dialog.py`、`settings_dialog.py`、`about_dialog.py`、`font_embedding.py`
   - `prompts.py`：从 MainWindow 迁入签名文件提示与重大更新提示两个内联对话框
   - `_builders.py`：`build_button_row()` / `build_toggle_group()` / `build_card()` / `build_dir_picker_row()` / `repolish()`，一次性消掉 20+ 处样板
2. `main_window.py` 瘦身：
   - `ui/settings_store.py`：QSettings 封装（load/persist/键映射）
   - `ui/selection_model.py`：预设/勾选状态机（apply_preset/toggle_preset/restore/clear/favorite），对外发信号
   - 六大区块拆为构建模块（header/nav/queue_panel/rules_panel/preset_bar/footer_bar），MainWindow 本体缩为组装 + 公共 API（目标 ~200 行）
3. `refresh_selection_summary` 纯函数化（输入状态 → 输出各按钮 enabled/visible/tooltip），配 Phase 4 的 FooterState。

**出口**：全量测试绿 + `test_theme.py` 扩展覆盖新对话框包 import；手工冒烟每个对话框明暗主题。**估算 ±1600 行，样板消除后 UI 层净减 15-20%。**

### Phase 6：工程化收尾（可选，随时可做/可不做）

- `pyproject.toml`（项目元数据 + 可选 ruff 配置），CI 加 lint job。
- 入口显式化：`app.run(enable_update_check=...)` 参数替代 env 副作用契约（保留 env 兼容）。
- 新公开 API 补类型标注；pdf 层静默 `except Exception` 分级收敛（可预期失败记日志、意外失败上抛）。
- README「项目结构」小节更新。

## 5. 风险与缓解

| 风险 | 缓解 |
|---|---|
| `process_document` 零测试下重构出错 | Phase 0 先建特征测试；Phase 3 每一步拆分单独 commit，出错可 bisect |
| 中文状态字符串被日志导出/用户习惯依赖 | Phase 2 只引入常量不改字面值；结构化导出已有 `process_log_rows` 双轨 |
| 并行处理（mp.Process + Pipe）行为微妙 | ProcessWorker 编排逻辑本轮不重写算法，只动 import 与信号签名；`ra_test/` 样例做串行/并行手工冒烟 |
| 打包回归（PyInstaller/Nuitka 找不到模块） | 不动根目录入口与 bat；每个大阶段结束跑一次 `build_pyinstaller.bat` 冒烟 |
| 阶段过大难以 review | 每 Phase 内部再按 commit 粒度拆（一个移动/一个合并一个 commit），CHANGELOG 记入 Unreleased |

## 6. 非目标（本轮不做）

- 不改任何处理算法与规则语义；不新增功能。
- 不迁移到 pytest（CI 与 CONTRIBUTING 均为 unittest，保持不变）。
- 不采用 `src/` 布局（沿袭上轮决策，待包 API 稳定后另议）。
- 不引入 i18n 框架（文案集中化即可，翻译不在范围）。
- 不重写 log_export 的旧文本解析（标记 deprecated，双轨运行）。

## 7. 执行顺序与量级总览

```
Phase 0  安全网          ~500 行测试    低风险   ★ 必须最先做
Phase 1  零风险清理      净减 ~300 行   极低风险
Phase 2  共享基础设施    ±750 行        低风险   ★ 后续阶段的地基
Phase 3  pdf 层手术      ±1500 行       高风险   ★ 收益最大（拆环+管线化）
Phase 4  controller 拆分 ±1800 行       中风险
Phase 5  UI 层拆分       ±1600 行       中风险
Phase 6  工程化收尾      ±300 行        低风险   可选
```

推荐节奏：0→1→2 可以连续做完（合计 1-2 天量级，全程低风险）；3、4、5 各自独立成 PR，每个完成后回 main 合并一次再继续，避免长期分叉。
