# 版本化处理配置

处理配置只携带规则、规则参数、版本和来源说明。输出目录、覆盖原件、处理后打开目录、并行设置、主题与其他个人设置不会导出。中国/美国 eCTD 是沿用应用现有规则的便利预设，来源与适用范围是应用说明，不代表官方认证，也不能替代对当前申报要求的独立确认。

## 桌面操作

1. 在主窗口选择“中国eCTD”“美国eCTD”，或手动勾选处理规则。图像压缩卡片可以选择预置或自定义 DPI（72–600）和 JPEG 质量（1–95）。
2. 点击快速预设栏右侧的“导出配置”，选择一个新的 `.json` 文件名。已有文件不会被覆盖。取消文件选择不会创建文件。
3. 在另一台电脑的主窗口点击“导入配置”，选择该 JSON。
4. 阅读“配置差异预览”：它显示配置 ID/版本、机构或地区、增加规则、移除规则、每个参数的前后值，以及验证级别。只有确认后才应用。点击取消不会改变规则、参数或 settings.ini。
5. 点击“保存为常用”保存当前完整配置，再点击“我的常用”可恢复规则与图像参数。旧 settings.ini 的 `Presets/FavoriteOptions` 列表和逗号字符串仍可迁移；旧键保留，新增 `Presets/FavoriteProfile` 保存版本化 JSON。只读取常用规则和必要图像参数，不读取无关设置来构建配置。
6. 开始处理时，处理计划冻结当前配置。预览期间更换规则、图像参数或配置身份会使计划失效，需要重新生成。处理结束后从“日志”点击“导出最近批次审计包”，在 `evidence.json` 的 `configuration.profile` 查看精确快照，`configuration.profile_digest` 查看其 SHA-256。

导入确认后改变任一实际处理规则或图像参数，后续获取的配置将改为 `custom.local` / `1.0.0`，避免继续使用原配置的身份。仅改变输出目录或覆盖选项不改变处理配置身份。关闭程序会沿用现有选项保存行为；会话重启后未重新应用“我的常用”或内置预设时，当前组合作为自定义配置。

## 格式和拒绝条件

当前 `schema_version` 为整数 `1`，`profile_version` 使用 `major.minor.patch`。必需字段为：

- `profile_id`：1–80 个 ASCII 字母、数字、点、下划线或连字符，首字符为字母或数字。
- `profile_version`、`agency_or_region`、`schema_version`。
- `enabled_rules`：已知处理规则 ID 的列表，不允许重复或未知规则。
- `rule_parameters`：当前仅支持已启用 `compress_images` 的 `{ "dpi": 整数, "quality": 整数 }`；不启用图像压缩时为 `{}`。
- `validation_level`：当前仅 `structural`，不意味着规则级或官方资格已验证。

可选的 `source` 和 `applicability` 是便携来源与适用范围说明。其他字段、重复 JSON 字段、未来 schema、未知参数、错误类型、越界值以及互斥组合会直接拒绝，不会截断数值或静默丢弃规则。互斥包括 A4/Letter、标准/极限压缩、删除/展平注释、相互冲突的链接边框、清除/写入标题，以及删除全部链接书签与修改这些对象。

桌面压缩卡片只支持三个完整级别：无损压缩规则为空；或 `compress_standard + generate_object_streams + realloc_flate`；或 `compress_aggressive + generate_object_streams + realloc_flate`。单独导入 `compress_standard` 等不完整组合会拒绝，避免控件自动加入额外规则。共享同一复选框的规则别名也必须完整匹配。导入会同时检查预期映射和控件实际选中结果，确认精确一致后才保存。

内置配置 ID 为 `builtin.china`、`builtin.us`，版本为 `1.0.0`；沿用的规则集合分别为 16 和 13 项。常用旧格式迁移为 `favorite.legacy` / `1.0.0`，手动配置为 `custom.local` / `1.0.0`。自定义版本不会自动递增；需要正式区分修订时，可以在导出的 JSON 中设定符合格式的新 ID/版本后再导入。摘要仍会区分相同 ID/版本下的不同内容。

## 证据与限制

PD02 的 `pd02.engine.v1` 增加可选的 `profile` 与 `profile_digest`；配置整体摘要同时绑定这些字段。PD03 审计包保留并验证两者。旧批次若没有 profile 快照，保持旧格式，不补造命名配置。摘要绑定内容完整性，不构成签名、作者身份或监管认证。

自动化验收使用真实 Qt 窗口的离屏控件和合成 PDF/结构化批次。真实业务样本、Acrobat、签名有效性、监管和外部设备资格仍待独立验收。当前批次的规则事实与有限独立检查见[PD-07 规则结果](rule-results.md)；未观测/未检查项与历史缺失证据仍为 unknown/not_verified，配置的 structural 级别不因此升级。
