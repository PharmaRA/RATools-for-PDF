# Acrobat 字体嵌入自动化探索记录

## 背景

RATools 曾尝试在“嵌入缺失字体”场景中自动调用 Acrobat Pro Preflight，对存在未嵌入非 Base14 字体的 PDF 执行修复，再由 RATools 做后验证。

实际验证中，用户可在 Acrobat 图形界面手动执行 Preflight 并成功嵌入字体，说明 PDF 文件和 Acrobat 本机 profile 本身可用。问题集中在 Acrobat 自动化调用链路。

## 已验证的路径

1. Acrobat COM `PDDoc.GetJSObject().eval()`

   该路径在本机返回 `E_NOTIMPL / 尚未实现`，无法作为稳定脚本执行入口。

2. Forms OLE `AFormAut.App.Fields.ExecuteThisJavascript`

   该路径可以执行部分 JavaScript，并能返回诊断信息。后续脚本改为通过该入口执行 Preflight 探测和调用。

3. Preflight profile 名称

   英文默认候选如 `Embed missing fonts`、`Embed fonts` 在中文 Acrobat 中匹配不到。用户本机可见的正确 profile 名称是 `嵌入缺失的字体`。

4. `preflight(profile, bOmitFixups)` 参数

   `bOmitFixups=true` 只做检查，不执行修复。若继续自动化，修复调用应使用 `false`。

5. Droplet / 快捷批处理

   曾尝试通过 Acrobat Droplet 绕过 JavaScript Preflight 调用，但本机出现“快捷批处理无法在较低版本的 Acrobat 中运行”一类提示。即使 `.pdf` 默认关联和 COM 注册指向 Acrobat Pro DC，Droplet 仍不稳定。

## 失败现象

在能枚举或指定到 `嵌入缺失的字体` profile 后，自动化仍失败于执行阶段：

```text
失败步骤：execute preflight: 嵌入缺失的字体
```

同时，Acrobat 手动执行同一 profile 可以成功嵌入字体。因此当前判断是：Acrobat 的 GUI Preflight 可用，但 COM/OLE 脚本上下文下执行 fixup 不可靠。

## 当前产品决策

自动字体修复工作流不再作为批处理选项执行。`embed_nonstandard_fonts` 旧配置残留会被后端忽略，避免再次触发 Acrobat 自动化错误。

当前稳定方案是手动交接：

1. RATools 预检发现未嵌入字体风险。
2. 用户在左侧队列选中目标 PDF。
3. 点击“嵌入缺失字体”按钮。
4. RATools 打开 Acrobat 或系统默认 PDF 程序。
5. 用户在 Acrobat 中进入“所有工具 > 印刷制作 > 印前检查”，选择“嵌入缺失的字体”，修复并保存。
6. 回到 RATools 重新预检确认风险是否消失。

## 后续如需继续自动化

优先研究 Acrobat 官方是否提供可用于 Preflight fixup 的受支持命令行或 SDK 接口。若继续使用 COM/OLE，应先解决脚本上下文下 Preflight fixup 执行失败的问题，再考虑恢复后验证闭环。
