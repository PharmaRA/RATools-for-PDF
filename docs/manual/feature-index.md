# 全部功能索引

[返回手册首页](index.md)

本索引按当前源码中的七个模块和可见独立工具排列。42 条规则定义全部有对应说明；其中四条结构压缩规则通过档位组合呈现，不能当作四个独立可见按钮。文档检测没有规则勾选框。

新增独立工具：[去除水印](watermarks.md)，入口位于“内容合规与安全性”；支持对象选择、前后预览、页码范围和另存，不增加批处理规则数量。

## 初始视图与文档属性（6 项）

| 名称 | 说明 |
| --- | --- |
| 设为首页打开 | [打开页及默认缩放](initial-view.md#first-page) |
| 重置页面布局 | [阅读器默认布局](initial-view.md#layout) |
| 重置缩放比例 | [同时影响打开页的边界](initial-view.md#zoom) |
| 设置导览标签 | [书签面板与页面](initial-view.md#navigation) |
| 折叠所有书签 | [书签初始展开状态](initial-view.md#collapse) |
| 同步文件名为标题 | [属性标题与输入名称](initial-view.md#title) |

## 页面与字体标准化（4 条规则及字体工具）

| 名称 | 说明 |
| --- | --- |
| 适配到A4尺寸 | [A4 适配](pages-fonts.md#a4) |
| 适配到Letter尺寸 | [Letter 适配](pages-fonts.md#letter) |
| 展平页面物理旋转角 | [旋转归一](pages-fonts.md#rotation) |
| 展平所有注释与批注 | [固化外观](pages-fonts.md#flatten-annotations) |
| 嵌入缺失字体、打开 Acrobat、保存后复检字体 | [手动交接与三类风险](pages-fonts.md#fonts) |

## 书签管理（5 条规则及数据向导）

| 名称 | 说明 |
| --- | --- |
| 书签设为承前缩放 | [缩放](bookmarks.md#zoom) |
| 书签动作：新窗口打开 | [外部动作](bookmarks.md#new-window) |
| 清理书签外部链接 | [URI 类型边界](bookmarks.md#external) |
| 清理失效书签 | [无效目标](bookmarks.md#invalid) |
| 清理非标准动作书签 | [特殊动作](bookmarks.md#unknown) |
| 书签数据导入/导出... | [CSV 格式、重建与结果位置](data-exchange.md#bookmarks) |

## 超链接处理（8 条规则及数据向导）

| 名称 | 说明 |
| --- | --- |
| 绝对路径转相对路径 | [保留文件名的实际行为](links.md#paths) |
| 超链接设为承前缩放 | [缩放](links.md#zoom) |
| 链接动作：新窗口打开 | [外部文件动作](links.md#new-window) |
| 链接文本设为蓝色 | [文字颜色](links.md#blue-text) |
| 链接区域加黑框 | [黑色边框](links.md#black-border) |
| 标准化有框链接 | [有框转蓝框黑字](links.md#bordered) |
| 标准化无框蓝字链接 | [无框蓝字转蓝框黑字](links.md#unbordered) |
| 清除所有链接边框 | [去框保留动作](links.md#remove-border) |
| 链接数据导入/导出... | [JSON、全部/外链、覆盖/增量](data-exchange.md#links) |

## 内容合规与安全性（11 项）

| 名称 | 说明 |
| --- | --- |
| 删除外部URI链接 | [URI 动作](cleanup.md#uri) |
| 删除外部URI链接并去色 | [URI 与文字去色](cleanup.md#uri-black) |
| 清理失效超链接 | [空动作等问题](cleanup.md#invalid) |
| 清理失效链接并去色 | [失效动作与文字去色](cleanup.md#invalid-black) |
| 清理非标准动作链接 | [特殊动作](cleanup.md#unknown) |
| 彻底清除动态内容 (JS/3D) | [动态内容及能力边界](cleanup.md#dynamic) |
| 移除所有内嵌附件 | [嵌入文件](cleanup.md#attachments) |
| 移除结构化标签 | [标签与辅助阅读](cleanup.md#tags) |
| 清理所有高亮/批注 | [普通注释](cleanup.md#annotations) |
| 清空文档元数据 | [属性清理](cleanup.md#metadata) |
| 移除全部链接和书签 | [全部导航](cleanup.md#all-navigation) |

## 文件级优化与输出（8 条规则）

| 规则名称 | 界面入口与说明 |
| --- | --- |
| PDF版本转换 | [固定目标 PDF 1.7](compression.md#version) |
| 启用线性化 (快速网页浏览) | [流式加载结构](compression.md#linearization) |
| 生成对象流 (Object Streams) | [智能无损/极限无损均包含](compression.md#lossless) |
| 无损重压缩内容流 | [智能无损/极限无损均包含](compression.md#lossless) |
| 标准文件压缩 | [智能无损包含](compression.md#lossless) |
| 深度文件压缩 | [极限无损包含](compression.md#lossless) |
| 压缩内嵌图像 ⚠️ | [界面为“压缩内嵌图像 (扫描件降采样)”及 DPI/JPEG 参数](compression.md#images) |
| eCTD文件名合规格式化 | [改名预览与命名边界](compression.md#filename) |

## 文档检测（2 个只读工具）

| 入口 | 说明 |
| --- | --- |
| 检测批注 (便签/高亮等) | [批注对象扫描](precheck.md#annotations) |
| 检测失效引用/链接文本 | [Word 固定错误文案](precheck.md#broken-references) |

## 队列、方案和批次控制

| 入口 | 说明 |
| --- | --- |
| 文件、文件夹、拖入、清空、右键移除/定位/详情、双击 | [文件队列](files-and-output.md) |
| 全选本页、本页清空、全部取消 | [界面范围](interface.md) |
| 中国eCTD、美国eCTD、我的常用 | [预设](profiles.md#builtins) |
| 配置管理 → 保存为我的常用 | [常用配置](profiles.md#favorite) |
| 配置管理 → 导入配置 (JSON)... | [差异确认](profiles.md#import) |
| 配置管理 → 导出配置 (JSON)... | [便携配置](profiles.md#export) |
| 预检、应用建议、处理建议文件 | [预检与子集](precheck.md#suggestions) |
| 智能处理、全部处理 | [模式含义](precheck.md#modes) |
| 开始处理、确认计划并开始、取消预览 | [处理计划](processing.md#plan) |
| 停止、跳过当前文件、并行终止选中项 | [停止](processing.md#stop)、[跳过](processing.md#skip) |
| 重试失败项 | [重试与新批次](processing.md#retry) |

## 安全中心、设置、日志与关于

| 入口 | 说明 |
| --- | --- |
| PDF 加密与权限保护 | [AES、打开/权限密码、打印/修改/提取/辅助权限、明文元数据、线性化](security.md#encrypt) |
| PDF 解密与权限脱壳 | [独立队列、已知密码、输出后缀、线性化、对象流](security.md#decrypt) |
| 解除数字签名限制、移除表单字典 | [外观、签名与表单边界](security.md#signatures) |
| 全局设置 → 外观主题；顶部深色/浅色 | [主题入口](interface.md#settings) |
| 全局设置 → 默认保存位置、自动打开输出、覆盖原始文件 | [保存规则](files-and-output.md#saving) |
| 全局设置 → 并行数量、每文件超时 | [性能与超时](processing.md#parallel) |
| 日志 → 筛选、搜索、记录详情 | [查看日志](logs.md#view) |
| 日志 → 导出日志、导出预检结果 | [CSV/TXT](logs.md#exports) |
| 日志 → 导出最近批次审计包 | [范围与四文件包](logs.md#audit) |
| 关于 → 版本、检查更新、打开发布页 | [安装与更新；NoUpdate 无更新入口](installation.md) |

本索引不把隐藏工作台或已退役的自动字体修复列为可用功能。维护者核对依据见[手册维护说明](../manual-maintenance.md)。
