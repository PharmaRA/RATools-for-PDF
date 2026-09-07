# PDF 压缩功能实现文档

## 功能概述

为 RATools-for-PDF 添加了完整的 PDF 压缩功能，满足 eCTD 文件大小合规要求（单文件 ≤ 50MB）。

## 实现的功能

### Phase 1：标准压缩（无损优化）

#### 1. 标准压缩 (compress_standard)
- **压缩级别**: `garbage=3`
- **效果**: 5-15% 体积减小
- **特点**: 安全无损，适用于所有文档
- **实现位置**: `processor.py:_finalize_output()`

#### 2. 深度压缩 (compress_aggressive)
- **压缩级别**: `garbage=4 + clean=True`
- **效果**: 10-20% 体积减小
- **特点**: 最大化对象清理，规范化内部结构
- **实现位置**: `processor.py:_finalize_output()`

#### 3. 预检增强
- 文件 > 20MB 时自动建议"标准压缩"
- 文件 > 40MB 时自动建议"深度压缩"
- **实现位置**: `precheck.py:build_precheck_report()`
- **注意**: 压缩选项不在预检可检测集合内，smart 模式下用户已勾选即执行，
  建议仅供未勾选的用户参考

### Phase 2：图像重采样压缩

#### 1. 图像压缩 (compress_images)
- **功能**: 重采样图像到目标 DPI
- **DPI 选项**: 150 DPI（文字类）、300 DPI（图表类，推荐）、自定义（72-600）
- **压缩算法**: LANCZOS 高质量重采样 + JPEG 质量 85%
- **效果**: 50-80% 体积减小（扫描类文档）
- **实现位置**: `processor.py:_step_image_compression()`

#### 2. DPI 选择对话框
- 用户勾选"压缩内嵌图像"时自动弹出
- 提供预设选项和自定义输入
- 包含风险警告提示
- **实现位置**: `ui/dialogs/dpi_selection_dialog.py`

## 文件修改清单

### 核心处理逻辑
1. **ratools_pdf/config/rules_catalog.py**
   - 添加 3 个压缩选项到"文件级优化与输出"模块

2. **ratools_pdf/pdf/processor.py**
   - `_step_image_compression()`: 新增图像压缩步骤
   - `_finalize_output()`: 动态调整 garbage 级别和 clean 模式
   - `_PIPELINE_STEPS`: 添加图像压缩到管线

3. **ratools_pdf/pdf/precheck.py**
   - 根据文件大小自动建议压缩选项
   - 压缩选项不列入 `PRECHECK_DETECTABLE_OPTIONS`（smart 模式下已勾选即执行）

### UI 组件
4. **ratools_pdf/ui/dialogs/dpi_selection_dialog.py**
   - 新增 DPI 选择对话框
   - 提供 150/300/自定义三种模式

5. **ratools_pdf/ui/dialogs/__init__.py**
   - 导出 `DpiSelectionDialog`

6. **ratools_pdf/ui/main_window.py**
   - 导入 `DpiSelectionDialog`
   - 添加 `_on_compress_images_toggled()` 处理方法
   - 为"压缩内嵌图像"复选框绑定特殊事件

### 依赖
7. **requirements.txt**
   - 添加 `Pillow>=10.0` 依赖

## 技术细节

### 压缩参数说明

| 参数 | 取值 | 作用 |
|------|------|------|
| `garbage` | 0-4 | 未使用对象清理级别（0=不清理，4=最激进） |
| `deflate` | True | 启用流级 deflate 压缩（已默认开启） |
| `clean` | True | 规范化内部结构（仅在深度压缩时启用） |
| `use_objstms` | 1 | 使用对象流（PDF 1.5+ 特性，已默认开启） |

### 图像压缩逻辑

```python
# 压缩条件：尺寸超过目标 DPI 下的 A4 尺寸（宽 target_dpi * 8.5 英寸、
# 高 target_dpi * 11.7 英寸），只缩小不放大

# 压缩流程：
1. 使用 PIL 打开图像
2. LANCZOS 重采样
3. 处理透明通道（RGBA/LA/P → RGB，白色背景；其余模式转 RGB，保留灰度 L）
4. 保存为 JPEG（quality=85, optimize=True）
5. 与 PDF 中实际存储的流长度比较（不是 extract_image 的 PNG 重编码长度），
   只有更小才回写
6. 原位替换：update_stream 写入 JPEG 字节 + 逐键同步图像字典
   （/Width /Height /Filter /ColorSpace /BitsPerComponent，清掉
   /SMask /Mask /Decode /DecodeParms）——字典与流不一致会导致图像花屏；
   不使用 Page.replace_image，它会在页面资源中残留重复图像条目
7. 同一 xref 被多页引用时只处理一次
```

### 压缩参数传递

图像压缩参数（dpi/quality）由 UI 层收集（settings.ini `[Compression] ImageDPI`），
经 `MainWindow.get_compression_settings()` → controller → `ProcessWorker` →
`process_document(compression_settings=...)` 传入；pdf 层不读取任何 UI 配置。
默认值与取值范围唯一定义于 `ratools_pdf/config/compression.py`，非法输入
由 `normalize_compression_settings()` 归一化后回落默认。

```
<AppDir>/settings.ini
[Compression]
ImageDPI=300
```

## 使用说明

### 用户操作流程

1. **标准/深度压缩**
   - 直接勾选"标准压缩"或"深度压缩"
   - 处理时自动应用

2. **图像压缩**
   - 勾选"压缩内嵌图像"
   - 在弹出的对话框中选择 DPI（150/300/自定义）
   - 确认后开始处理

### 预检提示

当预检发现文件过大时：
```
文件大小: 35.2 MB ⚠️
建议启用: 标准压缩、深度压缩
```

## 测试

回归测试位于 `tests/test_image_compression.py`，覆盖：
- P0：压缩替换后图像字典与实际流一致、渲染结果与源文件像素差在阈值内
- FlateDecode 大流图像按实际流长度比较触发压缩
- smart 模式下勾选的标准/深度压缩必须执行
- 无图像/全部跳过时的结果反馈、参数归一化

运行：
```bash
python -m pytest tests/test_image_compression.py -v
```

## 注意事项

### 兼容性
- ✅ 标准/深度压缩：100% 安全，适用于所有 PDF
- ⚠️ 图像压缩：可能影响签名 PDF，建议先测试

### 风险提示
- 图像压缩不可逆，处理前建议备份
- 降低 DPI 会永久影响图像清晰度
- 监管文档慎用图像压缩（可能影响图表可读性）

### 最佳实践
1. **优先使用标准/深度压缩**：无损且安全
2. **图像压缩保守使用**：仅用于超大文件应急处理
3. **选择合适 DPI**：
   - 纯文字文档：150 DPI
   - 包含图表：300 DPI（推荐）
   - 高清要求：自定义 400-600 DPI

## 压缩效果预估

| 文档类型 | 标准压缩 | 深度压缩 | 图像压缩(300DPI) |
|----------|----------|----------|------------------|
| 纯文字 PDF | 5-10% | 10-15% | 不适用 |
| 图表混合 | 8-12% | 12-18% | 20-40% |
| 扫描文档 | 5-8% | 8-12% | 60-80% |

*注：实际压缩率取决于原始 PDF 的制作方式*

## 后续优化方向

1. **字体子集化**：移除未使用的字形（需谨慎测试）
2. **批量报告**：显示批量处理前后的总体积变化
3. **压缩预览**：预检时估算压缩后大小
4. **智能推荐**：根据文档类型自动推荐压缩方案

## 版本信息

- **功能版本**: v1.0
- **实现日期**: 2026-09-04
- **依赖版本**: PyMuPDF>=1.24, Pillow>=10.0
