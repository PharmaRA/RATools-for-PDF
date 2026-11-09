# PD-01 PDF 资格矩阵与证据运行器

PD-01 提供一个版本化 JSON 清单、只读验证服务和显式 CLI。验证器读取清单声明的样本，检查路径是否在样本根内且没有符号链接、重解析点或硬链接绕过，计算输入 SHA-256，重新打开 PDF，提取结构指标，渲染清单指定页面，并运行 qpdf `--check`。它始终在一个全新的证据目录写入 `qualification-report.json`，不会覆盖旧目录或修改输入。

当前清单是 [pd01-manifest-v1.json](pd01-manifest-v1.json)。它列出一个明确标记的 PyMuPDF 合成样本和 Acrobat、Word、LibreOffice、扫描仪、PDF 打印驱动的外部来源行。外部行只有来源要求和稳定 ID；没有提供私有 PDF 时保持 `pending`。Acrobat Reader 和浏览器阅读器行使用 [手工证据模板](manual-reader-evidence-template.md)，没有记录时保持 `pending`。

## 运行

仓库附带 `samples/synthetic-minimal.pdf`，由 PyMuPDF 生成，只有一页公开合成文字和 CJK 文本；清单已登记生成引擎版本和文件 SHA-256。以下命令复制到仓库根目录的 PowerShell 即可运行合成样本并记录外部行 `pending`。加入其他样本时，先把已授权、非敏感样本放在清单的 `samples_root` 下，并填写相对 `path`、工具版本、SHA-256、预期指标和 `render_pages`。每次运行必须使用尚不存在的新目录：

```powershell
$evidence = "backups/product-roadmap-20261109/PD01/evidence-$([DateTime]::UtcNow.ToString('yyyyMMddTHHmmssZ'))"
python -X utf8 -X faulthandler -B scripts/run_qualification.py `
  --manifest docs/qualification/pd01-manifest-v1.json `
  --output $evidence
```

如果本机没有 qpdf，或者只想先记录 PDF 结构，可以显式留下 qpdf 检查为 `pending`：

```powershell
$evidence = "backups/product-roadmap-20261109/PD01/evidence-$([DateTime]::UtcNow.ToString('yyyyMMddTHHmmssZ'))"
python -X utf8 -X faulthandler -B scripts/run_qualification.py `
  --manifest docs/qualification/pd01-manifest-v1.json `
  --output $evidence --skip-qpdf
```

退出码为 `0` 代表清单全部要求实际通过；`1` 代表失败；`2` 代表存在 `pending` 或 `blocked`，也包括 qpdf 告警。qpdf 返回码 3 会保留 `outcome: warning`、诊断摘要哈希和 `pending` 状态，不会被记为完整 `passed`。样本缺失、手工阅读器未记录和外部来源未运行都不会被自动升级为通过。随附清单预期返回 `2`，因为 Windows 成品、真实业务来源和外部阅读器保持独立 `pending`。

报告中的引擎版本、样本哈希、结构指标、渲染像素摘要哈希和逐项状态可作为软件证据。报告不保存文档正文、密码或个人身份；真实来源 PDF、Acrobat/浏览器复核、Windows 成品和外部资格仍需独立完成。

## 清单接口 v1.0

顶层必须包含 `schema_version: "1.0"`、`matrix_version`、`samples_root`、`samples` 和 `matrix`。样本必须有稳定 `id`、相对 `path`（未提供时为 `null`）、`source.type` (`synthetic` 或 `real`)、来源工具/版本、非敏感授权声明、`sha256`、`features`、`processing_options` 和 `expected`。矩阵行必须有稳定 `id`、四种证据层次之一和 `sample_ids`。`status` 由运行器根据实际检查生成；清单中预填的 `pending` 只是初始状态。

支持的 `expected` 指标为 `page_count`、`pdf_version`、`signature_present`、`image_count`、`annotation_count`、`widget_count`、`link_count`、`text_char_count`，按等值检查。至少声明一项指标；未声明指标或未要求渲染页面时保留 `pending`。`render_pages` 是从 1 开始的页码列表，采用 72 DPI 真渲染，不保存正文像素。未知指标记 `blocked`。`features` 和 `processing_options` 描述样本要求与既往处理上下文，PD-01 不执行这些处理，也不据这些标签推断资格通过。

可复用接口：`run_qualification(manifest_path, output_directory, *, samples_root=None, qpdf_enabled=True) -> dict`，位于 `ratools_pdf.services.qualification`。报告 `schema_version` 同为 `1.0`，包含 `engines`、样本 `results/checks`、`summary` 和重新计算的 `matrix`。非自动化层次不会根据 JSON 内写入的 `passed` 自动批准；手工记录需要独立审阅归档。
