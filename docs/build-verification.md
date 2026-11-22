# Windows 可追溯构建与发布自检

在项目根目录使用 **CPython 3.11 / Windows AMD64**。`requirements-release-win311.lock.json` 固定本次运行与构建闭包的精确版本、官方 wheel URL、SHA-256 和 Requires-Dist。当前锁不适用于其它 Python/平台；`requirements.txt` 仍是开发环境的宽版本入口。

```powershell
python -B scripts/build_pyinstaller_safely.py --install-dependency --variant both
```

网络需要本机代理时：

```powershell
python -B scripts/build_pyinstaller_safely.py --install-dependency --proxy http://127.0.0.1:7890
```

`build_pyinstaller.bat` 调用同一脚本。默认构建更新版与 NoUpdate 版；可加 `--variant normal` 或 `--variant no-update`。安装使用 `--require-hashes --no-deps --only-binary`、显式官方 URL 和独立 target，先验证当前平台的传递依赖闭包。省略安装时必须用 `--dependency-path` 指向本项目先前的 `build_pyinstaller/<任务ID>/dependencies`；锁标记、版本、RECORD 与已安装文件哈希不符即停止。不会修改全局 Python 包或持久环境变量。

每次构建使用独立的 `build_pyinstaller/<任务ID>/` 和 `dist/<任务ID>/`，不会覆盖先前发布目录。最终 `BUILD_OUTPUT=` 给出产物目录。真正的分析/打包在新 Python 子进程中进行，锁定目标在导入 packaging/PyInstaller/运行库前生效；实际导入来自目标以外时拒绝。本地解释器、Windows 工具链、环境与时间并非密封，锁定 wheel 不等于位级可复现，也不构成官方或法规资格。

更新锁时只读取应用声明的四个发行包版本，并从公开 PyPI 获取相应 wheel 与传递依赖；输出必须是新文件。先审查差异、备份当前锁，再替换当前锁，不可拿开发环境的全部包清单当发布依赖：

```powershell
python -B scripts/lock_release_dependencies.py new-release-lock.json --pyinstaller-version 6.22.3
```

下载需要代理时生成器同样支持 `--proxy http://127.0.0.1:7890`。生成器不安装包。构建器默认的锁为仓库根目录 JSON，可用 `--lock` 显式选择已审查版本；安装失败只输出固定错误类别，避免泄漏用户配置中的凭据诊断。

构建主进程及 Python 分析子进程安装有界回收保护；无法回收时中止并保留文件，不回退物理删除。每次正式构建前必须出现：

```text
BUILD_GUARD parent=blocked isolated_child=blocked
```

对源码和两个实际 EXE 使用同一 `--verify-release` 入口，每次传入尚不存在的输出目录：

```powershell
$exe = 'C:\path\to\RATools-for-PDF.exe'
$checkDirectory = 'C:\path\to\new-release-check'
$process = Start-Process -FilePath $exe -ArgumentList @('--verify-release', ('"' + $checkDirectory + '"')) -WindowStyle Hidden -Wait -PassThru
$process.ExitCode
Get-Content -LiteralPath (Join-Path $checkDirectory 'release-verification.json')

# 源码：两个入口分别使用新目录
python -B main.py --verify-release new-source-normal-check
python -B main_no_update.py --verify-release new-source-no-update-check
```

验收要求退出码为 `0`、报告 `ok=true`，且 gui/core_pdf/qpdf/image_compression/update_boundary 五项 passed。自检使用 offscreen GUI 初始化、真实 About 按钮路由与关闭；只使用输出目录内的 `settings/settings.ini`。真实合成 PDF 进入核心标题/初始视图规则，保存后检查标题、页数、渲染与输入保全。qpdf 用实际受控 Windows Job API 执行 version、线性化重写和语法检查，冻结时必须解析到包内 qpdf。复用已有图片编码器检查，要求 JPEG 确实降采样、PDF 变小且能重新渲染。

自检观察并阻断当前 Python 进程的 socket 连接/解析/发送。NoUpdate 必须无更新按钮、无 worker、无 updater import，冻结时 updater 也不可被导入。普通版保留真实更新入口，隔离设置内写入当日检查标记以抑制在线更新；不宣称已测试实时 GitHub 更新、所有 OS 网络路径或全部 GUI 交互。所有自检 Python 临时清理受本次输出目录的回收保护。

`--verify-image-compression NEWDIR` 仍可单独使用，输出原有 `verification.json`。请求压缩而 Pillow 缺失时正常管线明确失败，并保留已有输出。

每个成品目录含：

- `release/build-provenance.json`：实际源码 commit、dirty 状态、逐源文件 SHA-256（包括未提交的用户主题）、实际导入运行库版本/文件哈希、构建依赖与 unsigned 签名状态。
- `release/requirements-release-win311.lock.json`：本次精确 wheel 锁。
- `release/sbom.json` 和 `release/licenses/`：根据本次 PyInstaller Analysis 实际输入关联组件及许可证，记录包内 qpdf、CPython 和输入哈希。未归属外部输入使构建停止；当前生成的 base_library.zip 单独校验各成员与本机标准库源码编译结果一致。
- `artifact-hashes.json`：成品全部文件的 SHA-256，自身除外；新增、缺失或改动均使校验失败。哈希证明内容一致，不能认证发布者。

```powershell
python -B -c "from scripts.release_metadata import verify_hash_manifest; import sys; sys.exit(0 if verify_hash_manifest(r'C:\path\to\bundle') else 1)"
```

可选签名仅通过 Windows 当前用户证书存储中的公钥 thumbprint 或引用该 selector 的环境变量。默认只生成 unsigned/pending 计划；不读取私钥文件、密码，不枚举证书，不购买证书、不隐式请求外部时间戳或远程发布：

```powershell
python -B scripts/sign_release.py 'C:\path\to\app.exe'
python -B scripts/sign_release.py 'C:\path\to\app.exe' --certificate-env RATOOLS_RELEASE_CERT_SHA1
# 有明确证书材料且准备执行时才添加 --execute
```

`--execute` 在签名前预检全部输出文件的链接/硬链接边界。只有 Authenticode 签名与独立验证均 Valid 才记录 signed；成功或失败均同步各层签名状态和重新计算成品哈希。多变体部分签名显示 mixed，未供应证书的本次产物保持 unsigned。正式签名、时间戳、法律许可审查、真实业务/外部阅读器资格需另行完成。

软件分层回归入口和资格边界见[测试说明](testing.md)、[资格矩阵](qualification/README.md)；当前成品实测路径与结论随 PD10 任务报告更新，不能由源码通过推定 EXE 已通过。

## Nuitka 发布 ZIP 的体积控制

GitHub 的 Nuitka 发布工作流使用 `requirements-nuitka.constraints.txt` 固定 PyMuPDF 版本。该约束与现有 Windows 发布锁中的 PyMuPDF 版本一致；开发环境继续使用 `requirements.txt`。此约束不等同于锁定全部 Nuitka 编译依赖或核验 wheel 哈希。

`build_nuitka.bat` 通过 `scripts/package_nuitka_release.py stage-pymupdf` 装配运行包：保留 Python 模块、原生库和配置，只排除 `pymupdf/layout/resources/onnx/*.onnx` 与包内 `__pycache__`。当前处理器及其直接 PDF 依赖未使用 layout 模型；PyMuPDF 核心初始化延迟导入 layout。如以后启用相应分析功能，应重新审查这项排除规则。旧输出和冲突插件只通过已校验路径的回收站操作清理，回收失败即停止。

工作流用同一脚本的 `zip` 子命令生成标准 Deflate ZIP，压缩级别为 9，逐文件回读并比较 SHA-256；输出路径必须尚不存在。也可对既有 Nuitka ZIP 使用 `repack` 验证缩包，保留 EXE/DLL 等剩余文件的原始字节：

```powershell
python -B scripts/package_nuitka_release.py zip --source dist/RATools-for-PDF_v0.8.1 --output RATools-for-PDF_v0.8.1.zip
python -B scripts/package_nuitka_release.py repack --source original.zip --output optimized.zip --report optimized-report.json
python -B scripts/run_tests_safely.py test_nuitka_packaging.py --layer integration
```

成功操作不会留下部分 ZIP；失败时独立命名的 `.partial` 文件保留用于诊断。重打包校验通过不代表重新执行了 Nuitka 编译，仍需针对实际解压后的 EXE 执行上述发布自检。


## 2026-11-10 本机最终证据

最终 `BUILD_OUTPUT`：`dist/f313fcee487b49a480d389fc1fafc5ca/`。正常版和 NoUpdate 各有独立目录；构建 exit 0，日志同时记录 `LOCK_VERIFIED packages=14` 与两个 guard 探针 blocked。源码基线为 `55ceb0c0edf01bafadece16ecfb5f97b85476295` 加记录的 dirty PD10/用户主题输入，源码 smoke 的逐文件哈希与实际构建来源完全一致。

| 变体 | 成品目录 | EXE SHA-256 |
| --- | --- | --- |
| normal | `dist/f313fcee487b49a480d389fc1fafc5ca/RATools-for-PDF` | `7a11c4f0b3ecc71cf231b2295832cd8d2a919d1e23b7ec5b685853c8c660ab49` |
| no-update | `dist/f313fcee487b49a480d389fc1fafc5ca/RATools-for-PDF-NoUpdate` | `3f89f5fb247a02dd760fd6c7a3a7bdf6fad2fac71c441b5538004880693f4c84` |

两个最终 EXE 都以隐藏进程运行 `--verify-release`：各 exit 0、ok=true，五项 passed；每个 SBOM 的 885 个实际 Analysis 输入均有明确归属，未使用诊断路径映射。图片从 504960 字节降到正常版 21994、NoUpdate 21996 字节，均为 432×528；包内 qpdf 12.3.2 的受控重写/语法检查都 exit 0。NoUpdate 无按钮/import/可导入 updater/worker，观察到 0 次 Python socket 请求。

最终软件回归 **575/575**，180.985 秒，unit 11 + integration 564，qualification not_selected。两个成品都含实际 smoke 报告、能力概要、软件分层摘要和明确的资格未运行说明；补充后各 296 个文件的完整哈希再次通过，EXE 字节不变。未执行正式签名，当前都是 unsigned。

本机机器证据位于 `backups/product-roadmap-20261109/PD10/`：`release-evidence-summary.json`、`artifact-smoke-summary.json`、`source-release-summary.json`、`build-final-4.log`、`full-release-final.log`。两个 EXE 的原始报告分别位于 `release-normal-final/release-verification.json` 与 `release-no-update-final/release-verification.json`。交付记录见 `.superpowers/sdd/product-development/task-10-report.md`；这些本机路径不表示已远程发版。

## 2026-11-23 Nuitka ZIP 缩包验证

以 GitHub v0.8.1 的实际标准版和 NoUpdate ZIP 为输入，保留所有剩余文件的原始字节，仅排除 16 个 layout ONNX 权重文件及 PyMuPDF 字节码缓存，再使用标准 Deflate 9 压缩。每包 231 个保留文件均经过回读 SHA-256 校验，两个重打包命令都 exit 0。

| 变体 | 原 ZIP 字节数 | 新 ZIP 字节数 | 新体积 | 缩减 |
| --- | ---: | ---: | ---: | ---: |
| standard | 108589995 | 62974602 | 60.06 MiB | 42.01% |
| no-update | 108486022 | 62870792 | 59.96 MiB | 42.05% |

成品位于 `dist/zip-size-20261123-adadeb91/`，保留原有版本文件名。原始 ZIP、配置备份及全部验证日志位于 `backups/zip-size-20261123-adadeb91/`；体积和文件校验明细见 `standard-repack.json`、`no-update-repack.json`。9 项定向打包测试通过，日志为 `nuitka-packaging-tests.log`。

实际 ZIP 中的 PyMuPDF 为 2.1。使用同一装配脚本复制缩包后的实际 PyMuPDF 运行库，并加载包内 qpdf，在本机 Python 3.15 与宿主 Qt 下执行当前源码自检：standard、no-update 都 exit 0、ok=true，gui/core_pdf/qpdf/image_compression/update_boundary 五项 passed、Python 网络尝试 0。报告分别位于 `pruned-source-standard/release-verification.json`、`pruned-source-no-update/release-verification.json`。Nuitka 工作流约束采用现有已审查发布锁的 PyMuPDF 1.28.2；该版本也通过本机源码五项自检，见 `pinned-source-check/release-verification.json`。

v0.8.1 tag 的旧程序入口没有 `--verify-release` 分支；两次实际 EXE 参数检查在 45 秒后结束本次测试进程，没有生成报告，不能记作 EXE 发布自检通过。本次未重新执行 Nuitka 编译；上述源码/实际运行库验证和逐文件字节校验不替代完整重编译或真实业务资格验收。
