# Windows 构建与压缩能力验收

在项目根目录使用已具备 `requirements.txt` 运行依赖的 Python。构建脚本只安装隔离的 PyInstaller 构建依赖；应用的 PySide6、PyMuPDF、Pillow 和 pywin32 来自当前 Python 环境。

```powershell
python -B scripts/build_pyinstaller_safely.py --install-dependency
```

网络需要本机代理时：

```powershell
python -B scripts/build_pyinstaller_safely.py --install-dependency --proxy http://127.0.0.1:7890
```

`build_pyinstaller.bat` 调用同一脚本。默认构建更新版与 NoUpdate 版；可加 `--variant normal` 或 `--variant no-update`。已安装 PyInstaller 时可省略 `--install-dependency`；也可用 `--dependency-path` 指向上次本项目 `build_pyinstaller/<任务ID>/dependencies`，复用隔离依赖。

每次构建使用独立的 `build_pyinstaller/<任务ID>/` 和 `dist/<任务ID>/`，不会覆盖先前发布目录。最终 `BUILD_OUTPUT=` 给出产物目录，其 `build.json` 记录版本、PyInstaller 版本和两个 EXE 路径。源代码、测试和文档未被打包成完整开发环境。

构建主进程及 Python 分析子进程安装有界回收保护；无法回收时中止并保留文件，不回退物理删除。每次正式构建前必须出现：

```text
BUILD_GUARD parent=blocked isolated_child=blocked
```

发布成品含 Pillow，支持无界面的压缩验证。对每个 EXE 传入一个尚不存在的输出目录：

```powershell
$exe = 'C:\path\to\RATools-for-PDF.exe'
$checkDirectory = 'C:\path\to\new-compression-check'
$process = Start-Process -FilePath $exe -ArgumentList @('--verify-image-compression', ('"' + $checkDirectory + '"')) -WindowStyle Hidden -Wait -PassThru
$process.ExitCode
Get-Content -LiteralPath (Join-Path $checkDirectory 'verification.json')
```

验收要求退出码为 `0`、`verification.json` 的 `ok` 为 `true`。检查会生成合成 `source.pdf` 和实际管线输出 `compressed.pdf`，要求图像确实降采样、PDF 变小且可重新渲染；目录已存在时拒绝复用，避免覆盖旧证据。该入口不打开主界面、不读取用户设置，也不需要外部测试 PDF。

若请求图片压缩但 Pillow 缺失，正常处理管线明确返回失败，已有输出保留。该验收不替代完整 GUI、各项 PDF 规则或真实业务文档验收。
