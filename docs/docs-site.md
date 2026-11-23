# Starlight 文档网站与本地预览

文档网站位于 `website/`，沿用 RATools-for-Word 的 Astro + Starlight + Pagefind 方案。正文唯一来源仍为 `docs/manual/`；网站提供中文搜索、分组侧栏、页内目录、深浅色主题、移动端菜单、任务首页和源码编辑链接。

## 环境与安装

需要 Node.js **22.23.4 或更新版本**及 npm。版本与 Word 站点一致：Astro 7.4.10、Starlight 0.42.10、sharp 0.35.5，完整依赖由 `website/package-lock.json` 锁定。可选的构建后链接检查使用 Python 3.9+ 标准库，不需要安装 Python 文档依赖。

从 PowerShell 运行：

```powershell
Set-Location C:\Repositories\RATools-for-PDF\website
node --version
node scripts/install-deps.mjs
```

安装入口调用锁文件安装并禁用依赖的安装脚本，使用本站独立 npm 缓存。已有 `node_modules` 会先经路径和链接检查后移入回收站，再重新安装。正常使用不必每次重装依赖。

如 npm 下载因网络失败，可在确认本机代理可用后重试：

```powershell
node scripts/install-deps.mjs --proxy=http://127.0.0.1:7890 --https-proxy=http://127.0.0.1:7890
```

这不会把代理写入项目或用户的 npm 配置。命令中不要填写账号、密码或令牌。

## 方式一：日常编辑预览

在 `website/` 运行：

```powershell
npm run dev -- --host 127.0.0.1 --port 4321
```

浏览器打开：

```text
http://127.0.0.1:4321/RATools-for-PDF/
```

主手册地址为 `http://127.0.0.1:4321/RATools-for-PDF/manual/`。终端显示服务就绪后再访问；如果端口被占用，使用终端实际提示的端口，或明确指定其他端口。

启动时自动同步全部手册，运行期间保存 `docs/manual/*.md` 会触发同步和开发服务器刷新。新增页面时也要更新侧栏；若同步提示不一致，应修正后重新保存或重启。修改截图或侧栏时，运行 `npm run sync`，或重启开发服务。

在启动服务的终端按 **Ctrl+C** 停止。开发模式适合检查内容与布局，**不作为 Pagefind 搜索验收方式**。

## 方式二：带中文搜索的完整预览

在 `website/` 依次运行；前一步成功后再继续：

```powershell
npm run check
npm run build
npm run check:site
npm run preview -- --host 127.0.0.1 --port 4322
```

浏览器打开：

```text
http://127.0.0.1:4322/RATools-for-PDF/
```

`npm run build` 自动同步正文、生成静态 HTML 和 Pagefind 索引，结果在 `website/dist/`。`preview` 只展示最近一次构建；正文修改后需要先停止预览，再重新构建。按 **Ctrl+C** 停止预览。

不要直接双击 `dist/index.html`：`file://` 方式不能正确模拟挂载路径，也不能正常加载完整搜索功能。

## 预览时检查什么

1. 首页的入门按钮、六张任务卡和完整目录是否进入正确章节。
2. 在“全部功能索引”点击某条规则，是否跳转到目标页面的指定小节。
3. 在完整构建预览中搜索“字体”“审计包”“书签”，打开结果并检查目标。
4. 查看“认识主窗口”的截图，以及“书签与链接数据交换”的表格和代码示例。
5. 切换深色/浅色主题，将窗口缩小到手机宽度，检查菜单、页内目录和表格横向滚动。
6. 打开不存在的站内路径，确认显示中文 404 页面；点击页头返回首页。

静态构建与链接检查不能代替浏览器中的搜索交互、视觉检查或真实 PDF 功能验证。

当前锁定版本在存在自定义中文 404 内容时，构建可能提示 `/404` 与更高优先级路由冲突。Starlight 的专用 404 路由仍读取该内容，实际生成 `404.html`；本轮已检查其中文标题及站内链接。该提示不是站点构建失败，本站没有修改第三方依赖来隐藏它。

## 内容和生成文件

| 位置 | 维护方式 |
| --- | --- |
| `docs/manual/*.md` | 正式手册正文，只在此处修改 |
| `docs/assets/screenshots/` | 正文引用的原始截图 |
| `website/src/data/sidebar.json` | 分组侧栏；21 页必须各出现一次 |
| `website/src/pages/index.astro` | 任务首页 |
| `website/src/styles/theme.css` | PDF 站点配色、阅读布局和响应式样式 |
| `website/src/content/docs/404.md` | 中文未找到页面 |
| `website/src/content/docs/manual/` | 自动生成的正文镜像，已忽略，不手工修改 |
| `website/public/manual-assets/` | 自动复制的截图，已忽略 |
| `website/dist/` | 构建产物，已忽略 |

`scripts/sync-docs.mjs` 为正文补标题、摘要和编辑链接，转换 `.md` 路由并保留显式锚点。手册内部链接转换成站内页面；手册外的开发说明转换成 GitHub 仓库链接，避免把 `docs/profiles.md` 错当成 `docs/manual/profiles.md`。截图只复制实际引用的资源。

同步发现已生成但正文中不再存在的页面/资源会停止，列出需处理的路径；按提示将对应生成文件移入回收站后重新同步，不会擅自删除。

## 清理与构建保护

本站 npm 的开发、检查、构建和预览入口都会预加载 `scripts/recycle-guard.mjs`。它将 Node 文件系统清理转交 Windows 回收站，目标限制在本站 `dist`、`.astro`、`node_modules`、`.npm-cache` 内，拒绝未经审查的符号链接、重解析路径和硬链接文件。

构建前将旧 `dist` 移入回收站；Astro 的临时构建目录清理也经同一入口处理。回收失败时停止并保留文件，不降级为物理删除。因此本地受保护工作流目前要求 Windows。不要通过直接调用 Astro 或普通清理命令绕过入口。

## 域名与挂载路径

普通 `npm run build` 的默认配置仍为 `https://PharmaRA.github.io/RATools-for-PDF/`。门户发布使用单独入口 `npm run build:portal`，固定构建到 `https://ratools.qzz.io/pdf/`，并依次完成 Astro 检查、构建和站内校验。这是构建配置，**不代表网站已经发布**。

需要将来挂载到另一个站点，例如 `/pdf/`，可在同一终端设置：

```powershell
$env:DOCS_SITE = "https://ratools.qzz.io"
$env:DOCS_BASE = "/pdf"
npm run build
npm run check:site
npm run preview -- --host 127.0.0.1 --port 4322
```

此时本地访问 `http://127.0.0.1:4322/pdf/`。恢复默认可直接设置：

```powershell
$env:DOCS_SITE = "https://PharmaRA.github.io"
$env:DOCS_BASE = "/RATools-for-PDF"
```

更改环境变量后重新构建；不要把另一个挂载路径生成的 `dist` 直接当作默认路径的产物。

## 接入 RATools 门户

本仓库 `.github/workflows/docs.yml` 在相关 PR/main 修改和手动运行时，使用 Windows + Node 24 + Python 3.11，安装锁定依赖并执行 `npm run build:portal`。只上传 `pdf-manual-preview`，不单独绑定或发布域名。

`PharmaRA/RATools` 门户的 `Deploy Portal & Manuals` 工作流从本仓库 main 拉取源码，在 Windows 构建 `/pdf/`，与 Ubuntu 构建的 Word `/word/` 产物一起装配。首页已经添加 PDF“使用手册”入口。门户 PR/main 推送只构建，手动选择 main 并勾选 `publish` 才发布完整站点。

上线顺序是：先推送并合并 PDF 站点及文档 CI，等 PDF Documentation 通过；再推送并合并门户接入修改，等完整装配通过；最后手动发布门户。不要给 PDF 仓库再次绑定 ratools.qzz.io，也无需为路径 `/pdf/` 添加 DNS。

完整上线、回退和联合预览步骤见门户仓库的 `docs/manuals-deployment.md`。本机路径为 `C:\Repositories\RATools\docs\manuals-deployment.md`；提交后可在门户仓库的 docs 目录中阅读。

本地使用正式构建时运行：

```powershell
npm run build:portal
$env:DOCS_SITE = "https://ratools.qzz.io"
$env:DOCS_BASE = "/pdf"
npm run preview -- --host 127.0.0.1 --port 4322
```

`build:portal` 的设置只传入子进程；预览命令需在当前终端显式设置上述环境变量。访问 `http://127.0.0.1:4322/pdf/`。

## 本轮实际验证记录

2026-11-23，本机 Windows、Node 26.13.1、npm 11.20.0；检查器使用 Python Launcher 可用的 Python 3.15.2。执行目录均为项目 `website/`：

| 实际命令 | 最终退出码 | 结果 |
| --- | --- | --- |
| `node scripts/install-deps.mjs` | 0 | 按锁文件安装 350 个包，依赖安装脚本禁用 |
| `node scripts/sync-docs.mjs` | 0 | 21 页手册、1 张截图、185 个正文链接转换，侧栏完整覆盖 |
| `npm run check` | 0 | Astro 检查 0 errors、0 warnings、0 hints |
| `npm run build` | 0 | 生成 23 个 HTML 页面及 Pagefind 搜索索引；有上文记录的 404 优先级提示 |
| `npm run check:site` | 0 | 21/21 手册页面存在，1421 个本地链接/资源与锚点检查通过，42/42 规则有文档覆盖 |

构建产物位于 `website/dist/`，中文索引元数据位于 `website/dist/pagefind/pagefind-entry.json`，`zh-cn.page_count` 为 22（首页加 21 页手册；404 排除）。受保护构建中的 `.prerender` 临时目录及安装时清理的一份缓存临时文件均已移入系统回收站，可恢复。

只读 review 覆盖同步、链接转换、侧栏、首页/主题、安装/构建包装器、回收目标限制和结果检查器；另以不存在的越界路径验证了同步、Promise、回调三类删除 API，共 9 次拒绝，未创建或删除测试文件。首次检查中的 Astro CLI 路径和 Windows Python 命令别名问题已在包装器中修正；最终结果以上表为准。

未启动本地服务，未运行浏览器搜索/视觉测试，未发布网站或执行 PDF 应用回归。自定义 `DOCS_BASE` 的接口已接入，但本轮构建验证使用默认 `/RATools-for-PDF`。

## 门户接入验证补充（2026-11-23）

后续已执行 `npm run build:portal`，正式域名 `https://ratools.qzz.io`、前缀 `/pdf` 的检查/构建/结果校验均退出 0：Astro 0 errors / 0 warnings / 0 hints，23 个 HTML、1421 个本地链接/资源、42/42 规则覆盖通过。

门户将实际 Word `/word/` 与 PDF `/pdf/` 产物合并，共 55 个 HTML，3206 个本地链接/资源和锚点检查通过，装配专项测试 7/7。完整本机证据与上线步骤记录在 `C:\Repositories\RATools\docs\manuals-deployment.md`；联合产物保留在门户 `.artifacts/portal-pdf-20261123/site/`。

PDF 和门户的 CI 配置已准备并完成 YAML/发布条件静态检查，尚未推送、运行 GitHub Actions 或部署线上。浏览器视觉与搜索交互也未作为本轮已通过项。
