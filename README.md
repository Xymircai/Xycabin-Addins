# WPS Office 加载项与 XyCabin 站点

本项目包含三个相互关联的子项目：两个面向 WPS Office 的 VSTO 加载项，以及一个用于展示与分发这两款加载项的离线站点生成器。

三者共享同一套文档体系与交付门禁，因此以单仓库（monorepo）形式维护。

## 项目组成

| 目录 | 类型 | 宿主 | 功能概述 |
|---|---|---|---|
| `C-00-Shared` | 共用源码与构建实现 | 编译到两个加载项 | 单一公共实现、窗口与网格组件、配置生成、PDF 工具、安装器及验证工具 |
| `C-01-ExcelIterate` | VSTO 加载项 | WPS 表格 / Excel | 公式引用与从属追踪、一键公式封装与数值计算、按列控制显示、文件批量整理与重命名、OCR 文本提取 |
| `C-02-WordUpdater` | VSTO 加载项 | WPS 文字 / Word | 将 Excel 区域粘贴为表格链接、对选中范围的链接批量更新 |
| `C-03-Website·Xycabin` | 站点生成器 | 浏览器 | 生成单文件离线站点，展示两款加载项的介绍、使用指南、版本与下载 |

两个加载项通过 `project.json` 中的 `@shared/` 引用直接编译公共源码，产品脚本只保留传入产品根目录的薄入口，不增加运行时共享 DLL。`verify_shared.py` 校验公共实现唯一、引用有效、入口纯转发及必要依赖一致；`verify_devlog.py` 校验统一文档与门禁。开发入口见 [B0-Entry.md](C-04-Devlog/B0-Entry.md)。

按界面功能定位代码请打开 [B4-Code Index.md](<C-04-Devlog/B4-Code Index.md>)：两个产品各自的 `AC-Original Code` 按选项卡、分组、菜单和末级功能归档，索引提供中文功能树、英文目录、实际回调和共用依赖。新增或迁移功能后运行 `python C-00-Shared/feature_index.py --write` 更新索引，省略 `--write` 执行只读校验。

## 技术栈

- **加载项**：C# 5（由 .NET Framework 自带的 `csc.exe` 编译），VSTO vstov4 清单，无 Visual Studio 工程文件；独立 PDF 工具按锁定版本和 SHA-256 恢复依赖。
- **构建与验证**：Python 3，仅使用标准库（无 `requirements.txt`）。
- **站点**：由 Python 模板渲染为单个 HTML 文件，内嵌 CSS、JavaScript 与图片资源，可离线打开。

## 目录结构

```
.
├── C-00-Shared/          公共源码、资源与构建实现（两个加载项直接引用）
├── C-01-ExcelIterate/
│   ├── AA-Program/        配置器、启动器与安装包
│   ├── AB-Graphic/        界面截图、功能区图标与示意 SVG
│   ├── AC-Original Code/  按菜单层级归类的功能入口、算法和弹窗
│   ├── src/               加载框架、生成文件、配置、安装器与测试工具
│   ├── build/             构建产物、缓存及验证证据（未入库）
│   └── lib/               微软版权互操作程序集（未入库）
├── C-02-WordUpdater/      结构同上
├── C-03-Website·Xycabin/
│   ├── AA-Program/        站点配置器、启动器与生成的网页
│   └── src/               模板、控制器、资产
├── C-04-Devlog/           唯一文档维护位置（B0、A-UserGuide、B1、B2、B3、B4）
├── verify_devlog.py       文档结构、概览、审计与规则注册表校验
├── verify_shared.py       公共实现、源引用和产品薄入口校验
└── AGENTS.md              仅指向 B0-Entry.md
```

## 构建

### 环境要求

1. Windows 10 / 11
2. .NET Framework 4.x（提供 `csc.exe`，系统自带）
3. Python 3.10 或更高版本
4. WPS Office —— 仅在运行与验收时需要，编译不需要

### 补回互操作程序集

仓库不包含 `lib/` 下的微软版权程序集（`Microsoft.Office.Tools.*.dll`、`Microsoft.VisualStudio.Tools.Applications.Runtime.dll` 等），因为其再分发受微软许可条款约束。首次构建前，请从本机的 VSTO 运行时或 Visual Studio Tools for Office PIA 目录将这些程序集复制到对应项目的 `lib/` 目录。

缺少任一会导致编译报错，错误信息会指出具体文件名。

### 编译

请保留 `C-00-Shared`、相应加载项、`C-03-Website·Xycabin` 和 `C-04-Devlog` 的同级目录关系。公共源码迁移后，单独下载一个加载项目录不足以构建。

```bash
# 仅编译 DLL 与依赖，不签名、不封装安装包
python C-01-ExcelIterate/build.py
python C-02-WordUpdater/build.py
```

### 完整重建

```bash
python C-01-ExcelIterate/rebuild.py --background --mode full
python C-02-WordUpdater/rebuild.py --background --mode full
python "C-03-Website·Xycabin/rebuild.py" --background --mode full
```

或双击各项目 `AA-Program/` 下的 `一键重建并打包.bat`。

完整重建依次执行：写入配置 → 编译 → 生成并签名清单 → 封装安装包 → 校验。`--mode incremental` 会复用 `build/incremental-cache/` 中的编译缓存。

`C-03-Website·Xycabin/rebuild.py` 是唯一的联合交付入口：它先构建并验证两个加载项，再把两个安装包内嵌进单文件网页，最后事务发布三件产物。只发布其中一个产品时，仍应走联合入口，否则网页内嵌的安装包会与独立附件不一致。单独运行 `AA-Program/C3-XycabinSite_Configure.py` 只会在同目录生成不含安装包的预览页，不能作为发布附件。

### 关于代码签名

完整重建的最后阶段会使用 Authenticode 证书为 DLL、应用清单与部署清单签名。签名身份由各项目 `AA-Program` 下的配置器决定：Excel-Iterate 为 `C1-Exceliterate_Configure.py`，WordUpdater 为 `C2-Wordupdater_Configure.py`，其中的 `SIGNING_CERT_THUMBPRINT` 是唯一来源。

**自行构建时，该指纹指向的证书在你的机器上并不存在，签名步骤会失败。** 有两种处理方式：

1. 在对应项目的配置器中将 `SIGNING_CERT_THUMBPRINT` 改为你自有证书的指纹，并确保证书位于当前用户的个人证书存储区且带私钥；
2. 跳过完整重建，改用 `build.py` 编译出 DLL，再手动安装 VSTO 加载项用于开发调试。

原项目的发布证书为自签名证书，未加入任何公共信任根。因此安装官方构建的安装包时，Windows 会提示「未知的发行者」，这属于预期行为。

## 未纳入版本管理的内容

| 路径 | 原因 |
|---|---|
| `build/`、`stage/` | 构建快照与中间产物，体积可达数百 MB，可由源码重建 |
| `.workbuddy/`、`.ai-memory/` | AI 辅助开发的本地工作记录，含本机路径与过程性内容 |
| `**/lib/*.dll` | 微软版权程序集，再分发受许可限制 |
| `AA-Program/*_Setup.exe` | 发布产物，通过 Releases 分发 |
| `C-03-Website·Xycabin/AA-Program/XyCabin.html` | 由模板渲染生成的产物 |
| `C-03-Website·Xycabin/AA-Program/index.html` | 同一产物在本地预览时的文件名 |
| `*.cer`、`*.pfx`、`*.snk` | 证书与私钥材料 |

## 新增功能、提交与 Release 发布

以下命令使用 **Windows PowerShell**，从仓库任意子目录开始。Python 至少 3.10；Git、GitHub CLI 需已安装，gh 登录账号需有仓库写入和 Release 更新权限。

### 准备终端

```powershell
$TaskRoot = git rev-parse --show-toplevel
if ($LASTEXITCODE -ne 0) { throw "请先进入本 Git 仓库" }
Set-Location -LiteralPath $TaskRoot
$TaskPython = "python" # 可替换为可用 Python 解释器的绝对路径
$RepoName = "Xymircai/Xycabin-Addins"
$ReleaseTag = "v0.9.9.3"
$env:PYTHONUTF8 = "1"
$env:PYTHONIOENCODING = "utf-8"
$env:PYTHONDONTWRITEBYTECODE = "1"
function Invoke-Checked {
    param([string]$Program, [string[]]$Arguments)
    & $Program @Arguments
    if ($LASTEXITCODE -ne 0) { throw "$Program 退出码 $LASTEXITCODE，停止后续步骤" }
}
Invoke-Checked -Program $TaskPython -Arguments @("--version")
Invoke-Checked -Program "git" -Arguments @("--version")
Invoke-Checked -Program "gh" -Arguments @("auth", "status", "--hostname", "github.com")
```

原证书与私钥、VSTO 依赖仍按上文准备。发布沿用固定身份，不能为了绕过签名失败而替换证书。

### 新增或调整功能

分类约束见 [B3 的 R-30](<C-04-Devlog/B3-Gate Control.md#6-功能源码分类与代码索引>)，代码定位见 [B4-Code Index](<C-04-Devlog/B4-Code Index.md>)。

1. 在对应 AA-Program 配置器的 BUTTONS / EXTRA_TABS 定义菜单与稳定 ID。
2. 在 project.json 的 code_layout 为新增节点 ID 登记英文 PascalCase 目录名。目录由“选项卡→分组→菜单→末级功能”的真实 ID 路径推导，不能仅修改 features.directory 蒙混通过；显示名称或同层排序变化不要求移动源码。
3. 在 AC-Original Code 对应叶子实现真实回调和专属代码；父菜单目录只放本菜单图标接线。共用业务源码放在使用功能的最近公共父级 _Shared，跨产品代码集中在 C-00-Shared 并登记 source-manifest.json。
4. 同步 sources 编译清单和 features 的 directory、entry、implementation、tests。AC 中每份源码均须归属功能；加载框架和生成文件保留在白名单位置，新增业务实现不得写回 src/addin。基础设施白名单变更需同步检查器并评审。
5. 显式生成 B4，再运行只读门禁和正反例：

```powershell
Invoke-Checked -Program $TaskPython -Arguments @("C-00-Shared/feature_index.py", "--write")
Invoke-Checked -Program $TaskPython -Arguments @("C-00-Shared/feature_index.py")
Invoke-Checked -Program $TaskPython -Arguments @("C-00-Shared/feature_index.py", "--project-root", "C-01-ExcelIterate")
Invoke-Checked -Program $TaskPython -Arguments @("C-00-Shared/feature_index.py", "--self-test")
Invoke-Checked -Program $TaskPython -Arguments @("verify_devlog.py")
Invoke-Checked -Program $TaskPython -Arguments @("verify_shared.py")
```

核验 Word 时将 --project-root 换为 C-02-WordUpdater。独立产品只核对自身和 B4 对应章节；重写整份 B4 需要两个产品完整源码。构建只读检查，不代替开发者修目录或索引。R-30 在单项目 rebuild/build/validate、联合构建和验证、批处理调用链中强制执行，完整与增量模式均不跳过。

### 联合重建与验证

先固定非审计文档，再执行以下两步，全部通过后才提交和发布。联合流程使用同一网页快照构建两个加载项，并将两个 EXE 与内含同一两包的 HTML 一起完成本地事务交付。

```powershell
Invoke-Checked -Program $TaskPython -Arguments @("C-03-Website·Xycabin/rebuild.py", "--background", "--mode", "full")
Invoke-Checked -Program $TaskPython -Arguments @("C-03-Website·Xycabin/validate.py", "--background", "--mode", "full")
```

重建期间不要编辑文档、共享源码或内嵌资源；发现变化应等待本轮结束、修正后重跑。后台回归、签名、五载荷、内嵌资源及网页中的两包均须通过。实际 WPS 点击与安装效果单列验收，构建不自动安装或关闭宿主。

### 分批提交和推送

```powershell
git status --short
git diff --stat
# 按本轮实际修改选择路径，勿混入无关改动。
git add -- "C-00-Shared" "C-01-ExcelIterate" "C-02-WordUpdater" "C-03-Website·Xycabin" "C-04-Devlog" "verify_shared.py" "verify_devlog.py"
git diff --cached --check
git diff --cached --name-status
git diff --cached
git commit -m "feat(shared): 强制校验功能源码分类与索引" -m "说明原因、影响范围和本轮真实验证。"
git add -- "README.md"
git diff --cached --check
git diff --cached
git commit -m "docs: 完善分类门禁与 PowerShell 发布步骤"
git fetch origin
git log --oneline --left-right origin/main...main
# 远端有新增时保留既有历史并合并；有冲突先解决并重新验证。
git merge origin/main
git push origin main
$SourceCommit = git rev-parse HEAD
git status --short
git log --oneline -8
git ls-remote origin refs/heads/main
gh api "repos/$RepoName/commits/main" --jq '.sha[0:7] + " " + .commit.message'
```

最后两条用于确认远端指针已指向 `$SourceCommit`：`git ls-remote` 比对哈希，`gh api` 读取 GitHub 网页同源的数据，两者一致才说明他人能在网页看到本轮代码。

| 批次 | 典型范围 | 提交信息前缀 |
|---|---|---|
| 1 | 共享源码修复（`C-00-Shared`） | `fix(shared): …` |
| 2 | 单个产品的行为或界面修复 | `fix(excel): …` / `fix(word): …` |
| 3 | 生成文件时间戳（`*.g.cs`） | `chore(build): …` |
| 4 | 文档、代码索引与发布证据 | `docs: …` |

一个提交只承载一个语义单元；`*.g.cs` 的时间戳变化与真实修复分开提交，回溯时可排除噪声。每批提交前先 `git diff --cached` 复核范围。

任一步失败均停止，冲突和门禁未解决时不推送。提交信息采用“类型(可选范围): 摘要”，正文说明原因、影响和验证；原有未推送提交随 main 一并推送，不强推或改写历史。长期分支是 main，历史 restructure/20261019 分支不合入。

不用 git add -A / git add . 代替范围审阅。build、stage、两个 EXE、生成 HTML、lib DLL、证书与私钥继续被忽略。git push 上传源码和文档，Release 附件需要单独上传。

### 常见阻塞与处置

| 现象 | 原因 | 处置 |
|---|---|---|
| `TrackerProbe.exe` 无任何输出即以 `0xE0434352` 退出，事件日志报 `Exception.ToString()` 失败 | `project.json` 被保存为带 UTF-8 BOM；`DataContractJsonSerializer` 把 BOM 当正文抛 `XmlException` | 已在 `SourcePaths.ReadManifest` 剥离 BOM。新增清单读取点须同样处理，或把 `project.json` 保存为无 BOM |
| `Feature099Probe` 报 `actual preview button has text N …` | 新增或改动对话框后，控件超出默认尺寸，被 `AutoScroll` 区域裁出可视范围，探针采样到的文字像素为 0 | 调整 `DialogSizes.cs` 的默认尺寸或窗体行比例，使全部控件可见；不要为通过而放宽探针 |
| 日志出现 `[safe-delete][SAFE_DELETE_BULK_CONFIRM_REQUIRED]`，联合重建中断 | 运行环境对批量删除按轮计数（常见阈值 50），清理 `build/stage-*` 时超限 | 改在系统终端直接执行联合重建；或先清理 `build/stage-*` 残留目录再重跑 |
| `git push` 长时间无响应，`api.github.com` TLS 超时 | 本机代理不可用 | 换用可用网络或修复代理后再推送；未推送前不要更换 Release 附件，避免源码与附件错位 |

### 备份和更新现有 Release

沿用 0.9.9.3，不移动已有标签。先确认 GitHub main 已显示 $SourceCommit，再备份旧说明和三个附件。每轮使用新目录，不以新附件覆盖旧备份。

```powershell
$ReleaseFiles = @(
    "C-01-ExcelIterate/AA-Program/Excel_Iterate_Setup.exe",
    "C-02-WordUpdater/AA-Program/WordUpdater_Setup.exe",
    "C-03-Website·Xycabin/AA-Program/XyCabin.html"
)
$ReleaseAudit = Join-Path "build" ("release-" + (Get-Date -Format "yyyyMMdd-HHmmss"))
$ReleaseBackup = Join-Path $ReleaseAudit "backup"
$ReleaseVerify = Join-Path $ReleaseAudit "downloaded"
New-Item -ItemType Directory -Path $ReleaseBackup, $ReleaseVerify -Force | Out-Null
$BeforeJson = & gh release view $ReleaseTag --repo $RepoName --json name,body,tagName,url,assets
if ($LASTEXITCODE -ne 0) { throw "无法读取 Release，停止发布" }
$BeforeJson | Set-Content -LiteralPath (Join-Path $ReleaseAudit "release-before.json") -Encoding utf8
Invoke-Checked -Program "gh" -Arguments @("release", "download", $ReleaseTag, "--repo", $RepoName, "--dir", $ReleaseBackup, "--pattern", "Excel_Iterate_Setup.exe", "--pattern", "WordUpdater_Setup.exe", "--pattern", "XyCabin.html")
$ReleaseFiles | ForEach-Object { Get-FileHash -LiteralPath $_ -Algorithm SHA256 }
```

将本轮变化、真实验证、三附件 SHA-256 与以下源码链接写入 $ReleaseAudit/release-notes.md，然后同批更新：

```powershell
"https://github.com/$RepoName/tree/$SourceCommit"
("https://github.com/" + $RepoName + "/archive/" + $SourceCommit + ".zip")
Invoke-Checked -Program "gh" -Arguments (@("release", "upload", $ReleaseTag, "--repo", $RepoName, "--clobber") + $ReleaseFiles)
Invoke-Checked -Program "gh" -Arguments @("release", "edit", $ReleaseTag, "--repo", $RepoName, "--notes-file", (Join-Path $ReleaseAudit "release-notes.md"))
Invoke-Checked -Program "gh" -Arguments @("release", "download", $ReleaseTag, "--repo", $RepoName, "--dir", $ReleaseVerify, "--pattern", "Excel_Iterate_Setup.exe", "--pattern", "WordUpdater_Setup.exe", "--pattern", "XyCabin.html")
foreach ($ReleaseFile in $ReleaseFiles) {
    $DownloadedFile = Join-Path $ReleaseVerify (Split-Path -Leaf $ReleaseFile)
    $LocalDigest = (Get-FileHash -LiteralPath $ReleaseFile -Algorithm SHA256).Hash
    $RemoteDigest = (Get-FileHash -LiteralPath $DownloadedFile -Algorithm SHA256).Hash
    if ($LocalDigest -ne $RemoteDigest) { throw "Release 附件不一致：$ReleaseFile" }
}
```

三文件必须同批重建和验证，HTML 内嵌的两个 EXE 与独立附件须逐字节一致。GitHub 的附件上传不具有三文件原子性；部分失败时使用 backup 原文件以同名 --clobber 恢复三个附件，并使用 release-before.json 的原 name/body 恢复说明，再核对原摘要。保留失败记录，修复后重跑，不能将部分成功标为发布成功。

已有标签与 GitHub 自动生成的“Source code”附件仍对应历史提交，本轮源码使用发布说明中的提交链接和 ZIP。相同版本但字节更新的用户需手动下载并重装；仅在明确指定新四段版本号时升版、新建标签和 Release。

### 回填发布证据

重下载核对通过后，在 B1“本轮交付证据”记录实现提交、三附件摘要和验收边界。审计区不参与内嵌正文，回填不需重建安装包。

```powershell
Invoke-Checked -Program $TaskPython -Arguments @("verify_devlog.py")
git add -- "C-04-Devlog/B1-Project-Files.md"
git diff --cached --check
git diff --cached
git commit -m "docs: 记录 GitHub 与 Release 发布验收"
git push origin main
git status --short
git rev-parse HEAD
git ls-remote origin refs/heads/main
gh release view $ReleaseTag --repo $RepoName
```

最终工作区应干净，本地 HEAD 与远端 main 一致，GitHub 网页可见 B3/B4/README，三附件摘要与本地正式产物相同。审计提交可晚于构建源码提交；发布说明保留实际构建对应的实现提交号。

## 免责声明

本项目为个人开发作品，**与金山办公（WPS Office 的开发方）无隶属或背书关系**，亦非 WPS Office 官方组件。

加载项通过 VSTO 接口与宿主进程交互，会读取并修改当前文档内容。请在测试文档中先行验证，并自行承担使用风险。作者不对因使用本软件导致的数据丢失、文档损坏或业务中断承担责任。

## 许可证

本项目采用 [Mozilla Public License 2.0](LICENSE)（MPL-2.0）授权。

MPL-2.0 是**文件级** copyleft 许可证，其边界与常见许可证不同，使用前请留意：

- 允许商业使用、修改与再分发，也允许收费分发。
- 你修改过的**源文件**，在分发时必须继续以 MPL-2.0 公开。
- 但 MPL-2.0 不要求你新增的文件或与之组合的其他代码开源，可以与其他许可证（含专有）代码组合成一个更大的作品。
- 含明确的专利授权条款；不授予商标使用权。
- 本项目未声明「Incompatible With Secondary Licenses」，因此按 MPL-2.0 默认规则，可与 GPL-2.0 或更高版本进行二次组合。

源文件的许可声明以仓库根目录的 `LICENSE` 文件为准，未在每个文件头部逐一附加 Exhibit A 通知。若你希望在自己维护的分支中补全文件头，可参考 `LICENSE` 末尾的 Exhibit A 模板。

## 问题反馈

请通过本仓库的 Issues 提交缺陷报告或功能建议。提交时请说明 WPS Office 版本、Windows 版本，以及可复现的操作步骤。

PDF 工具适配 0.9.9.3：使用与已知问题 PDF-001 见 [统一使用说明](C-04-Devlog/A-UserGuide.md)，实现边界见 [共享开发经验](C-04-Devlog/B2-DevExperience.md)。PDF-001 待复现、未修复。
