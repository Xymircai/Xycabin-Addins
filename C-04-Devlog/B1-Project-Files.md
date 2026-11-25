**三项目档案与版本**

最后更新：2026-11-25 ｜ 当前版本：**0.9.9.8** ｜ 交付物：`Excel_Iterate_Setup.exe`、`WordUpdater_Setup.exe`、`XyCabin.html`

# 一. Agent及开发者接手概览

本仓库维护三个输出：WPS 文字加载项 WordUpdater、WPS 表格加载项 Excel-Iterate，以及两者共用的单文件离线网站 XyCabin。文档分为接手入口、使用说明、项目档案、开发经验、门禁、代码索引六类。加载项主选项卡的“扩展→Agent扩展”显示 Agent-Evo，提供概览/全文、网站和自定配置。当前源码与本地交付版本为 0.9.9.8。

## 1. 目录与阅读顺序

| 文档                                         | 内容                               |
| ------------------------------------------ | -------------------------------- |
| [B0-Entry.md](B0-Entry.md)                 | 入口与阅读顺序、项目关系、交付边界                |
| [A-UserGuide.md](A-UserGuide.md)           | 三产品使用说明                          |
| B1-Project-Files.md                        | 本章与三个项目档案：定位、版本摘要、命令按钮、源码地图、当前状态 |
| [B2-DevExperience.md](B2-DevExperience.md) | 共享开发经验与三个专属经验                    |
| [B3-Spec Coding.md（B3）](<B3-Spec Coding.md>) | 编码、构建、验证、交付及网站门禁                        |
| [B4-Code Index.md](<B4-Code Index.md>) | 两个加载项全部末级功能的目录树、入口、核心实现与共用依赖 |

Agent 修改后的正式成果生成入口、完整验证方式及交付证据要求，统一见 [B3 的 R-17](<B3-Spec Coding.md>)，交付前先核对该约定。

人工专项 H-01「冗余清理」与 H-02「代码优化」的触发条件见 B3。

## 2. 项目与交付物

| 项目            | 程序集           | 交付物                     |
| ------------- | ------------- | ----------------------- |
| Excel-Iterate | Excel_Iterate | Excel_Iterate_Setup.exe |
| WordUpdater   | WordUpdater   | WordUpdater_Setup.exe   |
| XyCabin       | —             | XyCabin.html            |

## 3. 源码与资源目录

三个项目统一为 `AA-Program`（配置器、启动器、正式产物）与 `AB-Graphic`（界面截图、功能区图标、示意 SVG）；功能源码在产品 `AC-Original Code/`，按选项卡、分组和菜单叶子归档；加载框架、生成源码及工具仍在 `src/`。公共实现只在同级 `C-00-Shared`，依赖在 `lib/`，构建中间物与验证证据在各产品 `build/`。完整功能导航只在 [B4-Code Index](<B4-Code Index.md>) 维护。文档只在仓库根 `C-04-Devlog` 维护，项目内不再保留副本。

## 4. 项目状态维护

每次修改都以最新状态覆写各产品的接手摘要、交付回执和下一阶段目标。接手摘要按【现场基线】【目标与验收】【执行规范】【起手动作】四节组织，每个产品 600～1,000 字。只保留推进下一步所需的源码、配置和最新回执入口，把计数范围写进指标名。正式构建前固定接手正文，取得结果后按 R-17 回填本次实际构建项目的「本轮交付证据」。版本表按版本维护公开摘要。

# 二. Excel-Iterate 项目档案

## 1. 定位与输出

WPS表格VSTO加载项：引用/从属追踪、公式、筛选、链接、目录、文件整理/命名、文件名称/OCR、表头匹配合并及合并抵消。命名空间/程序集 `Excel_Iterate`，唯一安装器[AA-Program/Excel\_Iterate\_Setup.exe](../C-01-ExcelIterate/AA-Program/Excel_Iterate_Setup.exe)，内嵌五载荷；用户需完整VSTO运行库及官方授权。Agent-Evo同一DLL提供网站、离线文档、四项自定配置和需求记录，实验OCR占位灰色。

AA-Program放配置器、批处理与正式程序，AB-Graphic放界面截图、功能区图标和示意SVG；文档统一在仓库根C-04-Devlog维护；build是工具、中间产物和验证证据。C-01/C-02直接编译引用C-00-Shared公共实现，C-03是网站唯一源码及联合交付入口，不建立网站副本。

## 2. 版本摘要

版本来源与归并规则见[门禁](<B3-Spec Coding.md#1-文档编号与经验更新>)，当前为 0.9.9.8，沿用项目统一共享发布者（CN=Xymircai WPS Addins Dev，两产品同指纹）。公开配置见 C-00-Shared/signing/signing.json。证书缺失时由构建链从匹配的本机 PFX 导入当前用户个人证书库。

| 版本      | 重要变更                                                                                                                                                                                                 |
| ------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| 0.9.5.0 | 1. 首次提供WPS表格加载项，支持引用与从属追踪；2. 采用VSTO构建，提供单文件安装包。                                                                                                                                                      |
| 0.9.6.0 | 1. 完善引用与从属追踪；2. 统一窗口和图标；3. 增加运行库检查与固定证书签名，提升安装可靠性。                                                                                                                                                   |
| 0.9.7.0 | 1. 新增公式封装与计算、筛选增强、链接助手和工作表目录；2. 改善选区兼容与目录显示；3. 新增Agent-Evo选项卡、离线文档预览及网站入口。                                                                                                                           |
| 0.9.8.0 | 1. 新增文件整理与清单改名、链接清单保存替换，支持取消还原与确认清理；2. 改善筛选与反馈；3. Agent-Evo提供离线文档及树形尺寸设置；4. XyCabin提供离线下载、音乐和独立音量。                                                                                                   |
| 0.9.9.0 | 1. 文件整理、清单命名及多前缀OCR，按行断链、整簿恢复与编辑保护；2. 六列预览、布局回缩，即时配置及控件缩放，表头单行居中；3. 追踪引用新增显示错误来源：沿引用链定位错误源头格，可跳转标黄；4. 审计工具箱新增合并抵消：往来按单位对冲、坏账按比例冲回，收入成本逐行算未实现利润，列映射+分录预览+备份还原；5. 修复追踪弹窗锁死宿主：去除工具窗样式并对齐链接助手，根治抢焦点/失激活。 |
| 0.9.9.1 | 1. 合并抵消改净额法取数：资产端=债权方资产−负债、负债端=债务方负债−资产，债权方按净额较大方动态判定，净额保留真实符号；2. 公司对按实际往来自动配对，对方单位精确匹配，取消前缀匹配与别名归一；3. 差异表扩至15列，展示四类原始余额与净额过程，结果表跨表引用抵消金额；4. 修复未勾写数值时结果表金额引用错行。                                      |
| 0.9.9.2 | 1. 文档收敛为 C-04-Devlog 五份（B0 入口、A 使用说明、B1 档案、B2 经验、B3 门禁），项目内不留副本；2. 目录重构为 AA-Program/AB-Graphic，配置器更名 C1；3. 内嵌文档改按 C-04-Devlog 新路径读取，Agent-Evo 概览按章节标题与首段生成；4. 业务功能无变更。                               |
| 0.9.9.3 | 1. 开发/用户全文分别整份内嵌B2与A-UserGuide，导出文件名同源；2. 链接助手说明名称外链，按来源及作用域清理未使用名称后原生断链，保留整簿恢复；3. 断链指令改为“断开链接”，旧文本提示改填；4. 名称依赖预检区分A1地址与真实名称，保存后核对并准确报告失败状态；5. 名称外链状态改为“仅名称管理器定义引用（X 个名称）”。 |
| 0.9.9.8 | 1. 链接助手识别来源缺失的外链，支持替换到尚未就位的文件并保留外链公式，保存后核对新来源，支持按整簿备份撤销；2. 工作表批量操作按同名表预览文字和数值替换，支持格式设置及窗口内撤销；3. 精简交接与开发经验，更新内嵌说明。 |


## 3. 当前命令按钮概览

本表以配置器 BUTTONS 与 EXTRA_TABS 为准，保留顶层功能概览。逐项源码路径、回调与共用依赖统一见 [B4-Code Index](<B4-Code Index.md>)；Agent-Evo 默认隐藏，显示名称和排序可由用户自定。

| 命令按钮 | 选项卡与分组 | 子选项 | 功能摘要 |
| --- | --- | --- | --- |
| 追踪引用 | Excel-Iterate·数据 | — | 解析当前单元格公式，逐项展开被引用单元格的详情；错误行红字，选中后可「显示错误来源」回溯定位源头格并跳转 |
| 追踪从属 | Excel-Iterate·数据 | — | 遍历所有已打开工作簿，找出哪些单元格的公式引用了当前单元格 |
| 筛选增强 | Excel-Iterate·数据 | 按列筛选、快捷高筛 | 按活动单元格所在表格的表头选择显示列 |
| 审计工具箱 | Excel-Iterate·数据 | 批量迁移数据、表头匹配、合并抵消（往来抵消、收入成本抵消） | 按表头列名批量匹配多个工作簿，把列数据合并到目标工作簿；往来抵消按列映射生成抵消分录，收入成本抵消按三维聚合生成矩阵式底稿（借贷恒等式保证平衡），两者均写入前备份、可一键还原 |
| 一键转换 | Excel-Iterate·操作 | 文本提取与过滤、批量插入与删截、工作表批量操作 | 转换选区内容，预览并替换同名表的文字或底层数值，批量设置格式。 |
| 一键公式 | Excel-Iterate·操作 | 进行公式封装（IFERROR、ROUND、ABS、自定义函数）、移除公式封装、数值计算、相对-绝对转换、公式-数值转换 | 在预填参数或确认窗口应用，可撤销最新对应批次。 |
| 文件操作 | Excel-Iterate·操作 | 批量处理（文件批量整理、文件批量命名、文件提取名称）、文件转换与 PDF 工具（当前文档转 PDF、批量文件转 PDF、PDF 添加水印、PDF 拆分、PDF 合并） | 文件批量整理、命名、名称提取；当前文档或批量转 PDF，以及 PDF 水印、拆分和合并。 |
| 链接助手 | Excel-Iterate·操作 | — | 扫描普通及名称外链，支持替换至暂缺的新路径，断链先核对名称依赖，涉及缺失来源或目标的替换可整簿恢复 |
| 生成目录 | Excel-Iterate·操作 | — | 链接全部工作表，隐藏表点击后显示并跳转 |
| 问题反馈 | Excel-Iterate·反馈 | — | 有BUG或新增功能请发送邮箱（打开 QQ 邮箱网页版） |
| Agent扩展 | Excel-Iterate·扩展 | — | 显示或隐藏 Agent-Evo 选项卡（开关，按用户记住选择） |
| 调用API | Agent-Evo·详情 | 智能云-API（百度智能云、阿里智能云）、Agent-API（国内Agent（WorkBuddy、TRAE、Qoder、DeepSeek、豆包、Kimi）、国外Agent（ChatGPT、Claude、Gemini）） | 四级嵌套菜单，打开对应云控制台或 AI 网站（无 screentip，url 按钮） |
| 开发流程 | Agent-Evo·详情 | 开发预览、用户预览、XyCabin | 打开内嵌开发/用户文档与网站首页 |
| 自定配置 | Agent-Evo·操作 | 自定排序、自定名称、自定格式、交互弹窗 | Agent 弹窗的自定义项 |
| 实验计划 | Agent-Evo·操作 | OCR识别、需求记录 | 实验性功能入口 |

## 4. 源码地图与运行入口

| 子系统 | 位置与入口 |
| --- | --- |
| 功能实现 | AC-Original Code；完整定位见 [B4-Code Index](<B4-Code Index.md>)，机器映射在 project.json 的 features |
| 加载与生成 | src/addin 的 ThisAddIn、Connect、AssemblyInfo 及生成文件；AA-Program 配置器和 src/configuration |
| 跨产品共用 | C-00-Shared；产品只保留功能薄入口，经 @shared/ 直接编译引用共用实现 |
| 构建与验证 | 根目录构建脚本；src/installer、src/pdf、src/tools、src/testharness；父目录 verify_shared.py |
| 依赖与证据 | lib 为必要编译依赖；build 为中间物、候选包、缓存与验证证据 |

批量迁移数据位于审计工具箱，支持选择源区域、批量添加文件或文件夹、勾选目标、逐项预检及输出校验后的独立副本。入口及专属实现位于 `C-01-ExcelIterate/AC-Original Code/ExcelIterate/Data/AuditToolbox/DataMigration/`，验证入口为 DataMigrationProbe 与 DataMigrationLiveProbe；详细规则见 A 与 B2。 内部分为配置快照/纯计划、文件名索引、预设编解码、工作表快照与恢复实现，窗体不再维护这些实现。


单项目交付运行本产品 AA-Program 下的「一键重建并打包.bat」。源码诊断可运行 python rebuild.py --background --mode full，再运行 python validate.py --background --site-html build/site-embedded.snapshot.html。候选包使用 --installer-output 写入 build，并通过 validate.py 的 --installer 指定。

## 5. 当前状态与交付

### (1). 接手摘要

#### a. 【现场基线】

在仓库根的 `C-01-ExcelIterate` 工作。当前分支为 `main`，换链实现基准为 `7ffb77c95fbe`，以交付提交及干净工作树为基线。核心产物是 `AA-Program/Excel_Iterate_Setup.exe`。必读文件为 [换链入口](<../C-01-ExcelIterate/AC-Original Code/ExcelIterate/Operations/LinkAssistant/LinkAssistant.cs>)、同目录的 `LinkBatch.cs`、[产品配置](../C-01-ExcelIterate/AA-Program/C1-Exceliterate_Configure.py) 和下方本轮交付证据。按缺失目标换链成功行数、重开后公式保留格数和撤销恢复工作簿数核对独立 WPS 实测回执。

#### b. 【目标与验收】

验收 0.9.9.8 的链接助手。先用样本副本复测完整路径、文件名和 `file:///` 三类来源，以及新路径中来自公式的方括号。填写缺失目标和空白行，执行替换并保存重开，对照「外部链接」的处理状态、WPS 编辑链接中的新源及原公式。补齐一份目标再更新值，核对取值。另起一批执行取消，核对整簿恢复。可直接运行源码探针和副本验证。需安装或切换正在使用的加载项时安排现场验收。出现保存、公式或恢复异常时保留工作簿副本及日志，按失败阶段复核。

#### c. 【执行规范】

构建使用 Windows、Python 3.10 以上、.NET Framework 4.8 和 C# 5。改动本产品源码或专属说明后运行 Excel BAT，核对 G-11／G-12。用 `WorkbookToolsProbe` 检查映射、文件占用重试和恢复，实机用独立 WPS 处理临时样本。按扫描来源数、填写映射行数、匹配现存目标行数、成功写入缺失目标行数、跳过行数、失败行数、恢复工作簿数记录，附包摘要和运行标识。常规工程边界默认静默执行，遇阻或需决策时再说明。

#### d. 【起手动作】

在仓库根运行 `git status --short --branch` 与 `git log -2 --oneline`，核对分支、交付提交、配置器版本及安装包摘要。记录 WPS 与实际安装的加载项版本，取得副本和独立测试目录后开始扫描。用可复现的源标识与 C 列文本记录首条失败，确认同批其他行的处理状态，再定位对应代码。

### (2). 本轮交付证据

11 月 25 日运行 [网站联合 BAT](../C-03-Website·Xycabin/AA-Program/一键重建并打包.bat)，选择「2. 完整验证」。北京时间 01:54:32 至 01:59:09，退出码 0。18 个重建关口及 G-23 通过。重建标识 `f2e51caeaacc47e5b34aa7425c3874bd`，收尾标识 `3fda4f22f75040399379c10be3944f0d`。见 [BAT 日志](../C-03-Website·Xycabin/build/v0998-bat-20261125-015432.log)、[本轮交付数据](../C-03-Website·Xycabin/build/v0998-delivery-f2e51caeaacc.json)。

`Excel_Iterate_Setup.exe` 为 5775872 字节，SHA-256 为 `8de765f4fef45604a08cd3be89e74ac83a924f33f1588ef4d6fb0e4da256acdc`。Excel 后台实际启动 32 次验证调用，WorkbookToolsProbe 通过 402 项断言，覆盖共享占用、并发编辑及整簿恢复。核验 5 个安装有效载荷。

独立 WPS 使用 0.9.9.8 的 DLL 测试合成工作簿。主场景 12 条缺失目标和 1 条现存目标完成替换，1 条空白映射跳过。保存重开保留 14 格外链公式，补齐目标后刷新取值 1 次成功。另通过 2 个关闭工作簿的待核对来源场景，共按原字节恢复 3 份工作簿。见 [WPS 实测日志](../C-03-Website·Xycabin/build/v0998-link-wps-f2e51caeaacc.log)。

# 三. WordUpdater 项目档案

## 1. 定位与输出

WPS文字VSTO加载项：逐格链接Excel、更新/查看链接及选择正文顶层表格。曾用名WpsLinkUpdater，仅迁移自身旧身份。命名空间/程序集 `WordUpdater`，唯一安装器[AA-Program/WordUpdater\_Setup.exe](../C-02-WordUpdater/AA-Program/WordUpdater_Setup.exe)，内嵌五载荷；用户需完整VSTO运行库及官方授权。Agent-Evo同一DLL提供网站、离线文档、四项自定配置和需求记录，实验OCR占位灰色。

AA-Program放配置器、批处理与正式程序，AB-Graphic放界面截图、功能区图标和示意SVG；文档统一在仓库根C-04-Devlog维护；build是工具、中间产物和验证证据。C-01/C-02直接编译引用C-00-Shared公共实现，C-03是网站唯一源码及联合交付入口，不建立网站副本。

## 2. 版本摘要

版本来源与归并规则见[门禁](<B3-Spec Coding.md#1-文档编号与经验更新>)，当前为 0.9.9.8，沿用项目统一共享发布者（CN=Xymircai WPS Addins Dev，两产品同指纹）。公开配置见 C-00-Shared/signing/signing.json。证书缺失时由构建链从匹配的本机 PFX 导入当前用户个人证书库。

| 版本      | 重要变更                                                                                                                                      |
| ------- | ----------------------------------------------------------------------------------------------------------------------------------------- |
| 0.9.5.0 | 1. 首次提供WPS文字加载项，支持链接Excel数据；2. 采用VSTO构建，提供单文件安装包。                                                                                         |
| 0.9.6.0 | 1. 完善Excel链接、数据更新与查看功能；2. 统一窗口和图标；3. 增加运行库检查与固定证书签名，提升安装可靠性。                                                                              |
| 0.9.7.0 | 1. 按操作流程整理链接菜单，区分操作与反馈入口；2. 新增Agent-Evo选项卡；3. 支持离线阅读开发、使用文档及访问网站。                                                                         |
| 0.9.8.0 | 1. 完善链接建立、更新和查看；2. Agent-Evo提供离线开发/用户预览、全文应用选择和树形尺寸设置；3. XyCabin统一指南、离线下载、音乐和独立音量；4. 规范交付、文档与清理流程。                                        |
| 0.9.9.0 | 1. Agent-Evo同步排序、改名及控件缩放；2. 名称排序即时、格式尺寸下次开窗生效；3. 独立STA预览、窗口快照和需求记录；4. 通用确认隔离并完善Owner校验；5. 开发全文口令验证及离线配图缓存；6. 网格表头完整单行、居中且默认加粗。            |
| 0.9.9.1 | 1. 随联合交付重新打包，同步版本号与内嵌文档；2. WordUpdater 自身功能无变更。                                                                                           |
| 0.9.9.2 | 1. 查看链接窗体去除工具窗样式，修复 WPS 文字宿主单元格无法键入与功能区子菜单失效；2. 文档收敛为 C-04-Devlog 五份，内嵌文档与 Agent-Evo 概览同步新路径与新规则；3. 目录重构为 AA-Program/AB-Graphic，配置器更名 C2。 |
| 0.9.9.3 | 1. 开发/用户全文分别整份内嵌B2与A-UserGuide，导出文件名同源；2. 链接Excel仅校验所选规则矩形，允许选区外合并、拆分或嵌套；3. 保留实际目标Cell并在写入前复核，修复合并表头误拒与物理列号错位；4. 根据来源路径定位已打开的工作簿，兼容隐藏表格实例并恢复Ctrl多区域明细。 |
| 0.9.9.8 | 1. 更新内嵌开发经验及使用说明；2. 同步离线站点和安装包版本。 |

## 3. 当前命令按钮概览

本表以配置器 BUTTONS 与 EXTRA_TABS 为准，保留顶层功能概览。逐项源码路径、回调与共用依赖统一见 [B4-Code Index](<B4-Code Index.md>)；Agent-Evo 默认隐藏，显示名称和排序可由用户自定。

| 命令按钮 | 选项卡与分组 | 子选项 | 功能摘要 |
| --- | --- | --- | --- |
| 更新链接 | WordUpdater·链接 | — | 一键批量更新所选范围内的全部链接 |
| 链接Excel | WordUpdater·链接 | — | 按所选规则矩形逐格链接，允许选区外合并；按来源路径定位工作簿并读取Ctrl多选 |
| 查看链接 | WordUpdater·链接 | — | 查看本文档引用了哪些Excel文件及其工作表/区域 |
| 全选表格 | WordUpdater·操作 | — | 选中本文档中的全部表格（正文不受影响） |
| 文件操作 | WordUpdater·操作 | 文件转换与 PDF 工具（当前文档转 PDF、批量文件转 PDF、PDF 添加水印、PDF 拆分、PDF 合并） | 当前文档或批量转 PDF，以及 PDF 水印、拆分和合并。 |
| 问题反馈 | WordUpdater·反馈 | — | 有BUG或新增功能请发送邮箱（打开 QQ 邮箱网页版） |
| Agent扩展 | WordUpdater·扩展 | — | 显示或隐藏 Agent-Evo 选项卡（开关） |
| 调用API | Agent-Evo·详情 | 智能云-API（百度智能云、阿里智能云）、Agent-API（国内Agent（WorkBuddy、TRAE、Qoder、DeepSeek、豆包、Kimi）、国外Agent（ChatGPT、Claude、Gemini）） | 四级嵌套菜单，打开对应云控制台或 AI 网站（无 screentip，url 按钮） |
| 开发流程 | Agent-Evo·详情 | 开发预览、用户预览、XyCabin | 打开内嵌开发/用户文档与网站首页 |
| 自定配置 | Agent-Evo·操作 | 自定排序、自定名称、自定格式、交互弹窗 | Agent 弹窗的自定义项 |
| 实验计划 | Agent-Evo·操作 | OCR识别、需求记录 | 实验性功能入口 |

## 4. 源码地图与运行入口

| 子系统 | 位置与入口 |
| --- | --- |
| 功能实现 | AC-Original Code；完整定位见 [B4-Code Index](<B4-Code Index.md>)，机器映射在 project.json 的 features |
| 加载与生成 | src/addin 的 ThisAddIn、Connect、AssemblyInfo 及生成文件；AA-Program 配置器和 src/configuration |
| 跨产品共用 | C-00-Shared；产品只保留功能薄入口，经 @shared/ 直接编译引用共用实现 |
| 构建与验证 | 根目录构建脚本；src/installer、src/pdf、src/tools、src/testharness；父目录 verify_shared.py |
| 依赖与证据 | lib 为必要编译依赖；build 为中间物、候选包、缓存与验证证据 |

单项目交付运行本产品 AA-Program 下的「一键重建并打包.bat」。源码诊断可运行 python rebuild.py --background --mode full，再运行 python validate.py --background --site-html build/site-embedded.snapshot.html。候选包使用 --installer-output 写入 build，并通过 validate.py 的 --installer 指定。

链接来源模型与语法统一由 Links/_Shared/LinkFieldCode 解析，WordComCollection 提供 COM 集合适配；查看与粘贴不再调用更新引擎的内部辅助方法。跨产品纯坐标实现位于 C-00-Shared/src/addin/CellAddress，构建工具统一由 tool_catalog 与 project.json 登记。

## 5. 当前状态与交付

### (1). 接手摘要

#### a. 【现场基线】

在仓库根的 `C-02-WordUpdater` 工作。当前分支为 `main`，本轮基准为 `7ffb77c95fbe`，以交付提交及干净工作树为基线。核心产物是 `AA-Program/WordUpdater_Setup.exe`。必读文件为 [产品配置](../C-02-WordUpdater/AA-Program/C2-Wordupdater_Configure.py)、[链接粘贴入口](<../C-02-WordUpdater/AC-Original Code/WordUpdater/Links/_Shared/LinkPasteEngine.cs>) 和下方本轮交付证据。按 Word 后台实际执行用例数、安装包有效载荷数和包摘要核对回执。链接操作以功能区按钮及保存重开后的字段表现为验收对象。

#### b. 【目标与验收】

验收 0.9.9.8 的安装、功能区和表格逐格链接。在独立 Windows 用户或测试机上核对安装包摘要及 VSTO 清单签名，安装后检查加载项版本和 WordUpdater 选项卡。准备含顶层表格的文档及有已知数据的工作簿，用矩形和 Ctrl 多区域各粘贴一次，检查 LINK 代码、单元格显示和更新值。再运行查看与定位，保存重开后复核字段。可直接运行后台回归和样本文档验证。需要改动用户当前安装时进入现场安排。关键项失败时记录错误及最小选区，待安装、建链、刷新全部通过后登记验收。

#### c. 【执行规范】

使用 .NET Framework 4.8、C# 5 和 Python 3.10 以上，正式产物经 Word BAT 更新，核对 G-14／G-15。后台验证和 Word LINK 探针使用隔离样本。现场记录 Windows 与 WPS 版本、安装器 SHA-256、目标表选择范围、粘贴生成字段数、更新成功字段数、定位结果以及文档保存重开结果。按后台回执和现场记录归档各自的运行标识。常规工程边界默认静默执行，遇阻或需决策时再说明。

#### d. 【起手动作】

在仓库根核对 `git status --short --branch`、交付提交及配置版本，运行 `py -B -X utf8 C-02-WordUpdater/verify_package.py` 核对 VSTO 清单发布者与签名，用 `Get-FileHash` 对照安装包与本轮回执的摘要。确定可安装的测试机、空白文档和来源工作簿后开始安装验收。安装通过后按矩形、Ctrl 多选、更新及定位的顺序记录首个失败步骤。

### (2). 本轮交付证据

11 月 25 日运行 [网站联合 BAT](../C-03-Website·Xycabin/AA-Program/一键重建并打包.bat)，选择「2. 完整验证」。北京时间 01:54:32 至 01:59:09，退出码 0。18 个重建关口及 G-23 通过。重建标识 `f2e51caeaacc47e5b34aa7425c3874bd`，收尾标识 `3fda4f22f75040399379c10be3944f0d`。见 [BAT 日志](../C-03-Website·Xycabin/build/v0998-bat-20261125-015432.log)、[本轮交付数据](../C-03-Website·Xycabin/build/v0998-delivery-f2e51caeaacc.json)。

`WordUpdater_Setup.exe` 为 4249088 字节，SHA-256 为 `d8eae2859cdd36e900610d22c665647f6b662119fefad56f62957b37d5e9bcd6`。Word 后台实际启动 25 次验证调用，核验 5 个安装有效载荷。真实安装、功能区逐格链接和刷新由现场验收记录。

# 四. XyCabin 项目档案

## 1. 定位、关系与源码地图

网站介绍两个加载项，提供指南/版本/下载、开发流程、反馈与音乐；邮件/网盘需要网络。HTML内嵌CSS、JS、SVG、MP3和下载指引PNG，无CDN、外部字体、fetch邻近文件依赖。AA-Program/XyCabin.html独立网页含两EXE；两DLL资源是同一无EXE快照。Website不是第三个WPS加载项，不产第三份EXE。

| 当前源码                                         | 职责                            |
| -------------------------------------------- | ----------------------------- |
| AA-Program/C3-XycabinSite_Configure.py       | 唯一品牌、文案、指南、版本、颜色、声音和下载配置      |
| src/site_core.py、site_template.html          | 校验/转义/序列化、单页布局及通用交互           |
| src/site_audio.js、site_downloads.js          | 音频状态机、下载解码/菜单/弹窗              |
| src/assets/tranquility.mp3、unblock-guide.png | 必需项目内音乐片段与下载提示图，普通重建不依赖开发机原文件 |
| rebuild.py、src/delivery.py、build_cache.py    | 固定快照、两包暂存、联合验证/事务发布/回滚、内容缓存   |
| validate.py、src/tools                        | 前端、目录/文档、资源、事务、无旧build集成与清理探针 |

三个输出区统一AA-Program（唯一HTML、配置与批处理）、AB-Graphic（当前PNG与图标）；文档统一在C-04-Devlog维护；build留工具、收据与证据。两加载项不保留网站模板副本；相关输入路径按脚本位置计算。

## 2. 携带源码复现

完整交付带三个同级项目、根目录 verify_shared.py、两加载项 lib、公开 signing.json 和项目内资产。正式签名使用本机已配置的固定证书及私钥，证书缺失时从本机 PFX 导入。仅重建单加载项也带完整 Website。环境：Windows、可执行Python≥3.10、Node.js、.NET Framework4.8/C#5。先实际--version核解释器，子进程参数列表/UTF-8；普通构建不需网络或FFmpeg。

预览运行 `python AA-Program/C3-XycabinSite_Configure.py`，只改HTML；仅使用成功交付记录中版本/大小/摘要匹配的包，无记录或变更时默认下载不可用。不要手改生成HTML或把配置器移出固定目录。

联合交付直接运行本项目AA-Program批处理：1增量、2完整（均--background），或 `python rebuild.py --background --mode full`，不必先跑两单项目批处理。随后 `python validate.py --background` 只核对正式产物，默认共 19 个 G，两个产品业务回归各执行一次；机制自检另用 `python validate.py --background --developer-checks`，独立摘要保存到 development-run-summary.json。后台子进程统一无控制台、关闭默认输入，焦点切换只读记录于本轮日志；回填审计后 `python src/tools/layout_probe.py --strict-evidence`核引用。安装、真实file下载/播放、WPS点击单列验收。

## 3. 当前状态与交付

### (1). 接手摘要

#### a. 【现场基线】

在仓库根的 `C-03-Website·Xycabin` 工作。当前分支为 `main`，本轮基准为 `7ffb77c95fbe`，以交付提交及干净工作树为基线。核心产物为 `AA-Program/XyCabin.html`，页面携带两款 0.9.9.8 安装包。必读文件为 [站点配置](../C-03-Website·Xycabin/AA-Program/C3-XycabinSite_Configure.py)、[下载实现](../C-03-Website·Xycabin/src/site_downloads.js) 和下方本轮交付证据。按完整重建通过关口数、联合收尾通过关口数及内嵌 EXE 字节一致包数核对回执。

#### b. 【目标与验收】

验收离线打开、版本和本地下载。在独立目录用浏览器以 `file://` 打开本轮 HTML，核对两个产品的版本、指南与下载入口。各下载一个安装包，逐个比对文件名、字节数和 SHA-256，目标为两份下载内容均与本地独立包一致。可直接复核本地产物、运行前端门禁并在测试目录下载。遇到浏览器阻止下载时保存提示、版本和页面入口，待两包均能保存且摘要匹配后登记通过。GitHub Release 的处理另按发布任务安排。

#### c. 【执行规范】

Windows 上使用 Python 3.10 以上和 Node.js。收到用户明确的联合打包要求后，用 `AA-Program/一键重建并打包.bat` 选项 2，冻结正文和签名配置，按退出码、18 个构建关口和 G-23 核对结果。回填三处交付证据后运行 `src/tools/layout_probe.py --strict-evidence`，复核剔除审计节后的正文摘要。记录浏览器版本、本地 HTML 摘要、各下载文件大小与 SHA-256。常规工程边界默认静默执行，遇阻或需决策时再说明。

#### d. 【起手动作】

先核对 `git status --short --branch`、交付提交、站点配置版本和本轮收尾标识。用 `Get-FileHash` 核对 HTML 与两包。摘要相符后复制 HTML 到独立目录，打开离线页并开始两次下载，逐条填入结果。发生摘要偏差时保留原产物和回执，检查网页与独立包的交付源是否对应。

### (2). 本轮交付证据

11 月 25 日运行 [网站联合 BAT](../C-03-Website·Xycabin/AA-Program/一键重建并打包.bat)，选择「2. 完整验证」。北京时间 01:54:32 至 01:59:09，退出码 0。18 个重建关口及 G-23 通过。重建标识 `f2e51caeaacc47e5b34aa7425c3874bd`，收尾标识 `3fda4f22f75040399379c10be3944f0d`。见 [BAT 日志](../C-03-Website·Xycabin/build/v0998-bat-20261125-015432.log)、[本轮交付数据](../C-03-Website·Xycabin/build/v0998-delivery-f2e51caeaacc.json)。

`XyCabin.html` 为 14591511 字节，SHA-256 为 `29cc83d9e58436f1f44f860428e9c53e562a14a85838bc68c81684222edc02b0`。内嵌安装包与独立安装包逐字节一致 2 份。两个安装包均为 0.9.9.8。实际浏览器离线下载由现场验收记录。
