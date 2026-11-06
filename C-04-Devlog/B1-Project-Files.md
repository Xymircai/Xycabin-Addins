**三项目档案与版本**

最后更新：2026-11-05 ｜ 当前版本：**0.9.9.3** ｜ 交付物：`Excel_Iterate_Setup.exe`、`WordUpdater_Setup.exe`、`XyCabin.html`

# 一. Agent及开发者接手概览

本仓库维护三个输出：WPS 文字加载项 WordUpdater、WPS 表格加载项 Excel-Iterate，以及两者共用的单文件离线网站 XyCabin。文档分为接手入口、使用说明、项目档案、开发经验、门禁、代码索引六类。加载项主选项卡的“扩展→Agent扩展”显示 Agent-Evo，提供概览/全文、网站和自定配置。版本 0.9.9.3，构建不安装或操作用户 WPS。

## 1. 目录与阅读顺序

| 文档                                         | 内容                               |
| ------------------------------------------ | -------------------------------- |
| [B0-Entry.md](B0-Entry.md)                 | 入口与阅读顺序、项目关系、交付边界                |
| [A-UserGuide.md](A-UserGuide.md)           | 三产品使用说明                          |
| B1-Project-Files.md                        | 本章与三个项目档案：定位、版本摘要、命令按钮、源码地图、当前状态 |
| [B2-DevExperience.md](B2-DevExperience.md) | 共享开发经验与三个专属经验                    |
| [B3-Spec Coding.md（B3）](<B3-Spec Coding.md>) | 编码、构建、验证、交付及网站门禁                        |
| [B4-Code Index.md](<B4-Code Index.md>) | 两个加载项全部末级功能的目录树、入口、核心实现与共用依赖 |

## 2. 项目与交付物

| 项目            | 程序集           | 交付物                     |
| ------------- | ------------- | ----------------------- |
| Excel-Iterate | Excel_Iterate | Excel_Iterate_Setup.exe |
| WordUpdater   | WordUpdater   | WordUpdater_Setup.exe   |
| XyCabin       | —             | XyCabin.html            |

## 3. 源码与资源目录

三个项目统一为 `AA-Program`（配置器、启动器、正式产物）与 `AB-Graphic`（界面截图、功能区图标、示意 SVG）；功能源码在产品 `AC-Original Code/`，按选项卡、分组和菜单叶子归档；加载框架、生成源码及工具仍在 `src/`。公共实现只在同级 `C-00-Shared`，依赖在 `lib/`，构建中间物与验证证据在各产品 `build/`。完整功能导航只在 [B4-Code Index](<B4-Code Index.md>) 维护。文档只在仓库根 `C-04-Devlog` 维护，项目内不再保留副本。

# 二. Excel-Iterate 项目档案

## 1. 定位与输出

WPS表格VSTO加载项：引用/从属追踪、公式、筛选、链接、目录、文件整理/命名、文件名称/OCR、表头匹配合并及合并抵消。命名空间/程序集 `Excel_Iterate`，唯一安装器[AA-Program/Excel\_Iterate\_Setup.exe](../C-01-ExcelIterate/AA-Program/Excel_Iterate_Setup.exe)，内嵌五载荷；用户需完整VSTO运行库及官方授权。Agent-Evo同一DLL提供网站、离线文档、四项自定配置和需求记录，实验OCR占位灰色。

AA-Program放配置器、批处理与正式程序，AB-Graphic放界面截图、功能区图标和示意SVG；文档统一在仓库根C-04-Devlog维护；build是工具、中间产物和验证证据。C-01/C-02直接编译引用C-00-Shared公共实现，C-03是网站唯一源码及联合交付入口，不建立网站副本。

## 2. 版本摘要

版本来源与归并规则见[门禁](<B3-Spec Coding.md#1-文档编号与经验更新>)，本次标记为0.9.9.3并沿用原固定发布者。

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


## 3. 当前命令按钮概览

本表以配置器 BUTTONS 与 EXTRA_TABS 为准，保留顶层功能概览。逐项源码路径、回调与共用依赖统一见 [B4-Code Index](<B4-Code Index.md>)；Agent-Evo 默认隐藏，显示名称和排序可由用户自定。

| 命令按钮 | 选项卡与分组 | 子选项 | 功能摘要 |
| --- | --- | --- | --- |
| 追踪引用 | Excel-Iterate·数据 | — | 解析当前单元格公式，逐项展开被引用单元格的详情；错误行红字，选中后可「显示错误来源」回溯定位源头格并跳转 |
| 追踪从属 | Excel-Iterate·数据 | — | 遍历所有已打开工作簿，找出哪些单元格的公式引用了当前单元格 |
| 筛选增强 | Excel-Iterate·数据 | 按列筛选、快捷高筛 | 按活动单元格所在表格的表头选择显示列 |
| 审计工具箱 | Excel-Iterate·数据 | 批量迁移数据、表头匹配、合并抵消（往来抵消、收入成本抵消） | 按表头列名批量匹配多个工作簿，把列数据合并到目标工作簿；往来抵消按列映射生成抵消分录，收入成本抵消按三维聚合生成矩阵式底稿（借贷恒等式保证平衡），两者均写入前备份、可一键还原 |
| 一键转换 | Excel-Iterate·操作 | 文本提取与过滤 | 组合空格、英文、中文、数字和其他，按显示结果提取或过滤并原位写回文本。 |
| 一键公式 | Excel-Iterate·操作 | 进行公式封装（IFERROR封装（默认返回空、自定义封装）、ROUND封装（默认保留2位、自定义封装）、ABS封装、自定义函数）、移除公式封装、数值计算（加（默认加、自定义计算）、减（默认减、自定义计算）、乘（默认乘、自定义计算）、除（默认除、自定义计算））、相对-绝对转换、公式-数值转换（转公式格式、转文本格式、仅保留数值） | 封装、计算、相对-绝对转换与公式-数值转换（转公式格式、转文本格式、仅保留数值）。 |
| 文件操作 | Excel-Iterate·操作 | 批量处理（文件批量整理、文件批量命名、文件提取名称）、文件转换与 PDF 工具（当前文档转 PDF、批量文件转 PDF、PDF 添加水印、PDF 拆分、PDF 合并） | 文件批量整理、命名、名称提取；当前文档或批量转 PDF，以及 PDF 水印、拆分和合并。 |
| 链接助手 | Excel-Iterate·操作 | — | 扫描并说明名称外链，C列填新路径或“断开链接”；名称清理先预检，整簿备份可恢复 |
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

批量迁移数据位于审计工具箱，支持选择源区域、批量添加文件或文件夹、勾选目标、逐项预检及输出校验后的独立副本。入口及专属实现位于 `C-01-ExcelIterate/AC-Original Code/ExcelIterate/Data/AuditToolbox/DataMigration/`，验证入口为 DataMigrationProbe 与 DataMigrationLiveProbe；详细规则见 A 与 B2。


单项目：运行 python rebuild.py --background --mode full，再运行 python validate.py --background。源码结构迁移验证使用 --installer-output 将候选包写入 build，并通过 validate.py 的 --installer 指定候选包。联合交付仍使用 Website 入口。

## 5. 当前状态与交付

### (1). 当前实现与待验收

新增“操作 → 一键转换 → 文本提取与过滤”，以及“一键公式 → 相对-绝对转换／公式-数值转换”。两个参数窗支持手输地址及隐藏窗口后原生框选，按实际首格显示修改前后示例，默认高度随内容收紧。五类字符可提取或过滤，四种引用锁定可统一转换；“转公式格式”展开“转为公式文本”和“还原文本公式”，后者将等号开头的文本恢复为公式并局部计算。六个入口沿用全批快照、预检和内容／格式回滚，取消与失效目标零写入。真实宿主框选、焦点、菜单和 DPI 外观另列验收。

工程规范由 B3-Gate Control 更名为 B3-Spec Coding.md，仍承担 B3 逻辑角色，六文档边界、R/H/G 编号及校验强度保留。入口、链接、公共文档清单、资源范围及测试夹具同步维护。

新增“数据 → 审计工具箱 → 批量生成过程表 → 单体模板生成”：读取 xlsx/xls 磁盘模板快照，按用户选择的来源字段及显式数据范围逐公司完整复制，可选更新 Lead 指定行的两个字段。命名、来源与目标字段联动加载，异常预检阻断；逐文件取消、保存重开核验和不覆盖发布独立实现。原单体汇总表铺列逻辑不变。实现入口与文件见 B4；使用步骤见 A-UserGuide 对应菜单说明。实际 Excel/VSTO 保存保真、按钮交互及真实业务模板尚待本地验收。

单体汇总表现改为先明确选择工作簿：支持已打开文件下拉、浏览文件及主动使用当前工作簿，初始无选择；汇总工作表独立选择。预检、执行与还原绑定同一工作簿对象及选择版本，不随活动窗口切换；关闭重开或另存为后要求重选，待还原批次锁定选择。浏览打开的目标文件保持打开且不自动保存，供用户核对；实际宿主交互待验收。

单体汇总表新增可选年度跨度：每家公司展开为开始至结束年度的多列，每列重复公司表头，下方独立行写年度，默认倒序。预检显示逐列公司年度及来源路径，识别或配置年度目录段，默认检查年度文件；只改指定目录年份及模板公司文件名，其他引用保持原样。列上限、备份还原范围、日志和读回校验均按公司×年度计划处理，旧单列模式保留；实际 Excel/WPS 写入和整批还原待验收。

单体模板生成的进度区显示已处理任务数、总数、百分比、成功、失败和剩余项数，同时展示当前文件、累计耗时、平均每项耗时与预计剩余。任务收尾后才计数，取消保留真实比例，首项完成前不估计速度；新批次及预检失败清除旧状态。真实宿主中的绘制效果与大文件耗时待验收。

公司工作簿增加独立文件名预览及末尾去除字符数配置。默认自动清理来源名称末尾的 Excel 后缀，再统一使用模板扩展名，避免 `.xlsx.xlsx`；手动裁剪先作用于来源原文。预览展示原名和最终文件名，处理后重名、空名称及已有文件阻断，不改来源数据或 Lead 字段。命名预览与运行日志分标签页展示，生成前仍完整预检。

批量迁移源文件按钮位于源工作簿文本框右侧、源工作表左侧，标签与复选框按行高垂直居中。主窗批量勾选入口整合到“选择”表头的带框箭头，存在未勾选目标时显示蓝色漏斗；弹窗提供搜索及清空、全选搜索结果、反选、选择计数、空态、确定/取消和尺寸拖动，草稿确认后才回写，不隐藏主表目标。

批量迁移默认改为校验后更新目标原文件并保留原版本备份，也可取消勾选另存副本。新增 MigrationOutputTransaction 与 MigrationWindowScope，目标读取器逐项释放、临时窗口按所有权隐藏，修复预检累积工作簿及读写激活压低迁移窗口的问题；更新模式要求目标先保存并关闭。后续段落中的“仅输出副本”描述属于前版实现或当前另存模式。实际 WPS 标签页、焦点及二进制保存行为另列实机证据。

新增审计工具箱“批量迁移数据”：Connect.DataMigration → DataMigrationForm/Host/Engine，单源矩形区域到多个目标位置，内部引用平移、第三方地址冻结、必要名称按作用域迁移，仅输出原格式副本；预检覆盖与冲突，保存重开核对，目标独立失败。源区域旁新增“在表中选择…”按钮，通过宿主鼠标框选回填地址，取消保留原值。顶部缩短三个输入框，选择源文件、读取当前选区与允许全空白复选框同排，控件按中心线对齐。添加目标在同一原生窗口接受文件或文件夹，识别后逐项勾选，只有勾选目标进入本轮请求。目标清单列宽按当前窗口分配，保留标题与工作表可读下限，长路径通过悬停查看。工作表下拉框贴合单元格，编辑与非编辑态保持宽度、字号和行高一致，单击展开长表名列表。首版边界及十万单元格上限见 A-UserGuide；离线托管夹具不等于真实 WPS 或四种格式验收，本轮不安装或发布。

源选择修复：窗口不再自动绑定活动工作簿；xlsx/xlsm 的表名直接从文件读取，选文件不引起临时工作簿开关。预检区分用户未保存原件与 WPS 重算后的自建只读副本，所有权按会话路径映射判定，重复解析仍有效；字体和换行写入后恢复目标行高列宽。实际按钮与宿主焦点体验单列验收。

批次交互加固：取消源区域的 A1 默认值，全空白区域需明确允许及确认；缺少同名目标表时要求选择实际存在的表。汇总显示源非空数/公式数和目标成功失败数量，结果保存目标表及显示位置。窗口计时器防 COM 消息泵重入，已打开目标的宏指纹读取使用共享读写；日志增加异常阶段和完整堆栈。

兼容核验修复：仅在自建读取/输出副本中把 OOXML ISO 日期转换为相应日期基准的数字序列，防止旧 WPS 只读到年份；核验认可范围端点的等价排序，保留各轴 $ 锁定。预检阻止迁移后直接单格自引用。失败输出副本以可读名称归档到诊断目录，正式结果仍保留目标原名和扩展名。

保真与监听修正：增加源包数字格式和合并从属格常量的恢复与最终包级核验，避免宿主读回归一化掩盖实际文件差异。迁移监听细分事件与线程，排除单纯重算；异线程事件回宿主核对对象，真实编辑、关闭等保护保留。其他扫描/抵消仍使用默认监听语义。

本次将文档整理为五部分，共享方法与门禁分开，删重复流水和外部历史范例依赖；保留宿主独有算法/恢复边界，修正控件与四种拖动术语、窗口过度因果和内嵌网页一致性表述。新增表头匹配合并的接线、公式相对行规则与独立验收边界。

批量迁移已用探针自建样本验证 WPS 的 xlsx、xlsm、xls、et 四种格式迁移及保存重开，源与目标原文件摘要保持不变；实机按钮、键入焦点、真实宏代码和大区域耗时仍单列待验收。本轮不安装或发布；按用户要求单独生成本机开发签名安装包，正式签名配置保持不变。

本轮修复链接助手的名称外链：来源可能仅存在于名称定义，扫描仍应列出并说明数量。LinkNameAnalysis按完整来源路径、外部引用顺序与名称作用域分析依赖；混合来源、仍被公式/对象使用或无法安全分析时在修改前停止。预检通过后沿用整簿备份，精确删除所选来源未使用名称，再调用宿主BreakLink并保存核对；不直接删除ZIP关系，不触碰其他来源。C列仅接受“断开链接”，旧文本提示改填。四列清单、旧批次及恢复边界沿用；实测结果、安装包及待验收项目见本轮交付证据。

名称外链扫描状态按用户要求改为“仅名称管理器定义引用（X 个名称）”，数量仍按实际名称定义显示；同一文案同步到离线/实机探针、使用说明及网站指南。已有清单中的旧文本在重新扫描时更新，安装与真实按钮验收单独记录。

补修短名称依赖误判：旧版将原始报表差异!Y$5误认为局部名称Y的使用，并把筛选范围$A$8:$Y$271误认为全局名称Y的使用。UsesName复用FormulaRefRewriter的A1地址跨度，核有效行列边界后排除地址中的名称字母；裸Y、Sheet!Y及超出合法地址边界的真实名称仍执行作用域与依赖保护。仅修改内部分析，公开命令、批次格式和恢复流程沿用；旧版失败、修复后离线与样本副本实机结果见本轮交付证据。

本批（0.9.9.0）修复追踪弹窗锁死宿主并经真实WPS实机验证：弹窗打开后可正常键入单元格、点开有子选项的功能区菜单；合并抵消由置灰占位转为可用：引擎/宿主/窗体三层拆分，写入前备份、可一键还原。收入成本抵消已改为矩阵式审计底稿（三维聚合+活公式+动态年度列，详见第五章(3)），离线探针144项锁定会计模型。底稿写入与还原待联机验收。

原命令ID、用户设置格式、程序集与安装路径沿用。名称/排序即刻刷新；格式/控件/窗口尺寸下次开窗生效，旧窗口保留快照。共享确认隔离、字体所有权与安全预览方法见第三章，不将后台通过认作真实WPS键入/菜单/焦点问题已解决。本轮安装、宿主点击和云端调用单独验收。

本轮沿用版本与发布身份，公共源码迁至C-00-Shared，项目清单直接引用，生成文件与产品入口保留原位置。抵消窗口在首次可见前应用最大化，来源事件先订阅再回填；局部缓冲覆盖布局、网格和拖柄，尺寸预览保持普通窗口。默认布局按窗口及DPI延迟测量；从属扫描和抵消预检通过原宿主线程分块读取、后台纯数据运算，来源变化时中止。宿主不支持变化监听时回退同步读取并记录兼容路径。EliminationSession统一持有目标、快照和任务，大结果预览按需取值，仍完整测宽、支持复制；取消与失败重试保留尚未还原的目标。实际WPS黑影、键入、菜单、跨屏DPI与写入还原由用户验收，离线检查不能代替。

### (2). 本轮交付证据






单体模板生成与单体汇总批量链接已加入命名预设、应用、导入、导出和删除；仅保存参数，应用后重新预检。以当前上游 main 合并验证：Excel 与 Word 编译通过，单体模板检查 128 项、汇总检查 138 项、上游迁移检查 213 项、弹窗核心 127 项、宿主弹窗检查 17 项通过。正式证书缺失，重签未完成；未生成正式安装包，未安装或操作真实工作簿，实际 Excel/WPS 交互仍待本地验收。



本轮补齐被 Git 忽略的本机签名指纹配置和独立 Excel 测试打包入口，已通过该入口实际执行隔离编译、清单签名、标准完整后台验证 32 项（迁移 228 项）、安装器签名及五载荷字节、版本、卸载视图检查。源项目正式签名配置不变，未创建或导出私钥，未修改信任根、安装或启动 WPS；仅交付本机测试包，真实交互与两产品正式联合发布未验收。见 [本机入口验证结果](../C-01-ExcelIterate/build/local-delivery/local-self-build/result.json)。

本轮按当前方式循环切换、重算目标下拉、目标清理和活动窗口诊断更新生成 Excel_Iterate_Setup_20261106_19_TargetControls.exe，保留宿主弹窗隔离修复。编译、清单签名、标准 validate.py --background --mode full 的 32 个项目全量无缓存验证通过，其中迁移 228 项及隔离服务 17 项；安装器签名和 58 项包级检查通过。672 个输入在构建期间未变，20 个相关文件摘要及编译响应路径已核对。沿用原版本和既有本机测试证书，正式证书配置、门禁和信任根不变，开发证书链未受系统信任。本轮不是两产品正式联合发布，Word 未重签或重测；真实 WPS 焦点、保存交互、逐行重算及写入方式仍待验收，未安装、操作用户文件或提交/推送。见 [本轮打包状态](../C-01-ExcelIterate/build/local-delivery/local-installer-20261106-19/delivery-state.json)。

本轮按当前原生框选模拟点击、选区恢复及计时器生命周期保护更新生成 Excel_Iterate_Setup_20261106_18_RangeFocus.exe，保留宿主弹窗隔离修复。编译、清单签名、标准 validate.py --background --mode full 的 32 个项目全量无缓存验证通过，其中迁移 217 项及隔离服务 17 项；安装器签名和 56 项包级检查通过。672 个输入在构建期间未变，20 个相关文件摘要及编译响应路径已核对。沿用原版本和既有本机测试证书，正式证书配置、门禁和信任根不变，开发证书链未受系统信任。本轮不是两产品正式联合发布，Word 未重签或重测；真实 WPS 焦点、保存交互、逐行重算及写入方式仍待验收，未安装、操作用户文件或提交/推送。见 [本轮打包状态](../C-01-ExcelIterate/build/local-delivery/local-installer-20261106-18/delivery-state.json)。

本轮按当前四种写入方式、独立逐行重算、预检保存策略、框选焦点及计时更新生成 Excel_Iterate_Setup_20261106_17_MigrationModes.exe，保留宿主弹窗隔离修复。编译、清单签名、标准 validate.py --background --mode full 的 32 个项目全量无缓存验证通过，其中迁移 217 项及隔离服务 17 项；安装器签名和 54 项包级检查通过。672 个输入在构建期间未变，20 个相关文件摘要及编译响应路径已核对。沿用原版本和既有本机测试证书，正式证书配置、门禁和信任根不变，开发证书链未受系统信任。本轮不是两产品正式联合发布，Word 未重签或重测；真实 WPS 焦点、保存交互、逐行重算及写入方式仍待验收，未安装、操作用户文件或提交/推送。见 [本轮打包状态](../C-01-ExcelIterate/build/local-delivery/local-installer-20261106-17/delivery-state.json)。

本轮按当前预设详情显示及缩进更新生成 Excel_Iterate_Setup_20261105_16_PresetDetails.exe，保留宿主弹窗隔离修复。编译、清单签名、标准 validate.py --background --mode full 的 32 个项目全量无缓存验证通过，其中迁移 213 项及隔离服务 17 项；安装器签名和 50 项包级检查通过。672 个输入在构建期间未变，20 个相关文件摘要及编译响应路径已核对。沿用原版本和既有本机测试证书，正式证书配置、门禁和信任根不变，开发证书链未受系统信任。本轮不是两产品正式联合发布，Word 未重签或重测；真实 WPS 操作与详情外观仍待验收，未安装、操作用户文件或提交/推送。见 [本轮打包状态](../C-01-ExcelIterate/build/local-delivery/local-installer-20261105-16/delivery-state.json)。

本轮按当前批量迁移入口、引擎、窗口和宿主更新生成 Excel_Iterate_Setup_20261105_15_BatchUpdate.exe，包含已提交的预设隔离修复。新增探针的 var 多变量声明引起 CS0819，已只拆成两句，未删断言或更改业务逻辑。编译、清单签名、标准 validate.py --background --mode full 的 32 个项目全量无缓存验证通过，其中迁移 213 项、宿主隔离 17 项；安装器签名及 50 项包级检查通过。包内核对批量交付、缺表处理与隔离接口。672 个输入在最终构建期间未变，20 个相关文件摘要及编译响应路径已核对。原版本、正式证书配置、门禁和信任根保持不变；既有开发证书链未受系统信任。本轮只出表格本机测试包，不是两产品正式联合发布；Word 未重签或重测，真实 WPS 新批量行为仍待验收。未安装、操作用户文件或提交/推送。见 [本轮打包状态](../C-01-ExcelIterate/build/local-delivery/local-installer-20261105-15/delivery-state.json)。

本轮修复预设业务窗体的七处未隔离 ShowDialog：命名、管理和定位窗口由 HostDialogService 结果型接口在独立 STA 创建、读取纯数据及释放，宿主仅接收 JSON/路径后应用设置；文件夹和打开/保存选择统一经隔离服务。原线程门禁及白名单未修改，两产品门禁和编译通过，隔离服务真实消息循环各 17 项通过。Excel 标准 validate.py --background --mode full 的 32 个项目全量无缓存通过；Word 共享回归 2450 项及核心弹窗 68 项通过。672 个源输入在验证期间未变。测试夹具曾因直接隐藏窗口而被 WinForms 取消，已仅修正夹具为透明不激活窗口并重跑。正式证书配置与信任根不变，仅独立 Excel 副本用现有测试证书签清单；Word 正式链因原证书私钥缺失未执行。真实 WPS 交互仍待验收，未安装、打包、操作用户文件或提交/推送。见 [修复记录](../C-01-ExcelIterate/build/local-delivery/preset-modal-fix-20261105/修复结果.txt)。

本轮在本地 feature/batch-data-migration 合入主仓库 7ab1fbb，保留迁移/预设/逐行重算及上游一键转换/公式数值互转。四处冲突位于窗口登记探针、生成的 Ribbon 和开发经验；合并双方窗口与断言，Excel 目录数为 44，功能区由组合配置重新生成，日志均保留。迁移目录 10 文件与原分支一致，上游操作目录 18 文件与主仓库一致，双方菜单标识均保留。两产品编译、迁移/预设 206 项、转换算法 148 项、转换交互 142 项、公司生成 71 项、单体汇总 99 项、Excel/Word 共享界面 11505/2450 项、核心弹窗 125/68 项和两产品功能区检查均通过。672 个输入在验证期间未变，收尾只读核对 main 仍为 7ab1fbb。已保留合并前备份分支；本轮未执行完整 validate 全链或 WPS 实机验收，未打包、安装或推送。见 [合并检查结果](../C-01-ExcelIterate/build/local-delivery/upstream-integration-20261105/integration-summary.json)。

本轮转换完善已交付：两个参数窗读取实际首格并展示修改前后示例，默认高度按内容收紧；“转公式格式”展开“转为公式文本”和“还原文本公式”。规范最终文件名为 B3-Spec Coding.md，六文档边界和 R/H/G 校验保留。

完整 Excel 后台回归 32 次执行均通过；转换引擎及生产适配器 148 项、隐藏控件交互与渲染 146 项、Excel 弹窗核心 116 项、Word 弹窗核心 68 项、R-30 正反例 55 项、文档自测 25 项及目录／规范改名自测 13 项通过。Word 重新编译、清单重签和 Ribbon 检查通过。构建前后输入指纹一致，重建后的新 MVID 已重新完成后台回归。

正式产物为 [AA-Program/Excel_Iterate_Setup.exe](<../C-01-ExcelIterate/AA-Program/Excel_Iterate_Setup.exe>)，版本 0.9.9.3，SHA-256：`9d98306de268320ac7722a25428aa13b4a9a4a696d1d66ad8b9188bd7a66d63c`。正式路径五载荷、内嵌资产及清单双签名已核对；事务替换完成，本轮临时 EXE 已清除。沿用固定证书，EXE 本身仍未签名，未改动证书或系统信任。

见[本轮交付状态](../C-01-ExcelIterate/build/local-delivery/conversion-refinement/delivery-state.json)、[完整后台回归](../C-01-ExcelIterate/build/local-delivery/conversion-refinement/validation-final.log)、[窗口交互与缩放布局](../C-01-ExcelIterate/build/local-delivery/conversion-refinement/ConversionInteractionProbe-final.log)、[输入一致性](../C-01-ExcelIterate/build/local-delivery/conversion-refinement/input-consistency.json)和[正式包核验](../C-01-ExcelIterate/build/local-delivery/conversion-refinement/payload-formal-details.log)。控件渲染及四档缩放为离线证据；真实宿主框选、键入焦点、菜单展开及跨屏 DPI 仍待人工验收。本轮未安装、保存用户工作簿、关闭 WPS、提交或推送。

此前转换功能交付记录：

本轮新增文本提取与过滤、四种引用锁定转换及三种公式结果转换，工程规范已更名为 B3-Spec Coding.md。完整 Excel 后台回归通过；转换引擎及生产适配器夹具 117 项、隐藏参数窗交互及渲染 75 项、Excel 弹窗核心 116 项、Word 弹窗核心 68 项、分类门禁正反例 55 项通过，Word 已重新编译并重签清单。修正新增菜单与弹窗后的精确数量断言，并使分类自测按生产约定接受 UTF-8 BOM。

沿用原版本和固定证书，VSTO 两清单 SHA-256 双签名、五个包内载荷、内嵌文档和网站快照核验通过。正式安装包为 [AA-Program/Excel_Iterate_Setup.exe](<../C-01-ExcelIterate/AA-Program/Excel_Iterate_Setup.exe>)；本轮临时候选 EXE 已清除；Windows 拒绝用该固定证书签署 PE，提示不适用于代码签名，因此 EXE 本身未签名，未改证书配置或系统信任。见[交付状态](../C-01-ExcelIterate/build/local-delivery/conversion/delivery-state.json)、[完整后台回归](../C-01-ExcelIterate/build/local-delivery/conversion/validation.log)、[包内载荷核对](../C-01-ExcelIterate/build/local-delivery/conversion/payload.log)、[正式包验收](../C-01-ExcelIterate/build/local-delivery/conversion/payload-formal.log)和[签名边界](../C-01-ExcelIterate/build/local-delivery/conversion/installer-signature.json)。

本轮最终产物已统一到 AA-Program 正式入口；常规交付沿用该路径，临时验证包验收后完成替换和清除，验证记录保留。

本轮未安装、关闭 WPS、操作用户工作簿、提交或推送；真实 Excel/WPS 的鼠标框选、窗口焦点、菜单及跨屏 DPI 外观待实机验收。离线生产控件渲染见 build 中 previews，未作为真实宿主截图。







本轮保留当前设置菜单对齐、菜单行和框选切换排版，生成 Excel_Iterate_Setup_20261105_14_SettingsMenu.exe。构建登记检查发现本次窗口文件缺少三处稳定尺寸 ID 及预览隔离，已只补回四处必要修正，未以旧文件覆盖新菜单。编译、206 项离线迁移/预设回归、120 项核心弹窗、10052 项表头、清单与安装器测试签名、47 项包级检查均通过；包内确认 MigrationMenuRow。653 个输入在最终构建期间未变，17 个相关文件摘要与副本一致，编译响应路径已核对。共享登记与第 13 包一致，本次不重复 Word 检查。原版本、正式证书配置和系统信任根保留，现有开发证书链未受系统信任。未执行完整 validate 全链或 WPS 实机验收，未安装、操作用户工作簿/预设库或提交/推送。见 [本轮打包状态](../C-01-ExcelIterate/build/local-delivery/local-installer-20261105-14/delivery-state.json)。

本轮基于预设新版生成 Excel_Iterate_Setup_20261105_13_Presets.exe，含 JSON、预设库、管理及定位等当前代码。补齐命名/管理/定位三窗口的稳定尺寸 ID、预览入口、尺寸树和登记探针，预览不读取真实预设库；核心探针同步新增窗口数量并分别检查三个 ID。Excel/Word 编译、206 项离线迁移/预设回归、Excel 核心弹窗 120 项、表头 10052 项、Word 核心弹窗 67 项及共享回归 2450 项均通过。清单及安装器测试签名、46 项包级核对通过，包内确认预设类和保存/应用接口。653 个输入在最终构建期间未变，17 个相关源码摘要与编译副本一致。沿用原版本和现有本机开发证书，正式签名配置及信任根不变；开发证书链未受系统信任。未运行完整 validate 全链或 WPS 导入导出/定位交互实机验收，未安装、操作用户工作簿或预设库、提交/推送。见 [本轮打包状态](../C-01-ExcelIterate/build/local-delivery/local-installer-20261105-13/delivery-state.json)。

本轮按逐行重算首版的最新入口、窗口、宿主和探针代码生成 Excel_Iterate_Setup_20261105_12_RowRecalc.exe，包含触发配置、文件名匹配与明细及逐行重算代码。无需修改业务源码，主程序和探针一次编译通过；653 个输入在打包期间未变，12 个迁移/共享源文件摘要及编译响应路径已核对。199 项离线迁移回归、111 项核心弹窗、10052 项表头检查、清单及安装器本机测试签名、43 项包级检查均通过。包内直接核对触发请求字段、模型、匹配弹框及 ApplyTrigger/WaitCalculated 接口。沿用原版本及现有本机开发证书，正式签名配置和信任根保持不变；开发证书链未受系统信任。真实 WPS 修改触发单元格、等待重算及逐行写入效果仍待实机验收；未执行完整 validate 全链，未安装、操作用户工作簿或提交/推送。见 [本轮打包状态](../C-01-ExcelIterate/build/local-delivery/local-installer-20261105-12/delivery-state.json)。

本轮按当前磁盘的最新入口、窗口和宿主改动生成 Excel_Iterate_Setup_20261104_11_Latest.exe，包含目标区域选择及迁移进度/耗时提示更新，未改动业务源码。653 个输入在构建期间未变，12 个迁移/共享源文件摘要与编译副本一致且编译响应文件已核对。编译、194 项离线迁移回归、清单及安装器本机测试签名、40 项包级检查均通过。沿用原版本和现有本机开发证书，正式签名配置及系统信任根保持不变；开发证书链未受系统信任。未执行完整 validate 全链或 WPS 实机验收，未安装、操作用户工作簿或提交/推送。见 [本轮打包状态](../C-01-ExcelIterate/build/local-delivery/local-installer-20261104-11/delivery-state.json)。

本轮基于合并后的工作区及新修改的入口、引擎、宿主、探针四个文件，生成 Excel_Iterate_Setup_20261104_10_Latest.exe，包含工作簿占用确认处理及同名外链按实际使用阻断的更新。未改动业务源码，653 个输入在打包期间保持不变，12 个迁移/共享源文件摘要与编译副本一致。编译、清单和安装器本机测试签名、194 项离线迁移回归及 40 项包内检查均通过；包内直接核对新增占用处理和外链接口。沿用原版本和现有本机开发证书，源项目正式证书配置及系统信任根保持不变；开发证书链未受系统信任。未运行完整 validate 全链、新增保存/关闭交互的 WPS 实机验收，未安装、操作用户工作簿或提交/推送。见 [本轮打包状态](../C-01-ExcelIterate/build/local-delivery/local-installer-20261104-10/delivery-state.json)。

本轮合入主仓库至 826460d，八个冲突文件均已处理：五份生成代码保留主仓库新增菜单，假宿主 ClearContents 保留一份，两份日志保留双方记录且日期统一。相对主仓库保留 UiKit、迁移窗口、迁移探针及两份日志共五个文件的修订。独立副本中两产品编译通过；迁移回归 191 项、单体模板生成 71 项、单体汇总 99 项、Excel/Word 共享界面 11342/2450 项及核心弹窗 111/67 项均通过。初轮沙箱临时目录访问失败经正常权限构建排除；年度夹具因测试目录包含日期产生额外目录候选，同一二进制改用无日期夹具路径后通过，未更改业务规则。未执行完整 validate 全链或 WPS 实机验收、未生成新包或安装，合并尚未提交或推送。见 [合并检查结果](../C-01-ExcelIterate/build/local-delivery/merge-upstream-20261104/merge-check.json)。

**最新交付：门禁冗余清理与当前 main 完整重建。** 隔离副本不携带旧 build、生成 EXE/HTML 或字节码缓存；原批处理选择 2 完成联合完整重建及收尾独立验证。清理后在原目录再次运行同一批处理，返回 0；原证书、0.9.9.3 与稳定 ID 沿用。

| 检查 | 本轮实际结果 |
| --- | --- |
| 清理 | 7 个目标、18 个文件，实际删除 20,148,605 字节（19.22 MiB）；涉及 `build/stage-feature-layout`、`build/stage-optimization`、两个 `build/manifest-fixture-*` 及三处 build 内字节码缓存 |
| 当前安装包 | [Excel_Iterate_Setup.exe](../C-01-ExcelIterate/AA-Program/Excel_Iterate_Setup.exe)，5,265,408 字节；SHA-256：25568089d8b4a98853530ad7b2c122809e66ea1bf98e88b0adcefeb6b8d0fffb |
| 验证 | 30 项后台任务全通过，缓存复用 0；R-30=pass、cached=false；原证书清单双签名、五载荷、文档/图形及网页快照通过 |
| 本地审计 | [清理实绩](../C-03-Website·Xycabin/build/cleanup-report.json)、[隔离证明](../C-03-Website·Xycabin/build/cleanup-isolated-proof.json)、[原目录批处理日志](../C-03-Website·Xycabin/build/cleanup-official-build.log)；build 证据不入 Git |
| GitHub 与 Release | 构建源码 [938470e](https://github.com/Xymircai/Xycabin-Addins/tree/938470efe79810f3fcd95e9275d561a3ee5693cc) 已普通推送至 main；[现有 Release](https://github.com/Xymircai/Xycabin-Addins/releases/tag/v0.9.9.3) 三附件已同批替换并回下载，大小与 SHA-256 均匹配本表；[远端核验记录](../C-03-Website·Xycabin/build/cleanup-publication-state.json)；原标签未移动 |
| 边界 | 仅源码布局探针修复无旧 HTML 时的构建依赖；功能源码、配置、必要 lib、证书、文档资源与真实验收资料保留；真实 WPS 点击与安装效果仍待实机验收 |

以下为历史交付记录，旧安装包摘要不再代表本轮产物。

本轮解决上游 PR #11 合并冲突：保留批量生成过程表功能，合入上游六个数据迁移文件更新，合并双方交付记录，生成文件保留功能分支内容。迁移检查替身补齐 ClearContents，以匹配上游新增的清空单元格调用。编译、单体汇总 99 项、公司工作簿 71 项、数据迁移 168 项及共享源码分类检查通过；见[合并编译日志](../C-01-ExcelIterate/build/pr11-merge-build-2.log)与[迁移检查日志](../C-01-ExcelIterate/build/pr11-migration-tests-2.log)。检查使用受管替身，未进行 Excel/WPS 实机验收；固定签名证书缺失，未更新安装包。

本轮将“2025年度”等年度表头单元格显式设为水平、垂直居中，仅修改年度表头的两项对齐属性，其他模板及数据区域格式保留。编译与既有 SingleSummaryProbe 99 项回归通过，见[编译日志](../C-01-ExcelIterate/build/summary-year-alignment-build.log)及[回归结果](../C-01-ExcelIterate/build/summary-year-alignment-tests.json)。实际 Excel/WPS 显示未验收；固定签名证书仍缺失，未生成新安装包。

本轮实现单体汇总表多年度展开：每家公司按年度占多列，每列重复公司表头，下方指定行显示年度，默认倒序；年度路径只按完整目录规则换年，含来源预检、逐列计划、整批还原范围及实际公司行勾选回写。编译通过，SingleSummaryProbe 99 项、CompanyWorkbookProbe 71 项、DisciplineProbe 31 项、DialogSizesProbe core 111 项及 RibbonProbe 通过，弹窗目录和模态隔离检查通过。见[多年度测试结果](../C-01-ExcelIterate/build/summary-annual-tests.json)、[详细日志](../C-01-ExcelIterate/build/summary-annual-tests.log)和[编译日志](../C-01-ExcelIterate/build/summary-annual-build.log)。检查使用模型、托管宿主及自建空文件，不代表真实 Excel/WPS/VSTO 写入及还原已验收；固定证书仍缺失，未生成新安装包。

本轮修复公司工作簿重复扩展名，加入默认 Excel 后缀清理、末尾去除字符数及独立文件名预览；生成前按处理后的名称校验碰撞，保持来源记录和 Lead 字段原值。编译通过，CompanyWorkbookProbe 71 项（含后缀大小写/重复、字符裁剪、处理后重名及既有文件保护）、SingleSummaryProbe 70 项、DisciplineProbe 31 项及 DialogSizesProbe core 111 项通过，弹窗目录与模态隔离检查通过。见[命名测试结果](../C-01-ExcelIterate/build/company-filenames-tests.json)、[详细日志](../C-01-ExcelIterate/build/company-filenames-tests.log)和[编译日志](../C-01-ExcelIterate/build/company-filenames-build.log)。未验证真实 Excel/WPS/VSTO；固定签名证书仍缺失，未生成新安装包。

本轮按完成任务数优化公司工作簿生成进度，增加数量、百分比、成功/失败、剩余、当前文件、平均耗时及预计剩余时间；取消保留真实比例，新任务和预检失败重置。编译通过，CompanyWorkbookProbe 55 项、SingleSummaryProbe 70 项及 DialogSizesProbe core 111 项通过，弹窗目录与模态隔离检查通过，见[进度测试结果](../C-01-ExcelIterate/build/company-progress-tests.json)及[编译日志](../C-01-ExcelIterate/build/company-progress-build.log)。未运行真实 Excel/WPS/VSTO；固定证书仍缺失，签名及新安装包未完成。

本轮将单体汇总表改为显式选择工作簿和汇总工作表，预检、执行及还原使用同一绑定对象。编译通过；SingleSummaryProbe 共 70 项（含原有换链回归与工作簿选择夹具）、CompanyWorkbookProbe 45 项、DisciplineProbe 31 项及 DialogSizesProbe core 111 项通过，源码分类、文档、弹窗目录与模态隔离检查通过。证据见[编译日志](../C-01-ExcelIterate/build/summary-workbook-selection-build.log)与[选择及绑定测试](../C-01-ExcelIterate/build/summary-workbook-selection-probe.log)。未执行真实 Excel/WPS/VSTO 测试；固定签名证书仍缺失，签名未完成，未生成新安装包。

本轮新增“单体模板生成”的源码、菜单及弹窗登记已编译通过；CompanyWorkbookProbe 的 45 项规则/托管宿主检查、SingleSummaryProbe 的 46 项既有回归、DisciplineProbe 的 31 项检查、RibbonProbe 和 DialogSizesProbe core 的 111 项检查均通过，见[本轮针对性测试结果](../C-01-ExcelIterate/build/company-workbooks-final-tests.json)及[详细日志](../C-01-ExcelIterate/build/company-workbooks-final-tests.log)。源码分类、文档、弹窗目录及模态隔离检查通过。扩大回归中 TrackerProbe、ErrorSourceProbe、FormulaProbe core 通过，WorkbookToolsProbe 超时后中止后续扩大回归，不能认作完整后台验证通过。固定签名证书在当前环境缺失，重新签名失败，完整 validate 另被新 DLL 与旧清单摘要不匹配阻止；未生成新安装包，原正式安装包保持原状。未进行真实 Excel/VSTO/WPS 测试、安装或发布。编译、签名失败和完整验证受阻分别见 build/company-workbooks-build.log、build/company-workbooks-sign.log 和 build/company-workbooks-validation.log。

本轮新增审计工具箱「批量生成过程表 → 单体汇总批量链接」：配置器登记 menuBatchProcessSheets 与 btnSingleSummary，源码按功能目录 BatchProcessSheets/SingleSummary 落地；Lead 表按表头文字动态定位「编码&简称」「全称」与「选用」三列，候选行默认全勾、执行时回写选用，模板列铺列固定覆盖并逐列换链，生成日志表与隐藏备份表，全程不自动保存工作簿。离线探针 SingleSummaryProbe 37 项通过（换链边界、路径与 ]BS 不动产、其他公司文件名、本簿引用、字符串常量、含（合并）公司、列字母、表头归一化、N=0／全称重复／列越界阻断、弹窗默认值）；R-30 分类门禁、弹窗登记与 Feature099 网格表头检查沿用通过。真实 WPS 表格自建样本回归 41 项通过：Lead 定位、跨列合并阻断、备份与日志、表头写编码&简称、逐列换链、本簿引用原样带出、区间外不读写、重跑追加且不重复建表、两批备份分别还原。实测发现 WPS 拒写「未加单引号且文件名含括号」的外链公式（0x8FE3000B），已按「只补引号、不改文件名与地址」处理并锁进探针。本机缺少 lib 依赖目录与网站快照，未执行正式重建、签名、安装包与安装，版本号沿用；真实按钮点击与安装后界面验收另行记录。






















本轮针对本机缺少 WordUpdater 固定证书的签名中断，独立生成表格加载项本机测试包 Excel_Iterate_Setup_20261104_9_RegressionFix.exe，未修改两产品正式发布者配置。包已包含当前共享网格释放保护及迁移断言/假宿主修订；638 个输入在构建期间未变，三个修订文件与此前通过两产品共享回归的源码摘要一致。编译、清单签名、安装器签名和 37 项包级检查通过；新构建的离线迁移回归 191 项通过。首轮受限环境 File.Replace 访问拒绝后，同一二进制正常权限重跑通过。开发证书为现有本机证书，链不受系统信任且未修改信任根。未生成 Word 安装包、未执行完整 validate 全链或 WPS 实机验收、未安装或提交/推送。见 [本轮打包状态](../C-01-ExcelIterate/build/local-delivery/local-installer-20261104-9/delivery-state.json)。

本轮检查 PR #12（已合并为 826460d，末提交 63e982a），复现原布局断言失败及表头释放空引用。保留本地已有 UiKit/迁移窗口释放保护与两条需求一致的断言修订，新增 FakeCell.ClearContents 以补齐实际空白写入使用的假宿主接口，未改变 AddBook 的候选池语义或删除业务断言。独立重编两产品后，迁移回归 191 项、Excel 共享界面 11306 项、Word 共享界面 2450 项、Excel/Word 核心弹窗 107/67 项及 Excel 纪律 29 项全部通过。受限环境中的 File.Replace 权限拒绝经同二进制正常权限重跑排除；未作为代码缺陷修复。测试不连接 WPS、不操作用户工作簿；本轮未执行完整 validate 全链、未出新包、未安装或提交/推送。此前第 8 包没有当前释放保护修订，包级检查不能替代功能回归。详见 [本轮检查报告](../C-01-ExcelIterate/build/local-delivery/project-audit-20261104/检查结果.md)。

本轮按与备份一致的最新窗口和测试源码生成 Excel_Iterate_Setup_20261104_8_SharedMappings.exe，确认新版已将 AddSourceRows 调整为 AddSideRows，并增加共享源映射及页面侧栏接口，因此包级检查同步核对当前接口。主程序与探针一次编译通过，无需修改业务代码；清单和安装器本机测试签名完成，37 项包级核对直接确认 AddSideRows、AddMappingRow、SidebarAvailable 及现有框选接口。643 个输入在打包期间未变，编译响应文件引用正确快照；原版本与源码正式证书配置保留，开发证书链不受系统信任且未修改信任根。未运行回归、共享源映射/侧栏交互或 WPS 实机验收，未安装或提交/推送。见 [本轮源码及包内核对记录](../C-01-ExcelIterate/build/local-delivery/local-installer-20261104-8/shared-mappings-proof.json)。

本轮核对四个更新文件与 _claude_新版窗口 备份一致，按含 AddSourceRows 的新版窗口生成 Excel_Iterate_Setup_20261104_7_WorkbookPicker.exe，包含源工作簿逐行加入、真实工作簿框选及本机选择方式设置代码。主程序及探针一次编译通过，无需修改业务代码；清单和安装器本机测试签名完成，35 项包级检查直接确认 AddSourceRows、齿轮控件、框选事件及新版宿主接口。643 个输入在打包期间未变，编译响应文件引用正确快照；原版本和源码固定发布者配置保留，开发证书链不受系统信任且未修改信任根。未运行回归、真实只读框选/取消/设置持久化或 WPS 实机验收，未安装、提交或推送。见 [本轮源码及包内核对记录](../C-01-ExcelIterate/build/local-delivery/local-installer-20261104-7/workbook-picker-proof.json)。

本轮已核对四个更新文件与 _claude_新版窗口 备份一致，按含 OpenRangePicker 的最新窗口生成 Excel_Iterate_Setup_20261104_6_RangePicker.exe。主程序与新版探针一次编译通过，无需修改业务代码；清单及安装器本机测试签名完成，32 项包级核对直接验证逐行框选入口、纯三角边栏、区域条、ShowPage 及框外关闭监测器。旧 Glyphs 字段按本轮产品设计移除，核对已改为新接口，不以旧版本特征阻止新包。642 个输入在打包期间未变，编译响应文件引用正确副本。原版本及源码固定发布者配置保留；开发证书链不受系统信任，未修改信任根。未运行回归、拆行/框选/框外关闭交互或 WPS 实机验收，未安装、提交或推送。见 [本轮源码及包内核对记录](../C-01-ExcelIterate/build/local-delivery/local-installer-20261104-6/range-picker-proof.json)。

本轮已读到用户提供的约 159 KB 新版窗口；项目原件与 _claude_新版窗口 备份摘要一致，包含 MigrationMappingBoard、MigrationTabStrip 和 ShowPage。按该版本生成 Excel_Iterate_Setup_20261104_5_MappingBoard.exe，主程序及新版探针一次编译通过，无需修改业务代码；清单及安装器本机测试签名完成，31 项包级检查直接验证新类型、实际 Mappings 字段类型、ShowPage 和 Glyphs。640 个输入文件在打包期间未变，编译响应文件引用对应快照。已检查打包入口，未发现 Git checkout/restore/stash/reset 命令；不能由此判定此前文件修改者。本轮未提交或回滚 Git、未修改信任根，开发证书链仍不受系统信任；未运行回归或 WPS 实机测试，未安装或推送。见 [新版区域条来源及包内核对记录](../C-01-ExcelIterate/build/local-delivery/local-installer-20261104-5/mapping-board-proof.json)。

本轮核对用户反馈：此前两个包保存的 DataMigrationForm.cs 均为固定 Height=120 且无 Glyphs 的同一版本；当前原件含 Glyphs，源码摘要不同，现有记录不能确定此前修改者。已从当前原件生成 Excel_Iterate_Setup_20261104_3_Glyphs.exe；绑定源码、编译副本及保留备份摘要，构建前后均核对一致，编译响应文件确实引用该副本，并从安装包内 DLL 直接验证 Glyphs 字段、五行文字和框内点击/悬停方法。编译、清单及安装器签名完成，28 项包级核对通过；本轮未改写业务源码，未修改信任根，开发证书链仍不受系统信任。未进行实际界面、WPS 实机或完整回归，未安装或推送。见 [Glyphs 来源及包内核对记录](../C-01-ExcelIterate/build/local-delivery/local-installer-20261104-3/glyphs-proof.json)。

本轮按用户要求再次从最新本地工作区生成 Excel_Iterate_Setup_20261104_2.exe，独立 staging 编译、清单重签及安装器本机测试签名完成，无需修改业务代码。25 项包级载荷、身份及卸载入口核对通过，638 个输入在打包期间未变。原版本及源码固定发布者配置保留，既有开发证书链不受系统信任且未修改信任根；未做功能回归、WPS 实机验收、安装或推送。见 [本轮打包记录](../C-01-ExcelIterate/build/local-delivery/local-installer-20261104-2/delivery-state.json)。

本轮按最新本地工作区重新生成 Excel_Iterate_Setup_20261104.exe，独立 staging 编译、清单重签及安装器本机测试签名完成，无需修改业务代码。25 项包级载荷、身份及卸载入口核对通过，638 个输入在打包期间未变；原版本及源码固定发布者配置保留，沿用既有本机开发证书，证书链不受系统信任且未修改信任根。未运行功能回归、预览交互或 WPS 实机验收，未安装或推送。见 [本轮打包记录](../C-01-ExcelIterate/build/local-delivery/local-installer-20261104/delivery-state.json)。

本轮将用户写回的六文件对照表结构改写生成 Excel_Iterate_Setup_20261103_4.exe，包含工作簿多选、两侧交换、区域预览及只写数值代码。独立 staging 一次编译通过，无需修改业务代码；新版回归探针已编译但未运行。清单与安装器本机测试签名完成，25 项包级载荷、身份及卸载核对通过；仅将包级检查中被移除的 MigrationTargetSelectionPopup 类型更新为实际新版 MigrationBookPopup。638 个输入在打包期间未变；原版本及源码固定发布者配置保留，既有开发证书链不受系统信任且未修改信任根。未做功能回归、交换/区域预览/数值写入或 WPS 实机验收，未安装或推送；内置用户指南仍为旧版界面描述。见 [本轮打包记录](../C-01-ExcelIterate/build/local-delivery/local-installer-20261103-4/delivery-state.json)。

本轮将最新源区域预览框选及目标完整区域独立选择代码生成 Excel_Iterate_Setup_20261103_3.exe。独立 staging 一次编译通过，无需修改业务代码；清单重签、安装器本机测试签名及 25 项包级载荷、身份、卸载入口核对完成，638 个输入在打包期间未变。原版本和源码固定发布者配置保留，沿用既有本机开发证书且未修改信任根，证书链不受系统信任。未做完整回归、鼠标框选/自动同步交互或 WPS 实机验收，未安装或推送。见 [本轮打包记录](../C-01-ExcelIterate/build/local-delivery/local-installer-20261103-3/delivery-state.json)。

本轮将编译修复后的工作表预览、列宽自适应和固定明细面板代码生成测试包 Excel_Iterate_Setup_20261103_2.exe，包含 sheetPreview 字段重名修复。独立 staging 编译、清单重签及安装器本机测试签名完成，25 项包级载荷、身份及卸载入口核对通过；638 个输入在打包期间未变。原版本和源码固定发布者配置保留，沿用既有本机开发证书，证书链不受系统信任且未修改信任根。未做完整回归、XLSX 快照实测、预览交互或 WPS 实机验收，未安装或推送。见 [本轮打包记录](../C-01-ExcelIterate/build/local-delivery/local-installer-20261103-2/delivery-state.json)。

本轮按用户要求仅编译新版工作表预览及网格布局代码。首次编译发现新增 MigrationSheetPreview 字段与既有预检按钮均名为 preview，报 CS0102；已将新增控件及其调用改名 sheetPreview，保留旧按钮名称和调用。修正已写回 DataMigrationForm.cs，独立 staging 再编译成功，源码输入在成功编译期间未变。未运行功能回归、XLSX 解析实测或 WPS 实机验收，未重签、打包、安装或推送。证据见 [本轮编译记录](../C-01-ExcelIterate/build/local-delivery/compile-20261103-preview/compile-result.json)。

本轮按用户写回的第一阶段多源迁移代码生成 Excel_Iterate_Setup_20261103.exe，包含迁移窗口、入口编排和选择弹窗改动。独立 staging 一次编译通过、清单及安装器本机测试签名完成，25 项包级载荷、身份及卸载入口核对通过；638 个源码输入在打包期间未变，本轮无需修改业务代码。原版本及正式证书配置保留，沿用既有本机开发证书且未修改信任根，证书链未受系统信任。未执行完整回归或 WPS 实机验收；多源映射、部分提交、还原链和另存汇总行为不能据编译通过认定验收通过。未安装、关闭 WPS 或推送。证据见 [第一阶段测试包记录](../C-01-ExcelIterate/build/local-delivery/local-installer-20261103/delivery-state.json)。

本轮按最新本地工作区重新生成 Excel_Iterate_Setup_20261101_2.exe，包含更新后的迁移入口；编译、清单与安装器本机测试签名完成，25 项包级载荷、身份及卸载入口核对通过，638 个输入文件在打包期间未变。原版本和源码固定发布者配置保留，沿用既有本机测试证书且未修改信任根，证书链仍不受系统信任。本轮未做完整回归、WPS 实机验收、安装或推送。证据见 [本轮打包记录](../C-01-ExcelIterate/build/local-delivery/local-installer-20261101-2/delivery-state.json)。

本轮按最新本地工作区生成 Excel_Iterate_Setup_20261101.exe，包含迁移入口、窗口、宿主、输出事务及窗口恢复改动。独立 staging 编译、清单重签及安装器签名完成，25 项包级载荷、CLR 身份及卸载入口核对通过；638 个输入文件在打包期间未变。版本及原源码固定发布者配置保持不变；沿用既有本机测试证书，根未受系统信任且未修改信任根。本轮未做完整回归或 WPS 实机验收，未安装或推送。证据见 [本轮打包记录](../C-01-ExcelIterate/build/local-delivery/local-installer-20261101/delivery-state.json)。

本轮按用户更新的“还原操作”按钮及相关本地实现重新打包，包含 DataMigrationForm 和 MigrationOutputTransaction.Restore；编译、清单与安装器本机测试签名完成，25 项包级载荷、身份及卸载入口核对通过。638 个输入在打包期间未变，原版本和源码固定发布者配置保持不变；既有开发证书链不受系统信任，未修改信任根。本轮未做完整回归、WPS 实机或实际还原操作测试，未安装或推送。输出 Excel_Iterate_Setup_20261031_4.exe，见 [本轮打包记录](../C-01-ExcelIterate/build/local-delivery/local-installer-20261031-4/delivery-state.json)。

本轮再次按最新工作区重新打包，修复新下拉菜单缺失 RadioButtonState 引用；采用该类型的完整命名空间别名，避免与 Drawing.ContentAlignment 重名。成功编译后重签清单及安装器，25 项包级载荷、身份与卸载入口核对通过，原版本及源码固定发布者配置保留。沿用既有本机开发证书，证书链不受系统信任且未修改信任根；未执行完整回归、WPS 实机验收、安装或推送。输出 Excel_Iterate_Setup_20261031_3.exe，证据见 [本轮打包记录](../C-01-ExcelIterate/build/local-delivery/local-installer-20261031-3/delivery-state.json)；前两次编译失败日志保留于同目录。

用户再次修改迁移窗口、宿主与窗口恢复代码后，已从当前工作区重新生成独立测试包 Excel_Iterate_Setup_20261031_2.exe；编译、清单与安装器本机测试签名、25 项包级五载荷及卸载入口核对通过，638 个源码输入在打包期间未变。正式版本及原作者固定证书配置保留，既有测试签名根未受系统信任且未修改信任根。本轮未进行完整回归、WPS 实机测试、安装或推送。证据见 [本轮重新打包记录](../C-01-ExcelIterate/build/local-delivery/local-installer-20261031-2/delivery-state.json)。

本轮按用户更新后的本地工作区生成测试安装包，包含 MigrationWindowScope 的所有者窗口及迁移弹窗恢复改动；独立 staging 编译、清单重签、安装器签名完成，包级五载荷、CLR 身份、签名结构与卸载入口 25 项核对通过。使用既有本机开发证书 60E01297DE2278451E865FF88B7124C984A9901C，自签名根未受系统信任且未修改信任根；原源码固定发布者及版本保持不变。638 个输入文件在打包期间未变。本轮未执行完整回归或 WPS 实机验收，未安装、关闭 WPS、提交或推送。证据见 [本轮打包记录](../C-01-ExcelIterate/build/local-delivery/package-20261031/delivery-state.json) 与 [包级核对](../C-01-ExcelIterate/build/local-delivery/package-20261031/installer-artifact.log)。

本轮筛选菜单与首行对齐代码已在独立 staging 编译，以已有本机开发证书 60E01297DE2278451E865FF88B7124C984A9901C 重签清单并生成测试安装包；原项目固定发布者及版本保持不变。包级载荷、CLR 身份、签名结构和卸载入口探针 25 项通过；自签名根未受系统信任，未修改信任根。后台验证在文档资源探针访问自建临时目录时因 WinError 5 中止，不代表完整回归通过；随后用户要求无需检查，因此未重跑、未进行本轮界面或 WPS 验收。未安装、未关闭 WPS、未提交或推送。证据见 [打包记录](../C-01-ExcelIterate/build/local-delivery/filter-alignment/delivery-state.json)、[未完成验证日志](../C-01-ExcelIterate/build/local-delivery/filter-alignment/package-validation.log)。

本轮源文件按钮移位与“选择”表头下拉已完成，主窗不再显示全选/全不选。迁移探针 168 项通过，覆盖同名不同路径搜索、勾选草稿取消、一次失效、无变化保留预检及运行中拒绝回写。交互样本验证 Esc 保留四项，取消后重开，再仅对两个搜索结果反选并确认，主表更新为两项且四行均可见，样本正常退出。初次交互发现 Closed 内同步 Dispose 触发 WinForms 激活筛选器的 ObjectDisposedException，现保留已关闭对象到后续正常事件释放，源码记录此寿命约束。独立测试包完整后台验证通过，Feature099 10873 项（较前版少两个已移除按钮的控件断言）、弹窗核心 107 项、安装包载荷及卸载入口 25 项通过；沿用本机开发证书，根证书仍不受信任，正式签名配置未改。见[交付状态](../C-01-ExcelIterate/build/local-delivery/selection-menu/delivery-state.json)、[交互核对](../C-01-ExcelIterate/build/local-delivery/selection-menu/interactive-check.json)和[完整后台验证](../C-01-ExcelIterate/build/local-delivery/selection-menu/package-validation.log)。本轮未安装、关闭 WPS、提交、推送或发布；独立交互样本不代替用户 WPS 版本安装后验收。

按用户要求生成迁移覆盖与静默修复的独立本机测试安装包，使用既有开发证书 60E01297DE2278451E865FF88B7124C984A9901C，正式源码的发布证书及版本不变。独立副本包含当前九份迁移实现源码，构建期间 637 份规范输入未变；两份清单 SHA-256 双签名、实际 CLR 标识和五载荷摘要通过，完整后台验证通过，迁移 154 项、Feature099 10875 项及卸载流程 54 项通过，安装包实际载荷和维护卸载入口 23 项通过。EXE 的签名者与开发证书一致，证书链根仍不受系统信任，未改信任根。见[安装包交付状态](../C-01-ExcelIterate/build/local-delivery/quietmigration-installer/delivery-state.json)、[完整后台验证](../C-01-ExcelIterate/build/local-delivery/quietmigration-installer/package-validation.log)和[包内载荷核对](../C-01-ExcelIterate/build/local-delivery/quietmigration-installer/installer-artifact.log)。本轮仅生成供用户测试的安装包，未安装、关闭 WPS、提交、推送或发布。

本轮批量迁移覆盖与静默修正已完成本地验证：默认更新目标原文件，临时副本经保存重开及包级校验后以 File.Replace 提交，并保留替换前版本备份；可显式选择原另存模式。迁移 154 项断言通过，包含十二目标独立失败、取消、共享占用、旧摘要、备份冲突及逐项释放。真实 WPS 的 xlsx/xlsm/xls/et 均通过另存与更新原件、备份字节和交付文件重开可见核验；实际点击生产窗体预检与执行，86 个进度观测点失焦为 0，均未最小化，新增临时工作簿最多 2 个。初次无可见宿主的实验出现调用拒绝，现预检提前说明可见宿主要求；单认后端 PID 的恢复失焦实验未通过，补齐根窗口及 Owner 的前台框架 PID 后才通过。两个产品重新编译，Excel 25 项、Word 22 项后台安全探针全部通过。见[修正与验证状态](../C-01-ExcelIterate/build/local-delivery/quiet-replace/delivery-state.json)、[真实窗口焦点及四格式记录](../C-01-ExcelIterate/build/local-delivery/quiet-replace/quiet-replace-ui.log)、[迁移断言](../C-01-ExcelIterate/build/local-delivery/quiet-replace/DataMigrationProbe.log)。本轮未生产签名打包、安装、提交、推送或更新 PR；用户截图所用版本、大区域和真实宏代码另行验收。

本轮本地源码合并：接入上游 `09fc962` 的七个提交，将批量迁移七份源码归入审计工具箱 DataMigration 功能目录；保留上游表头匹配更名、分类门禁与代码索引。迁移七份源码的规范化文本与合并前备份一致，迁移事件接收契约随上游拆分迁至 DirectoryNavigation。两个产品均重新编译，Excel 本轮后台探针全部通过，含迁移 133 项、表头匹配 104 项、Feature099 10875 项、弹窗核心 107 项及卸载流程 54 项；分类门禁正反测试 55 项通过。见[合并与验证摘要](../C-01-ExcelIterate/build/local-delivery/upstream-integration/integration-summary.json)、[逐探针结果](../C-01-ExcelIterate/build/local-delivery/upstream-integration/regressions.json)。本轮仅本地未签名编译及离线回归，未运行正式签名打包、完整包验收或真实 WPS 四格式复验；未安装、提交、推送或发布，版本与正式签名配置沿用。

本轮分类门禁实现提交 [3cff8c5](https://github.com/Xymircai/Xycabin-Addins/commit/3cff8c5856b6d32c2fc20a10994dd2228cb80892)；构建源码与 README 提交 [bfccdec](https://github.com/Xymircai/Xycabin-Addins/tree/bfccdecb9bc3dab0af0030def01d1f9aac6a07ae)，原有四个未推送提交完整保留并一并推送。详见 [本轮验证与发布记录](../C-03-Website·Xycabin/build/github-release-r30/delivery-review.md) 与 [完整证据摘要](../C-03-Website·Xycabin/build/github-release-r30/delivery-summary.json)。

以下为本地迁移开发期间的历史验证记录，独立于上述远端发布记录；不代表本轮合并已重新发布。

工作表下拉框修正留空与编辑态缩小，保留 ComboBox 形态。[交互尺寸核对](../C-01-ExcelIterate/build/local-delivery/sheetfit/interactive-check.json)在自建清单实际点击：编辑器与单元格同宽 124，编辑器高 33 对应单元格高 34（预留一像素行分隔），两者字号均为 10，内边距为零；展开列表宽 217，可完整选择中文长表名并即时提交。[交互记录](../C-01-ExcelIterate/build/local-delivery/sheetfit/migration-sheet-interactive.log)保留过程；此为交互桌面托管清单，不等同 WPS 安装后的宿主验收。[完整后台验证](../C-01-ExcelIterate/build/local-delivery/sheetfit/package-validation.log)通过，含迁移 133 项和弹窗核心 107 项；[载荷及卸载入口](../C-01-ExcelIterate/build/local-delivery/sheetfit/installer-artifact.log)19 项通过。[交付状态](../C-01-ExcelIterate/build/local-delivery/sheetfit/delivery-state.json)记录本机开发签名及摘要，未安装、关闭 WPS 或修改原件及信任根。

目标清单新增按窗口视口分配列宽，勾选列紧凑，标题和工作表保留下限，完整路径使用提示。[布局记录](../C-01-ExcelIterate/build/local-delivery/responsivecolumns/layout-preview.log)核对 40 行清单在 900、1180、1600 再回 900 的窗口宽度，每种宽度重复刷新三次；列宽总和恰好填满扣除竖滚动条后的可用宽度，往返分配一致，无多余横向滚动条。[完整后台验证](../C-01-ExcelIterate/build/local-delivery/responsivecolumns/package-validation.log)通过，含 Feature099 10875 项、迁移 133 项、弹窗核心 107 项；[安装包实际载荷及卸载入口](../C-01-ExcelIterate/build/local-delivery/responsivecolumns/installer-artifact.log)19 项通过。[交付状态](../C-01-ExcelIterate/build/local-delivery/responsivecolumns/delivery-state.json)记录本机开发签名及摘要。本轮为托管窗口布局与后台回归，未安装、关闭 WPS、修改信任根或声称已完成 WPS 安装后的界面验收。

添加目标统一支持文件和文件夹，清单首列勾选控制本轮请求；没有新增主窗文件夹选项。[迁移回归](../C-01-ExcelIterate/build/local-delivery/foldertargets/migration-folder-probe.log)133 项通过，覆盖四格式扫描、当前层边界、去重与过滤、全选/全不选、勾选变更失效、重复添加保留取消状态、未勾选目标无输出且原件摘要不变、坏目标独立失败。[交互桌面记录](../C-01-ExcelIterate/build/local-delivery/foldertargets/migration-folder-picker-interactive.log)确认同一个原生选择器接受自建文件夹，第二次取消不返回路径；未声称已完成安装后的 WPS 功能区全流程验收。[完整后台验证](../C-01-ExcelIterate/build/local-delivery/foldertargets/package-validation.log)通过，Feature099 10875 项、迁移 133 项、弹窗核心 107 项通过；[实际载荷及卸载入口](../C-01-ExcelIterate/build/local-delivery/foldertargets/installer-artifact.log)19 项通过。[交付状态](../C-01-ExcelIterate/build/local-delivery/foldertargets/delivery-state.json)记录本机开发签名及摘要；未自动安装、关闭 WPS、修改信任根或用户原文件。

迁移窗口顶部紧凑布局：[布局记录](../C-01-ExcelIterate/build/local-delivery/compactlayout/layout-preview.log)确认默认宽度 1180 时三个源输入项、在表中选择、选择源文件、读取当前选区及允许全空白复选框同处一行，中心线偏差不超过一像素；带旧版保存宽度的开窗预览通过。仅在当前格式副本处理旧默认宽度，保留其他自定义尺寸及磁盘配置。最终包[完整后台验证](../C-01-ExcelIterate/build/local-delivery/compactlayout/package-validation.log)通过，包括自定义尺寸、121 项迁移检查与弹窗核心 107 项；[载荷和卸载入口](../C-01-ExcelIterate/build/local-delivery/compactlayout/installer-artifact.log)19 项通过，[交付状态](../C-01-ExcelIterate/build/local-delivery/compactlayout/delivery-state.json)记录摘要与本机开发签名。此轮只验证托管窗口布局和后台回归，未安装、关闭 WPS 或声称已完成安装后的宿主界面验收。

源区域新增原生鼠标选择：[托管回归](../C-01-ExcelIterate/build/local-delivery/rangepicker/packaged-range-probe.log)121 项通过，覆盖区域归属、离散选区、格数上限、取消及自建副本清理。[WPS 鼠标实测](../C-01-ExcelIterate/build/local-delivery/rangepicker/native-mouse.log)在自建样本拖选 B1:C3、确认返回正确地址，第二次取消返回空值，原文件摘要不变且读取副本已清理；此项不代替安装后功能区全流程验收。[本机开发签名包状态](../C-01-ExcelIterate/build/local-delivery/rangepicker/delivery-state.json)、[完整后台验证](../C-01-ExcelIterate/build/local-delivery/rangepicker/package-validation.log)与[实际载荷及卸载入口检查](../C-01-ExcelIterate/build/local-delivery/rangepicker/installer-artifact.log)通过。保持正式证书配置、版本与产品身份，未自动安装、改信任根、关闭 WPS 或修改用户工作簿。

用户成功文件的独立审计发现日期格式代码变化和 H21 隐藏常量丢失，现已另存两个迁移修正结果。[修复回归](../C-01-ExcelIterate/build/migration-fidelity-probe.log)107 项通过，[用户副本双目标修复](../C-01-ExcelIterate/build/migration-fidelity-repair.log)通过；[结果一审计](../C-01-ExcelIterate/build/audits/success-result-20261030/repaired-1.json)与[结果二审计](../C-01-ExcelIterate/build/audits/success-result-20261030/repaired-2.json)均为内容差异零、样式差异零，名称、合并、目标尺寸及区域外数据保持。四格式及原规则见[实机回归](../C-01-ExcelIterate/build/migration-fidelity-live.log)。原失败日志唯一证实来源监听保护触发；[原行为复跑](../C-01-ExcelIterate/build/migration-events-before.log)未自然复现旧失败，不能据此认定历史事件必为重算。旧入口丢失事件类别且只存布尔状态，无法回溯；托管测试复现计算误报并验证真实编辑和异线程关闭保护。本机开发证书签名后的保真与监听修复包完整后台验证通过，107 项迁移回归及实际五载荷/卸载入口 19 项通过；[交付状态](../C-01-ExcelIterate/build/local-delivery/fidelityfix/delivery-state.json)、[完整后台验证](../C-01-ExcelIterate/build/local-delivery/fidelityfix/package-validation.log)、[失败分析边界](../C-01-ExcelIterate/build/local-delivery/fidelityfix/failure-analysis.json)。已交付两个独立修正结果与新安装包，未覆盖原件、自动安装、关闭用户 WPS 或推送发布。

日期、范围等价与诊断文件回归：[托管探针](../C-01-ExcelIterate/build/migration-compat-probe.log)99 项通过；[真实 WPS 探针](../C-01-ExcelIterate/build/migration-compat-live.log)通过四格式、两个结果中的日历日期与中文文件名、逆序范围保存重开及 A1 直接自引用预检阻止。用户本轮两个文件实际未通过 C17 核验，长标识文件为未交付副本；历史未占用自建副本已移入 build/user-diagnostics，原件未改。本机开发证书签名后的兼容修复包完整后台验证通过，实际五载荷及卸载入口 19 项通过；[交付状态](../C-01-ExcelIterate/build/local-delivery/compatfix/delivery-state.json)、[后台验证](../C-01-ExcelIterate/build/local-delivery/compatfix/package-validation.log)与[EXE 检查](../C-01-ExcelIterate/build/local-delivery/compatfix/artifact-validation.log)。原发布配置不变，未自动安装、关闭用户 WPS 或推送发布。

用户反馈空白结果后检查原件与读取副本字节一致，实际批次复制空白 A1 且一目标在执行前 COM 失败。主动嵌套 Tick 的[旧实现回归](../C-01-ExcelIterate/build/migration-reentrancy-before.log)失败，修复后[托管回归](../C-01-ExcelIterate/build/migration-batch-probe.log)86 项通过。[两个已打开目标副本回归](../C-01-ExcelIterate/build/migration-batch-repro.log)通过，正式窗口消息泵的[完整区域与四格式回归](../C-01-ExcelIterate/build/migration-batch-live.log)单独记录；真实按钮焦点仍待用户验收。本机开发证书签名后的批次修复包完整后台验证通过，实际五载荷及卸载入口 19 项通过；[交付状态](../C-01-ExcelIterate/build/local-delivery/batchfix/delivery-state.json)、[后台验证](../C-01-ExcelIterate/build/local-delivery/batchfix/package-validation.log)与[EXE 检查](../C-01-ExcelIterate/build/local-delivery/batchfix/artifact-validation.log)。正式窗口消息泵实测源非空 112、公式 11，两个目标成功、各生成一个结果；逐一重开核内容、公式和显示位置。未自动安装、关闭 WPS 或推送发布。

本轮源选择故障回归：DataMigrationProbe 80 项通过，覆盖不打开宿主的 OOXML 表名选择、空初始来源、重复定位后的只读副本脏标记及用户原件未保存仍阻止；[托管回归](../C-01-ExcelIterate/build/migration-picker-probe.log)。[真实 WPS 回归](../C-01-ExcelIterate/build/migration-picker-live.log)完成四种格式及用户测试源副本 B5:I22 到 B100 的 144 格迁移、保存重开，核内容、名称、样式、多行合并、目标尺寸及原件摘要不变。实际按钮焦点及用户安装后体验仍待验收。本机开发证书重签后的安装包完成完整后台验证，实际 EXE 载荷及卸载入口 19 项通过；见[迁移修复包状态](../C-01-ExcelIterate/build/local-delivery/migrationfix/delivery-state.json)、[打包验证](../C-01-ExcelIterate/build/local-delivery/migrationfix/package-validation.log)与[产物检查](../C-01-ExcelIterate/build/local-delivery/migrationfix/artifact-validation.log)。仅交付新文件名的修复包，未替换用户安装或关闭 WPS。

用户尝试上轮本机包安装时，官方 VSTOInstaller 返回 -201：旧清单声明 Interop 2.8，实际文件是强签 office 15，安装拒绝。原后台门禁仅核摘要及签名，未核 CLR 依赖身份，不能据此认定可安装。现新增 PE/AssemblyRef 身份门禁和回归，按实际元数据重建清单；修正版仍为 WPS 加载项，微软授权窗口标题保持其原样。修正版实际清单已绑定 office 15 的真实身份，清单双签名及完整后台验证通过；PE 标识回归 12 项、实际 EXE 内嵌载荷验证 19 项通过，旧 EXE 在增强验证下明确失败。见[修正版状态](../C-01-ExcelIterate/build/local-delivery/fixed/delivery-state.json)、[后台验证](../C-01-ExcelIterate/build/local-delivery/fixed/package-validation.log)与[旧包拒绝回归](../C-01-ExcelIterate/build/local-delivery/fixed/old-package-regression.log)。已交付不同文件名的修正版，未自动安装或改动用户的 WPS 登记，实装结果待用户重试。

按用户要求完成本机开发签名 EXE：正式配置器的原证书指纹未改，隔离源码副本固定专用本机开发证书；清单双签名、五载荷及完整后台验证通过，安装产物 13 项检查与安装/卸载流程 54 项检查通过，共享 Word 安装器源码编译通过。新增 Windows 应用列表卸载入口及保留维护工具；删除下载包后仍可卸载。包摘要、签名身份与路径见[本机交付状态](../C-01-ExcelIterate/build/local-delivery/initial/delivery-state.json)，原始日志见[签名打包](../C-01-ExcelIterate/build/local-delivery/initial/package-build.log)、[后台验证](../C-01-ExcelIterate/build/local-delivery/initial/package-validation.log)与[产物验证](../C-01-ExcelIterate/build/local-delivery/initial/artifact-validation.log)。仅生成交付文件，未实际安装、卸载或推送发布；签名为本机开发证书，不声称为原发布者签名。

本轮批量迁移开发验证：两个加载项均已成功编译；新增 [DataMigrationProbe](../C-01-ExcelIterate/build/data-migration-probe.log) 72 项通过，[DataMigrationLiveProbe](../C-01-ExcelIterate/build/data-migration-live.log) 在 WPS 自建样本中完成 xlsx/xlsm/xls/et 的迁移、保存重开与原文件摘要不变核对。既有表头合并 104 项、公式核心 172 项、抵消 222 项、维护纪律 29 项均通过；弹窗核心 107 项、Feature099 10778 项、内嵌文档核心 113 项均通过。实际按钮焦点、含真实宏代码的样本和大区域耗时仍待验收。本轮未签名打包、安装或发布；以下保留的历史发布记录不代表新增功能已发布。

后续本地完整重建后的包摘要已变化，当前安装包以[远端发布校验记录](../C-03-Website·Xycabin/build/github-release-publication/publication-state.json)为准；原优化交付摘要和性能汇总保留为历史实验，未冒充本次重新计时。

本产品 54 个末级入口通过 [B3 R-30](<B3-Spec Coding.md#6-功能源码分类与代码索引>) 与 [B4 代码索引](<B4-Code Index.md>) 核验。完整与增量模式均只读检查；55 项正反例包含 16 次实际入口违规阻断，不自动修正目录或索引。

| 检查 | 结果 |
| --- | --- |
| 安装包 | [Excel_Iterate_Setup.exe](../C-01-ExcelIterate/AA-Program/Excel_Iterate_Setup.exe)，4,871,680 字节，版本 0.9.9.3；SHA-256：7a61efb06f15a02d5946dea233e22749407d12767d7b82432e1fe47c57cdee9d |
| 重建与后台 | [联合完整重建](../C-03-Website·Xycabin/build/last-run-summary.json)与[收尾独立验证](../C-03-Website·Xycabin/build/validate-run-summary.json)成功；本产品 26 项完整后台任务通过、缓存复用 0；[产品验证日志](../C-03-Website·Xycabin/build/Excel_Iterate-validation.log) |
| 分类与编译 | [验证收据](../C-01-ExcelIterate/build/validation-timings.json)记录 R-30=pass、cached=false；目录、实际回调、编译清单及 B4 一致 |
| 签名与载荷 | [重建及签名日志](../C-03-Website·Xycabin/build/Excel_Iterate-rebuild.log)与[五载荷和资源核对](../C-03-Website·Xycabin/build/Excel_Iterate-payload.log)通过；沿用原发布者、清单双签名，文档与图形资源匹配 |
| 接口兼容 | [公开接口对照](../C-03-Website·Xycabin/build/github-release-r30/public-api-summary.json)：680 条记录保持一致，新增 0、删除 0 |
| GitHub 与 Release | 已普通推送至 main；[现有 Release](https://github.com/Xymircai/Xycabin-Addins/releases/tag/v0.9.9.3)附件已替换并回下载，摘要匹配本地正式产物；HTML 内嵌 EXE 与独立附件逐字节一致 |
| 源码链接 | [本轮精确提交](https://github.com/Xymircai/Xycabin-Addins/tree/bfccdecb9bc3dab0af0030def01d1f9aac6a07ae)及[对应源码 ZIP](https://github.com/Xymircai/Xycabin-Addins/archive/bfccdecb9bc3dab0af0030def01d1f9aac6a07ae.zip)；既有版本标签保持历史指向，自动 Source code 附件仍是旧标签源码 |
| 验收边界 | 实际 WPS 点击、宿主交互、安装效果、焦点与跨屏待实机验收；本轮未安装、连接或关闭 WPS，未重置用户配置 |

# 三. WordUpdater 项目档案

## 1. 定位与输出

WPS文字VSTO加载项：逐格链接Excel、更新/查看链接及选择正文顶层表格。曾用名WpsLinkUpdater，仅迁移自身旧身份。命名空间/程序集 `WordUpdater`，唯一安装器[AA-Program/WordUpdater\_Setup.exe](../C-02-WordUpdater/AA-Program/WordUpdater_Setup.exe)，内嵌五载荷；用户需完整VSTO运行库及官方授权。Agent-Evo同一DLL提供网站、离线文档、四项自定配置和需求记录，实验OCR占位灰色。

AA-Program放配置器、批处理与正式程序，AB-Graphic放界面截图、功能区图标和示意SVG；文档统一在仓库根C-04-Devlog维护；build是工具、中间产物和验证证据。C-01/C-02直接编译引用C-00-Shared公共实现，C-03是网站唯一源码及联合交付入口，不建立网站副本。

## 2. 版本摘要

版本来源与归并规则见[门禁](<B3-Spec Coding.md#1-文档编号与经验更新>)，本次标记为0.9.9.3并沿用原固定发布者。

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

单项目：运行 python rebuild.py --background --mode full，再运行 python validate.py --background。源码结构迁移验证使用 --installer-output 将候选包写入 build，并通过 validate.py 的 --installer 指定候选包。联合交付仍使用 Website 入口。

## 5. 当前状态与交付

### (1). 当前实现与待验收

本次将文档整理为五部分，共享方法与门禁分开，删重复流水和外部历史范例依赖；保留宿主独有算法/恢复边界，修正控件与四种拖动术语、窗口过度因果和内嵌网页一致性表述。已审阅Excel专属合并、OCR与文件批次，只同步界面/维护方法，不移植工作表业务。

本轮修复“链接Excel”被选区外合并单元格阻止的问题：取消整表 Uniform/行列数拦截，按实际选中 Cell 与父表 XML 中对应单元格建立网格位置，选区自身仍须无合并、无嵌套且形成规则矩形；写入前复核同一批目标 Cell。离线失败回归与实机合并表头四格、左侧合并右侧四格已通过，字段引用、显示值、后续更新和选区外内容/格式均已核对。最终构建、签名与安装包证据见本轮交付证据；安装与真实按钮点击单独验收。

后续Ctrl多选反馈经新样本复现：默认Application实际指向Workbooks.Count=0的隐藏ET，来源工作簿在另一实例且只登记文件ROT；旧代码找到隐藏Application后漏查文件登记，窗口回读仍无来源。现无论是否已发现Application，都精确检查剪贴板来源路径对应的ROT登记，核对Workbook.FullName后通过Workbook.Application定位实际实例，再读全局/窗口选区；路径、表名、目标格数、精确外边界及不重叠约束继续生效，歧义拒绝。新CtrlSelectionLiveProbe由正常文件关联打开两个样本的临时副本，保留隐藏实例并存条件，覆盖第1/4行和第1/3行各6格到合并表头下2×3；原件不改动。安装与真实按钮点击仍单独记录。

原命令ID、用户设置格式、程序集与安装路径沿用。名称/排序即刻刷新；格式/控件/窗口尺寸下次开窗生效，旧窗口保留快照。共享确认隔离、字体所有权与安全预览方法见第三章，不将后台通过认作真实WPS键入/菜单/焦点问题已解决。本轮安装、宿主点击和云端调用单独验收。

本轮沿用版本、原证书及稳定命令，公共实现改为C-00-Shared单源引用。更新链接与链接查看统一使用WordHostStateScope：完整取得提示、刷新状态后才修改，设置失败时恢复已触碰状态，恢复失败明确报告；移除诊断前固定等待，不关闭非自有进程。共享窗口尺寸和网格测宽采用统一生命周期与缓存。原逐格链接、Ctrl来源定位、表格结构、数值变化标黄及配置继续回归；本轮不连接实际宿主、不自动安装或关闭WPS。

### (2). 本轮交付证据


本轮因共享弹窗结果接口修复同步重编 Word；原宿主模态扫描、17 项隔离服务真实消息循环、2450 项共享界面和 68 项核心弹窗回归通过。正式证书配置未修改；本机缺少原 Word 签名证书及私钥，因此没有重签、打包或记为 Word 完整验证通过。未连接或安装 WPS，未修改用户文档，未提交/推送。见 [共享修复验证](../C-02-WordUpdater/build/local-delivery/preset-modal-fix-20261105/fix-summary.json)。

**最新交付：门禁冗余清理与当前 main 完整重建。** 隔离副本不携带旧 build、生成 EXE/HTML 或字节码缓存；原批处理选择 2 完成联合完整重建及收尾独立验证。清理后在原目录再次运行同一批处理，返回 0；原证书、0.9.9.3 与稳定 ID 沿用。

| 检查 | 本轮实际结果 |
| --- | --- |
| 清理 | 2 个目标、2 个文件，实际删除 8,494,080 字节（8.10 MiB）；涉及 `build/stage-feature-layout`、`build/stage-optimization` |
| 当前安装包 | [WordUpdater_Setup.exe](../C-02-WordUpdater/AA-Program/WordUpdater_Setup.exe)，4,296,704 字节；SHA-256：85cd8d8077b21cdb7ca37f08472ad00d10d29bd4ea30700f8a4de94acdc3257b |
| 验证 | 25 项后台任务全通过，缓存复用 0；R-30=pass、cached=false；原证书清单双签名、五载荷、文档/图形及网页快照通过 |
| 本地审计 | [清理实绩](../C-03-Website·Xycabin/build/cleanup-report.json)、[隔离证明](../C-03-Website·Xycabin/build/cleanup-isolated-proof.json)、[原目录批处理日志](../C-03-Website·Xycabin/build/cleanup-official-build.log)；build 证据不入 Git |
| GitHub 与 Release | 构建源码 [938470e](https://github.com/Xymircai/Xycabin-Addins/tree/938470efe79810f3fcd95e9275d561a3ee5693cc) 已普通推送至 main；[现有 Release](https://github.com/Xymircai/Xycabin-Addins/releases/tag/v0.9.9.3) 三附件已同批替换并回下载，大小与 SHA-256 均匹配本表；[远端核验记录](../C-03-Website·Xycabin/build/cleanup-publication-state.json)；原标签未移动 |
| 边界 | 仅源码布局探针修复无旧 HTML 时的构建依赖；功能源码、配置、必要 lib、证书、文档资源与真实验收资料保留；真实 WPS 点击与安装效果仍待实机验收 |

以下为历史交付记录，旧安装包摘要不再代表本轮产物。

本轮随迁移文档及内嵌资源重新编译 WordUpdater，22 项后台安全探针全部通过，含 Feature099 2450 项、内嵌文档核心 113 项和弹窗核心 67 项。业务代码及正式身份未改，本轮未签名打包、安装或发布。见[两产品验证状态](../C-01-ExcelIterate/build/local-delivery/quiet-replace/delivery-state.json)。

本轮随本地迁移源码合并重新编译 WordUpdater；共享登记、清单程序集身份与卸载流程改动一并覆盖。后台探针全部通过，含 Feature099 2450 项、弹窗核心 67 项、内嵌文档 113 项、程序集身份 12 项及卸载流程 54 项。见[本轮结果](../C-02-WordUpdater/build/local-delivery/upstream-integration/regressions.json)及[合并摘要](../C-02-WordUpdater/build/local-delivery/upstream-integration/integration-summary.json)。本轮未签名打包、安装、推送、发布或操作真实 WPS 文字文档。

本轮分类门禁实现提交 [3cff8c5](https://github.com/Xymircai/Xycabin-Addins/commit/3cff8c5856b6d32c2fc20a10994dd2228cb80892)；构建源码与 README 提交 [bfccdec](https://github.com/Xymircai/Xycabin-Addins/tree/bfccdecb9bc3dab0af0030def01d1f9aac6a07ae)，原有四个未推送提交完整保留并一并推送。详见 [本轮验证与发布记录](../C-03-Website·Xycabin/build/github-release-r30/delivery-review.md) 与 [完整证据摘要](../C-03-Website·Xycabin/build/github-release-r30/delivery-summary.json)。

以下为本地迁移开发期间的历史验证记录，独立于上述远端发布记录；不代表本轮合并已重新发布。

本轮随共享弹窗登记同步编译 WordUpdater；弹窗核心 67 项、Feature099 2450 项、内嵌文档核心 113 项及功能区探针均通过。Word 业务功能无修改，本轮未签名打包、安装或发布。

后续本地完整重建后的包摘要已变化，当前安装包以[远端发布校验记录](../C-03-Website·Xycabin/build/github-release-publication/publication-state.json)为准；原优化交付摘要和性能汇总保留为历史实验，未冒充本次重新计时。

本产品 31 个末级入口通过 [B3 R-30](<B3-Spec Coding.md#6-功能源码分类与代码索引>) 与 [B4 代码索引](<B4-Code Index.md>) 核验。完整与增量模式均只读检查；55 项正反例包含 16 次实际入口违规阻断，不自动修正目录或索引。

| 检查 | 结果 |
| --- | --- |
| 安装包 | [WordUpdater_Setup.exe](../C-02-WordUpdater/AA-Program/WordUpdater_Setup.exe)，4,248,576 字节，版本 0.9.9.3；SHA-256：597d5b0fa1e841f5e410921d67a48289a9aff4ca33a108f7830216b6b5a7f291 |
| 重建与后台 | [联合完整重建](../C-03-Website·Xycabin/build/last-run-summary.json)与[收尾独立验证](../C-03-Website·Xycabin/build/validate-run-summary.json)成功；本产品 24 项完整后台任务通过、缓存复用 0；[产品验证日志](../C-03-Website·Xycabin/build/WordUpdater-validation.log) |
| 分类与编译 | [验证收据](../C-02-WordUpdater/build/validation-timings.json)记录 R-30=pass、cached=false；目录、实际回调、编译清单及 B4 一致 |
| 签名与载荷 | [重建及签名日志](../C-03-Website·Xycabin/build/WordUpdater-rebuild.log)与[五载荷和资源核对](../C-03-Website·Xycabin/build/WordUpdater-payload.log)通过；沿用原发布者、清单双签名，文档与图形资源匹配 |
| 接口兼容 | [公开接口对照](../C-03-Website·Xycabin/build/github-release-r30/public-api-summary.json)：327 条记录保持一致，新增 0、删除 0 |
| GitHub 与 Release | 已普通推送至 main；[现有 Release](https://github.com/Xymircai/Xycabin-Addins/releases/tag/v0.9.9.3)附件已替换并回下载，摘要匹配本地正式产物；HTML 内嵌 EXE 与独立附件逐字节一致 |
| 源码链接 | [本轮精确提交](https://github.com/Xymircai/Xycabin-Addins/tree/bfccdecb9bc3dab0af0030def01d1f9aac6a07ae)及[对应源码 ZIP](https://github.com/Xymircai/Xycabin-Addins/archive/bfccdecb9bc3dab0af0030def01d1f9aac6a07ae.zip)；既有版本标签保持历史指向，自动 Source code 附件仍是旧标签源码 |
| 验收边界 | 实际 WPS 点击、宿主交互、安装效果、焦点与跨屏待实机验收；本轮未安装、连接或关闭 WPS，未重置用户配置 |

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

完整交付带三个同级项目、父verify_shared.py、两加载项lib、原证书私钥和项目内资产；仅重建单加载项也带完整Website。环境：Windows、可执行Python≥3.10、Node.js、.NET Framework4.8/C#5。先实际--version核解释器，子进程参数列表/UTF-8；普通构建不需网络或FFmpeg。

预览运行 `python AA-Program/C3-XycabinSite_Configure.py`，只改HTML；仅使用成功交付记录中版本/大小/摘要匹配的包，无记录或变更时默认下载不可用。不要手改生成HTML或把配置器移出固定目录。

联合交付直接运行本项目AA-Program批处理：1增量、2完整（均--background），或 `python rebuild.py --background --mode full`，不必先跑两单项目批处理。随后 `python validate.py --background`；回填审计后 `python src/tools/layout_probe.py --strict-evidence`核引用。安装、真实file下载/播放、WPS点击单列验收。

## 3. 当前状态与交付

### (1). 当前实现与待验收

本次加载项内嵌文档改为整份导出：开发预览打开整份 B2-DevExperience.md，用户预览打开整份 A-UserGuide.md，不再分章裁剪、拼接或改写图片路径，导出文件名与源文件同名。加载项升至0.9.9.3并沿用原发布者；网站随联合交付重新生成，下载页携带 0.9.9.3 两包，站点版本摘要与 B1 版本表逐字一致，指南依据两使用说明同步。

源码与正式产物内容指纹控制增量；两包、独立页和交付记录统一事务发布。后台前端/下载/音频替身不发声、不真实下载，不能代替file页面、听感、浏览器许可或安装效果。当前交付摘要在审计区，方法不重复历史流水。

### (2). 本轮交付证据

**最新交付：按 R-21 完成三项目冗余清理并重建。** 隔离副本不携带旧 build、生成 EXE/HTML 或字节码缓存；原批处理选择 2 完成联合完整重建及收尾独立验证。清理后在原目录再次运行同一批处理，返回 0；原证书、0.9.9.3 与稳定 ID 沿用。

| 项目 | 删除文件数 | 实际删除字节 |
| --- | --- | --- |
| C-01-ExcelIterate | 18 | 20,148,605（19.22 MiB） |
| C-02-WordUpdater | 2 | 8,494,080（8.10 MiB） |
| C-03-Website·Xycabin | 1009 | 187,838,793（179.14 MiB） |
| 合计 | 1029 | 216,481,478（206.45 MiB） |

清理共 55 个目标，失败 0；类别为构建暂存、清单测试夹具、历史生成 HTML 与有对应源码的 build 内 Python 字节码缓存。网站删除 12 个 stage 目录、3 份历史/探针 HTML 和 31 处字节码缓存。完整路径与逐文件摘要见[应用后的 schema 3 清单](../C-03-Website·Xycabin/build/cleanup-report.json)，[隔离成功证明](../C-03-Website·Xycabin/build/cleanup-isolated-proof.json)绑定真实批处理日志及输入指纹；未伪造 fixture-state。

保留网站 `build/site-downloads.html`，因现行 rebuild.py 存在引用且门禁拒绝删除；保留 `document-refactor-20261012`、`grid-output-format-20261014`、`file-names-layout-filter-20261014`、`output-rename` 中的源码 ZIP、独有历史与验收资料，以及 `github-release-*` 的发布恢复证据。Excel 的 `optimization-audit`、实机日志/截图、两个 `.cer` 证书、两产品全部 lib 及锁定 PDF 包保留；源码区域缓存按专项入口的保守边界保留，未确认恢复资料不删除。上述历史目录仅删除单独核准的字节码缓存。

三目录包含新产物和证据后的总字节由 656,350,848 降至 443,201,825，净减少 213,149,023 字节；这是本地重建完成时的快照，后续审计日志会略有增长。根 build 另保留旧 Release 备份。两个本轮自建临时隔离目录约 95.1 MiB，自动审批拒绝删除（blocked by policy），因此保留，不计入项目内释放量。

当前 [XyCabin.html](../C-03-Website·Xycabin/AA-Program/XyCabin.html) 为 13,973,932 字节，SHA-256：8808ae597e62357feeccbfd6b74193550c7c3d3cdd9418479119cde076dde165；HTML 两包与独立 EXE 逐字节一致，两 DLL 内嵌同一无安装包网页。联合重建 14 个阶段与收尾验证 11 个阶段均返回 0，[重建摘要](../C-03-Website·Xycabin/build/last-run-summary.json)及[验证摘要](../C-03-Website·Xycabin/build/validate-run-summary.json)均 success；Excel/Word 分别 30/25 项完整后台任务、0 缓存复用。

本轮源输入复核：除五份配置生成文件的时间戳外，隔离输入副本对应的源码、配置、依赖和资源字节均保持一致。新发现的 SourceLayoutProbe 无旧网页构建问题已修复并归并 B2。构建源码 [938470e](https://github.com/Xymircai/Xycabin-Addins/tree/938470efe79810f3fcd95e9275d561a3ee5693cc) 已普通推送至 main，README 可在 GitHub 网页查看；[现有 Release](https://github.com/Xymircai/Xycabin-Addins/releases/tag/v0.9.9.3) 的标题、说明和三个附件已更新，回下载 SHA-256 与本地一致，HTML 内嵌两包及网页正文核验通过，原标签未移动。见[发布状态](../C-03-Website·Xycabin/build/cleanup-publication-state.json)。首次发布因说明文本的 Windows 换行比对失败，已精确恢复旧附件与原说明后修正并重发，失败与回滚记录保留在根 build 审计目录。本轮没有进行真实 WPS 安装或操作用户文档。

以下为历史交付记录，旧网页及附件摘要不再代表本轮产物。

本轮本地源码合并保留上游构建与发布流程，网站仅保留字体注册表非路径值的缓存兼容改动；CacheProbe 10 项与 WorkflowProbe 77 项通过。未重建发布网站或更新远端附件，既有远端发布记录如下。

本轮分类门禁实现提交 [3cff8c5](https://github.com/Xymircai/Xycabin-Addins/commit/3cff8c5856b6d32c2fc20a10994dd2228cb80892)；构建源码与 README 提交 [bfccdec](https://github.com/Xymircai/Xycabin-Addins/tree/bfccdecb9bc3dab0af0030def01d1f9aac6a07ae)，原有四个未推送提交完整保留并一并推送。详见 [本轮验证与发布记录](../C-03-Website·Xycabin/build/github-release-r30/delivery-review.md) 与 [完整证据摘要](../C-03-Website·Xycabin/build/github-release-r30/delivery-summary.json)。

以下为本地迁移开发期间的历史验证记录，独立于上述远端发布记录；不代表本轮合并已重新发布。

本轮补充字体注册表非路径值的缓存环境兼容处理，CacheProbe 10 项通过。网站配置及版本未变；构建用离线网页仅为本地预览，本轮未发布。

按用户后续授权，已将实现提交[102d8ec](https://github.com/Xymircai/Xycabin-Addins/commit/102d8ececaa8519001957ea4459e49987aa9d15f)推送到main，仓库代码与README已更新；现有[Release v0.9.9.3](https://github.com/Xymircai/Xycabin-Addins/releases/tag/v0.9.9.3)的两个EXE及XyCabin.html已替换，发布标题与维护说明同步更新。用户确认只更新HTML附件与仓库说明，不启用GitHub Pages。

三份产物同批联合完整重建，并通过收尾独立验证。B1 此审计区在构建完成、推送与远端附件回下载核对成功后回填，不改变内嵌文档正文。

| 检查 | 结果 |
| --- | --- |
| 独立网页 | [XyCabin.html](../C-03-Website·Xycabin/AA-Program/XyCabin.html)，13,384,792 字节；SHA-256：9e300daa5991206ef91b4d85cdc22da5c3baac68e8d170f820a6522486e60bdc |
| 联合重建 | [运行摘要](../C-03-Website·Xycabin/build/last-run-summary.json)状态 success，14 个阶段返回 0，G-04、G-08、G-09、G-17、G-19、G-20、G-21、G-22 通过；R-30=pass、cached=false |
| 收尾验证 | [独立验证摘要](../C-03-Website·Xycabin/build/validate-run-summary.json)状态 success，两个加载项 26+24 项后台任务通过且不复用缓存；签名、五载荷及内嵌文档/图形资源通过 |
| 附件一致性 | [发布验证状态](../C-03-Website·Xycabin/build/github-release-r30/publication-state.json)为 PUBLISHED_AND_DOWNLOADED_VERIFIED；三个远端附件已重新下载，SHA-256 与本地匹配，HTML 内嵌两包与独立附件逐字节一致，去除安装包数据后的网页与两 DLL 快照一致 |
| 发布记录 | [现有 Release](https://github.com/Xymircai/Xycabin-Addins/releases/tag/v0.9.9.3)说明已更新本轮变化、精确源码提交、[源码 ZIP](https://github.com/Xymircai/Xycabin-Addins/archive/bfccdecb9bc3dab0af0030def01d1f9aac6a07ae.zip)与三附件完整 SHA-256；原说明和附件已备份，未移动标签或强制推送 |
| 版本与边界 | 保持 0.9.9.3、原发布者及既有标签；自动 Source code 对应历史标签。本轮未自动安装或关闭宿主，实际 WPS 按钮和安装效果仍待实机验收 |

