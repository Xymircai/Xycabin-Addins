**可复现开发经验**

# 一. Word与Excel共享开发经验

本章规定两个 VSTO 加载项的环境、功能接线、窗口线程、配置和验证方法。后三章给出 Excel-Iterate、WordUpdater 与 XyCabin 的数据规则和验收步骤。接手时按 [项目档案](B1-Project-Files.md) 选择范围，按 [代码索引](<B4-Code Index.md>) 定位实现。

## 1. 从零构建加载项

### (1). 固定环境与产品身份

1. 在 Windows 上运行 `py --version` 和 `node --version`。使用 Python 3.10 以上、.NET Framework 4.8 及 `%SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\csc.exe`，按 C# 5 编译。Python 子进程沿用 `sys.executable`、参数列表和 UTF-8，路径以脚本目录为基准。
2. 按产品 `project.json` 的 `lib` 配置备齐六种引用。核对 VSTO Tools、Common、Common.v4.0.Utilities、宿主 Word 或 Excel、VSTO Runtime 和 Office.Core。安装验收使用完整 VSTO Runtime。
3. 以程序集分别确定安装路径、挂载、用户配置及白名单。Excel 使用 `Excel_Iterate` 与 `ET` 白名单，Word 使用 `WordUpdater` 与 `WPS` 白名单。Office 挂载位于 `HKCU\Software\Microsoft\Office\<宿主>\Addins\<addin_key>`，Kingsoft 的 `AddinsWL` 写入值名 `addin_key` 和空字符串。新产品的 `legacy_name` 留空。
4. 在 `ThisAddIn` 中绑定对应宿主工厂、`[StartupObject(0)]` 和 `AddInBase`，通过 `CreateRibbonExtensibilityObject` 接入功能区。Excel 兼容探活顺序为 `KET.Application`、`et.Application`、`Excel.Application`，用 `Workbooks` 检查类型。Word 对应 `WPS.Application`、`KWPS.Application`、`Word.Application` 与 `Documents`。现场回归显式选择已注册的 WPS 宿主。

### (2). 建立编译与安装闭环

1. 用信息按钮、只读窗口和一项回调建立最小功能。将模型和数据规则交给引擎，将界面和宿主读写交给各自的调用层。依次覆盖正常、空输入及异常路径。
2. 在 `project.json` 登记 `sources`、`tools`、`additional_sources`、功能索引及退出清理。通过 `@shared/` 引用 `C-00-Shared`，将共用实现直接编译进两个产品。用配置器生成 `.g.cs`。更新功能映射后执行 `py -B -X utf8 C-00-Shared/feature_index.py --write`，再运行 `py -B -X utf8 verify_shared.py`。
3. 用 `GetPkt --json` 从 PE 读取程序集名称、版本、文化、架构和公钥 token。`gen_manifest.py` 按实际标识及 AssemblyRef 写清单，`verify_package.py` 据此校验绑定和签名。用 `PackageIdentityProbe` 覆盖强签、弱签、名称与加载入口变体。
4. 读取 `C-00-Shared/signing/signing.json`，核对统一发布者 `CN=Xymircai WPS Addins Dev`、指纹及私钥。证书缺失时核对本机 PFX 摘要，从标准输入传入口令并导入当前用户个人证书库，再复核指纹、有效期及私钥。将 PFX、口令配置和私钥保留在本机。编译后重签应用和部署清单。
5. 打包五个 VSTO 载荷。用 `InstallFlowProbe` 验证注册和卸载顺序。`MaintenanceEntry` 登记显示名、版本、位置及带引号的本地维护工具命令。显式 `--uninstall` 进入卸载流程，官方订阅注销成功后移除本产品载荷、挂载、白名单及卸载项，保留用户配置和 VSTO Runtime。

### (3). 正式交付与经验维护

1. 固定业务源码、正文、资产和签名配置。按 [R-17](<B3-Spec Coding.md#2-重建签名与交付>) 运行本次涉及的产品 BAT。分别核对 Excel G-11／G-12、Word G-14／G-15 与实际退出码。用户明确要求网站联合打包时，运行网站 BAT 并选「2. 完整验证」，核对 18 个构建关口与 G-23。产品证据写入各自的 `build`，联合回执写入站点的 `build`。
2. 按运行标识保存重建、验证及收尾阶段日志。复用 `build_report`、`tool_catalog` 和产品 `project.json`，记录 profile、关口编号、状态、耗时、日志路径、缓存原验证时间及跳过原因。后台子进程使用 `CREATE_NO_WINDOW` 和 `stdin=DEVNULL`。真实 WPS 探针通过 `LiveProbeSafety` 确认宿主归属与样本目录。
3. 每次更新本文件时，将经验写成可执行步骤、固定参数、触发条件、恢复动作或输出字段。全文使用正向陈述句和明确动词。用最小样本及独立预期值保留有效边界，把单次运行、产物和现场验收结果写入 [B1 本轮交付](B1-Project-Files.md)。用有效做法替换过期做法。回填审计节后核对剔除审计节的 B1 正文摘要，再运行 `py -B -X utf8 verify_devlog.py`。

## 2. 功能区按钮、图标与菜单

1. 在产品配置器维护 `BUTTONS`、`EXTRA_TABS`、中文标签、组 ID、回调和帮助文字。沿用稳定 ID，通过 `Connect` 对应业务窗体或引擎。在 `project.json` 登记叶子对应的目录、入口和依赖，生成 B4 索引后核对功能树。
2. 对可排序菜单生成 `dynamicMenu` 内容。通过回调读命令 ID、显示名、可见性及顺序，复用默认的 Ribbon 项和局部图片路径。缩略图以透明背景和目标像素尺寸输出，在 100%、125%、150%、200% DPI 下检查边缘、对齐和高对比模式下的辨识度。
3. 保存名称或顺序时，先核对外部改动，再原子更新 XML，由 `CustomizationSaving.AfterSave` 投递到宿主 STA 刷新已登记 Ribbon。记录配置保存结果和 Ribbon 刷新结果。用真实菜单核对初始、隐藏、重启和空配置回退。

## 3. 弹窗美工、控件与网格

先选布局容器，再选数据控件。用根 `TableLayoutPanel` 分组，卡片内容按实际测量定位，按行增减的底部按钮随容器布局。自动高行里的卡片用 `AutoSize=true` 与 `Dock=Top`，用内容期望高度计算卡片高度。需要伸缩的日志、表格和预览使用有最小尺寸的剩余行。检查切换字体、DPI 和窄窗时的可达操作。

| 界面对象 | 实现与检查 |
| --- | --- |
| 表单、卡片及输入 | 按字段语义选择 Label、TextBox、ComboBox、NumericUpDown、Button 和 CheckBox。按最长有效文本测量标签、表头与按钮。 |
| 数据表与菜单树 | `DataGridView` 承载行和列、`TreeView` 承载层级。可编辑下拉用 `DataGridViewComboBoxCell` 保存每行选项，`Tag` 绑定业务身份。 |
| 进度、状态及临时菜单 | 区分持久结果和临时提示。将选项草稿保存在对象身份上，从确定事件统一提交。给 Esc、关闭、外部点击及失去焦点定义收尾。 |
| 网格宽度及滚动 | 测量文字、字距、箭头、最小列宽与 DPI。窄窗启用横向滚动。先核对句柄及 `IsDisposed`、`Disposing`，在 `OnHandleCreated` 消费缓冲日志。抵消矩阵达到 5000 行时核对虚拟列表的 `CellValueNeeded`。 |

按结构、控件和交互分别查看示意，并用实际 DPI 核对可见区域。

![表头匹配合并的结构与容器](../C-01-ExcelIterate/AB-Graphic/界面对象-表头合并-编号示意.png)
![表头匹配合并的选择、按钮与拖动](../C-01-ExcelIterate/AB-Graphic/界面对象-表头合并-交互操作示意.png)
![选项与网格操作](../C-01-ExcelIterate/AB-Graphic/界面对象-选项与网格操作.png)
![树与分区](../C-01-ExcelIterate/AB-Graphic/界面对象-树与分区.png)
![反馈与临时菜单](../C-01-ExcelIterate/AB-Graphic/界面对象-反馈与临时菜单.png)
![通用选择与输入](../C-01-ExcelIterate/AB-Graphic/界面对象-通用选择与输入.png)
![通用列表与展示](../C-01-ExcelIterate/AB-Graphic/界面对象-通用列表与展示.png)

## 4. 宿主线程、窗口寿命与安全预览

### (1). COM 与隔离窗口

1. 在加载项宿主 STA 通过 `ThisAddIn.ItemProvider.GetHostObject` 获取当前宿主实例，同一线程读取选区、访问 COM、提交业务并释放自取引用。独立线程接收可序列化值。需要外部发现时用 ROT 与完整文件路径核对，区分借用接口和自取引用。
2. 同步确认经 `HostDialogService.ShowIsolatedResult<T>` 在独立 STA 创建、显示、取值和释放。使用 `CoWaitForMultipleHandles` 等待回传，保持宿主消息处理。为选择器、取消、重复提交、关窗及异常定义结果。宿主调用留在原委托线程。
3. 用 `HostBatchRunner` 在宿主 STA 分批读取，单批最多 4096 行及 65536 格，待解析最多两批。用 `SourceChangeWatch` 监听相关文件或工作表，块间和提交前复核身份。关窗及卸载取消会话，用会话身份过滤迟到回调。监听不可用时切换到原线程同步读取并记录原因。

### (2). 配置实际窗口与预览

1. 新窗口依次登记 `DialogCatalog` 稳定 ID、96 DPI 默认和最小客户区、`DialogSizing.Prepare`、`DialogPreviewFactory`、`DialogTree` 对应 Ribbon 叶子及加载项退出清理。预览使用固定样例，在创建和加载阶段隔离宿主读写、业务、链接跳转和高亮。
2. `DialogSizes.xml` 按程序集隔离。禁用 DTD，验证版本与范围，锁内比较摘要并原子替换。损坏值恢复默认，未来版本以只读方式呈现。使用 `ResizeBegin`、`ResizeEnd` 捕获用户尺寸，显式保存草稿。以实际文字宽度和缩进计算目录命令列，绘制与命中共享坐标。
3. 在独立 STA 维护格式预览及字体寿命。隐藏源窗口初始化一次，依次应用样式、重排、捕获、同步选择矩形。`DrawToBitmap` 捕获控件，RichTextBox 通过 `EM_FORMATRANGE` 补绘并释放资源，文本框按实际样式补绘且保留密码遮罩。用相同裁剪快照处理命中与绘制。以实际可见的控件、默认尺寸和最小尺寸通过验收。

### (3). PDF 当前文档与后台转换

1. 用 `Connect.Pdf`、`PdfToolsForm`、`PdfCurrentDocument` 及固定版本引擎接线。用 `WPS_WORD` 区分宿主。登记 `file.pdf.current`、`convert`、`watermark`、`split`、`merge` 对应的完整窗口 ID。
2. 在宿主 STA 捕获文档、Excel 原工作表与连续选区，按 HWND 和 PID 核对归属。执行时选择输出目录，核对对象、活动目标、路径和可保存状态，再调用宿主保存。新文档使用原生保存对话框。保存取消或文件校验失败时停止。
3. Excel 按整簿、原工作表或原连续选区导出。选区导出忽略打印区域并保持原工作表打印设置。Word 导出全篇后按需提取页码范围。后台接收临时 PDF 路径。批量转换前检查 Office 源占用，遇到 `RPC_E_CALL_REJECTED` 或 `RPC_E_SERVERCALL_RETRYLATER` 提示暂停，交还原窗口。
4. 在自建临时目录完成批量转换或合并，全部通过再发布，同名自动编号。失败或取消清理本批临时输出。水印接收 RGB、居中或平铺。禁用请求和结果 XML 的 DTD。
5. 运行 `PdfToolsProbe`。执行 `PdfCurrentDocumentProbe --live` 前核对新实例的 HWND、PID 和空文档集合。处理分页样本时记录 WPS 版本、打印范围、缩放、源行高列宽、输出页数及问题所在页。用户步骤和 PDF-001 见 [统一使用说明](A-UserGuide.md)。

## 5. 四种拖动及配置生效

| 操作 | 执行规格 |
| --- | --- |
| 输入与按钮拖宽 | `WidthGrip` 绑定控件，按 DPI 将逻辑步长换算为像素，向外部容器请求重新布局。检查松开后按钮可见及提示框右边界。 |
| 卡片区块改高 | `RowGrip` 调整根表格布局行的高度，按最小内容高和相邻行可用空间限制，释放鼠标时以 96 DPI 逻辑高度保存，拖动失去捕获时恢复原行高。 |
| 格式预览控件缩放 | `PreviewSelectionOverlay` 以稳定控件键保存 `Width96`、`Height96`。范围 1～4096，0 使用默认尺寸。ComboBox 与 UpDown 仅调宽度。先复位原布局属性，再应用文本和字体、解析尺寸、重排并捕获。 |
| 命令顺序拖动 | `SortableCommandTree` 仅移动同一父级的可排序命令，保留功能和菜单 ID，释放后更新草稿并校验相邻顺序。 |

保存名称和顺序后刷新当前 Ribbon。格式、控件宽高及窗口尺寸在下次开窗时由 `FormatSession.Capture` 固定快照。用 `LiveCustomization` 去抖文件变化并将菜单刷新交给宿主。默认 Font 按实例克隆，字体登记字号 6～36，缺失时回退。保存不同设置时合并对应字段，取消时丢弃草稿，损坏配置保留最近有效快照。

## 6. Agent-Evo、内嵌文档与网站

1. 通过 `EXTRA_TABS` 登记 Agent-Evo。默认隐藏，主按钮 `get_pressed` 与选项卡 `get_visible` 读取同一状态。点击后修改、刷新并原子写入 `AgentExtension.state`。刷新失败恢复旧状态，保存错误进入提示。
2. 概览取登记位置的首段。将 A 与 B2 各整份内嵌，去 BOM 并统一换行。保留 `C-04-Devlog` 逻辑路径和指向 `AB-Graphic` 的相对图片链接，按摘要在 `WpsShared.Agent.Assets.xml` 登记现用资产。打开全文时按程序集与文档资产摘要修复缓存，随后通过 `SHOpenWithDialog` 选择应用。
3. 「开发预览 → 全文」经 `AgentDocumentAccess` 和 `DeveloperPasswordForm` 每次验证输入。使用掩码、校验摘要和当次窗口作用域。取消、Esc 或校验失败时停留在概览。用 `AgentDocumentCoreProbe` 核对允许、拒绝、缓存及导出路径。
4. 两个 DLL 用同一个 `WpsShared.Agent.Home.html` 空下载快照。缓存写到 `LocalAppData/Xy/AgentEvo/Home/<摘要>/index.html`，用 HTTP 或 HTTPS 关联选择浏览器。独立离线页在同一快照上附两个安装包。联合构建后在 B1「本轮交付证据」登记资源核验结果。
5. `RequirementsForm` 将 `Requirements.txt` 用 UTF-8 原子写入并核对外部修改。普通关闭提供保存、放弃、取消。宿主退出保存单份未保存草稿，下次打开提示恢复。将正式保存结果和草稿清理结果分别呈现。


# 二. Excel-Iterate 专属开发经验

本章规定工作表事务、外链、文件、审计底稿和批量数据操作的实现与回归。每个操作以当前工作簿身份和显式范围为输入，以保存结果及恢复状态为验收依据。

## 1. 单元格、选区与公式

### (1). 引用追踪与高亮

用 `FormulaRefParser` 解析公式。跳过双引号常量，保留单引号内的工作表及文件，按跨簿、跨表、同表解析冒号、感叹号和美元符。按实际数组下界及 `UsedRange` 起点将公式索引换算成地址。沿稳定的簿、表、地址定位单元格，用 `TrackerProbe` 覆盖中文表名、绝对引用、区域、错误源和跨簿路径。

跳转前取得实际对象并记录原填充，黄色使用 `Interior.Color=65535`，无填充按 `ColorIndex=-4142` 恢复。约 2.5 秒、再次跳转或关窗时恢复。选中单元格的展示行保留同类 `Tag`，计数取引擎返回条数。

### (2). 批处理、公式与数值

1. 用 `WorksheetIdentity` 捕获父簿、表及选区，按 RCW/IUnknown、FullName 和 Name 核对，CodeName 作为可选信息。按 Area 与 UsedRange 求交，分块上限 4096 格。对重叠格去重，合并区仅处理选中的主格。用 HasFormula 区分公式和以等号开头的文字。
2. 获取原公式、值类型、格式和保护状态。Formula2 属性缺失时回退 Formula。对数组、溢出、锁定或读取失败位置结束整批预检。提交前再次核对来源身份和快照，写入后读回。异常时将已尝试写入的格逆序恢复，再恢复 EnableEvents 和 ScreenUpdating。记录工作表、原地址、原值、失败阶段及待恢复项。
3. IFERROR 默认空文字，ROUND 默认两位，ABS 接收一项。用 `FormulaWrapper` 识别引号、数组、结构化引用与嵌套，添加一层完整封装或提取完整外层函数的第一参。引号双写，有限数值用不变区域性，ROUND 位数为有符号 32 位整数。
4. 从 `ARITHMETIC_DEFAULTS` 生成默认值和提示。加、减、乘用 100，除用 10000。原公式整体加括号，数值 `20` 可写成 `=20*100`。对文本、空值、布尔和错误跳过。除数为零时结束校验。由 `IFormulaValueCell` 保存原公式和值类型，按原接口恢复。
5. `FormulaInteraction` 在宿主 STA 捕获选区，参数窗只接收会话号、操作、Owner 和 `FormulaRequest`。确认投递至宿主，按会话号防重复或过期提交。关闭、Esc 和初始化失败释放会话。用 `FormulaInteractionProbe` 检查选区冻结、取消和撤销。

### (3). 列筛选与快捷高筛

按活动格所属的 ListObject、AutoFilter、CurrentRegion 顺序选择范围。正式表读 `HeaderRowRange`，普通区域取首行。以列身份绑定选项，搜索调整显示，应用至少保留一列。写入前核对范围、原表头、Hidden 与保护，按快照逆序恢复。

快捷高筛在同一连续表收集条件。要求有效且唯一的表头，同一行条件用 AND，不同行用 OR，去重并跳过全空条件行。自动补所选列的表头，将比较与通配条件交给 `AdvancedFilter`。新建结果表写值与数字格式，重名编号，失败时清理本次未完成的输出。

## 2. 工作簿、目录与表头匹配合并

1. 在首位生成目录，重名时编号，展示工作表名称和生成时状态。用隐藏名称绑定目标 Range，以 `SheetFollowHyperlink` 处理隐藏表的显示和跳转。为单引号表名双写引号，检查结构保护与目标删除。退出时解绑事件。
2. 通过 `MergeHeadersEngine` 匹配和规划，`MergeHeadersHost` 处理 COM、快照、读回及备份，`MergeHeadersForm` 展示源、目标、列映射和报告。接受已保存的 `.xlsx`。单批最多 20 个源、50 万格，超限分批。匹配时修剪表头首尾空白，忽略大小写。列缺失时标出供确认，每列独立从该列末个数据行追加。
3. 用 `FormulaRefRewriter` 跳过常量、函数、科学记数及定义名。同表相对列按映射，绝对列锁定，绝对行锁定，相对行按行差平移。跨表、跨簿保留前缀及列，相对行同样平移。未映射列沿用原列字母，越界保持相应轴并报告。
4. 执行前保存目标，在同目录「ExcelIterate合并备份」保存副本并捕获本批区间。提交后逐格核对，失败或取消从快照恢复。确认结束撤销状态并保留文件备份。用 `MergeHeadersProbe` 检查列匹配和公式，再用 `MergeLiveProbe` 核对宿主的预检、写入、样式及还原。

## 3. 外部链接与文件批次

### (1). 扫描与映射

1. 按工作簿 `externalReferences` 的顺序关联 `.rels`，读取 `externalLinkPath` 和 WPS 的 `xlExternalLinkPath/xlPathMissing`。保留原 Target 和实际编号，外链名称按工作簿级与工作表级区分。只有名称定义引用的来源写入「仅名称管理器定义引用（X 个名称）」，按定义数量计数。
2. 绑定接收工作簿和程序拥有的「外部链接」表。执行前保存相关工作簿及清单，核对四列表头和前两列的身份集合。允许排序并按原身份读 C 列，行数、重复、身份或文件状态异常时提示重扫。新工作簿通过原生保存流程获得路径。
3. C 列先 Trim。空白行跳过，精确文本 `断开链接` 进入断链。替换接受完整 Windows `.xlsx` 路径及 `目录\[文件名.xlsx]`，只在终端文件名去除方括号。将 URI、完整路径和单独的文件名归一为匹配候选。源只有文件名时，同时核对保存包 Target 和 WPS 原始来源令牌，要求唯一匹配。调用 `ChangeLink` 时传宿主返回的原始 `Name`。
4. 用 `File.GetAttributes` 验证新路径，现存文件再验证 XLSX 包。`FileNotFoundException`、`DirectoryNotFoundException` 标记为缺失目标并继续。目录、权限异常、损坏文件、指向自身或冲突映射进入失败提示。相同绝对目标标记「链接路径已一致，目标文件当前缺失」或「链接相同，无需替换」。按链式末端优先规划，多对一、路径占用及环路暂停。

### (2). 写入、核验与撤销

1. 含断链、待核对来源或缺失目标的整个工作簿走 `LinkWorkbookCheckpoint`。在修改前保存备份、`BeforeHash`、预期提交摘要和操作意图。打开的工作簿使用所属宿主，关闭的工作簿在同目录暂存并以 `UpdateLinks=0` 打开，验证 FullName 后操作。关闭自己的暂存簿前核对新增编辑及摘要，校验后用 `File.Replace` 提交。遇 Win32 32、33、1175 时，在原文件和副本均存在的条件下最多尝试五次，间隔 50、100、200、400 毫秒。提交与整簿恢复均在每次尝试前核对两份摘要。摘要变化时暂停并保留备份。诊断登记阶段、尝试次数、工作簿路径及 HRESULT。
2. 缺失目标执行宿主原生 `ChangeLink(Name, 新路径, 1)`，保存后核对新源的规范化绝对地址、其他来源、单元格公式和定义名称。通过工作表部件及地址索引公式，对照替换前后的实际来源，允许 WPS 将新源保存为 `xlPathMissing`。核验通过写入「链接已替换，目标文件当前缺失」，成功行纳入结果。文件补齐后在 WPS 编辑链接中更新值。
3. 现存目标的普通替换沿用原事务。已开簿原生换链并核保存关系。关闭簿修改 `externalLinkPath`，核暂存产物与原摘要后原子提交。运行前复核存在性，遇目标状态变化时按当前批次安全路径处理或提示重扫。
4. 断链按选定来源检查名称依赖与作用域。根据公式中的美元符和 A1 地址跨度区分定义名称。`Y$5`、`$Y$5`、`Y:Y` 按地址解析，`=Y`、`=工作表名!Y` 按名称处理。混合来源、动态引用、作用域歧义或无法证实安全时暂停。核对 `RefersTo` 后删除只指向该来源且无依赖者的名称，再使用 `BreakLink(Name, 1)`。保存后核该源、余下名称及公式。
5. 对每行记录原来源、新路径、阶段、成功/跳过/失败、目标是否缺失及原因。用户日志按事件输出，诊断日志附 HRESULT、请求参数及隐式保存阶段。核验无法完成时保留意图、备份和实际状态。整簿撤销要求关闭相关工作簿并匹配提交摘要，恢复执行前的完整字节。普通替换按逆序还原已成功项。全部恢复或确认保留后，回收本批记录。失败时保留可重试状态。
6. 用 `WorkbookToolsProbe` 和 `NameLinksProbe` 覆盖普通或待核对来源、绝对路径与相对文件名、真名称与地址、缺失目标、现存目标、空白行和执行期间变更。实机准备独立 XLSX 副本，验证保存重开后关系、公式、定义名称、补齐文件后刷新取值以及撤销摘要。按成功写入缺失目标行数、公式保留格数和逐字节恢复工作簿数登记 B1 回执。

### (3). 整理、重命名与记录

1. 整理先生成只读计划，列出纳入文件、来源目录和备份数量。递归跳过联接和程序目录，提交前核时间、摘要及目标身份。预留编号，保留较新文件，同时间优先现有目标。淘汰项进入保持相对路径的本批备份。
2. `OperationBatch` 记录意图、实际路径、摘要、目录身份及清理进度。取消时只移除本批新建且为空的目录。确认按清单回收原目录，保护分类输出、原备份、联接和执行后新目录。
3. 恢复记录持久化使用同一次完整序列化，遇 Win32 32、33、1175 且原记录与临时文件均在时最多尝试五次，重试间隔 50、100、200、400 毫秒。报告记录路径、HRESULT、尝试次数及原异常，失败保留快照。用 `FileOperationsProbe` 验证暂时占用、持续占用及重新打开后恢复。
4. 批量重命名扫描第一层，按路径、原名、新名和状态绑定清单。只改主体时保留扩展名。原名、非法名称、碰撞和外部变更进入预检。大小写变化或交换使用唯一临时名两阶段重命名。失败时按已提交身份逆向恢复。

## 4. 文件提取名称与云端 OCR

1. 用 `FileNameEngine` 按文件类型和规则扫描，在宿主 STA 通过 `FileNameWorksheet` 写入结果。输出路径和文字列使用文本格式，状态及计数按实际结果回填。独立保存 OCR 前缀版本、四类规则及源身份，旧版规则按 v1 字段映射到 v2。按既有 `OcrPrefixRules.xml` 恢复配置。
2. 百度接口为 `general`、`accurate`、`general_basic`、`accurate_basic`。标准版长边 4096、编码后 8 MB，高精版 8192、10 MB，短边至少 15 像素。按接口发送 Base64 表单与适用粒度参数，PDF 使用 `pdf_file` 和 `pdf_file_num`。
3. 阿里使用杭州端点 `RecognizeAdvanced` 和版本 `2021-07-07`。按 ACS3-HMAC-SHA256 对查询、请求头和原始请求体摘要签名。请求体发送图片，尺寸 16～8191、长宽比小于 50、体积不超过 10 MB。多页 TIFF 按页送入。
4. 上传前核对服务商、服务版本、样本、页数和配额，串行每页 120 秒。仅当用户选择记忆时写入 DPAPI CurrentUser 加密的独立凭据文件。百度 token 按 `expires_in` 缓存，110、111 重新鉴权一次。网络、限流或配额异常时暂停当前队列并报告，取消时停止后续页。
5. 预览六列，按 Unicode 字符边界应用前缀和匹配，记录文件、页码、耗时、错误码、命中规则与输出状态。离线运行 `Feature099Probe` 及 `MultiplePrefixesProbe`，云端实测使用获准的独立样本和对应服务账号。

## 5. 合并抵消

1. `EliminationEngine` 接受纯文本行，`EliminationHost` 复用读取、备份、快照与写入。通过 `EliminationKind` 区分往来和收入成本，以表头角色取得字段。往来余额的合并单位格从锚点填充，单位用实际全称精确配对，保留净额符号。
2. 往来按公司对、会计科目和年度汇总四项原始余额。资产净额为债权方资产减负债，负债净额为债务方负债减资产，差额为资产端减负债端。差异表 15 列，展示四个方向余额、两项净额、抵消金额、差额和备注。坏账准备随对应抵消项目全额冲回，写入独立两行。
3. 收入成本按核算主体、客商、年度三维汇总，年度降序。每组写收入 R、成本 C、成本还原 `(R-C)-U` 和存货 U，验算 `R=C+((R-C)-U)+U`。收入及成本取同口径单体金额。U 有映射列且组内有值时求和，否则逐行计算 `Σ((收入-成本)×未对外销售比例)`。保留负数方向。
4. 正式写入静态金额。对公式探针，映射 U 使用 SUMIFS，逐行加权使用 SUMPRODUCT 并使条件、金额数组同区间。扩展区间起于数据首行，例如 `$O$3:$O$1048576`。结果表引用实际差异表 M 列，先取得差异表名再生成公式。
5. 用独立的差异表与结果表记录目标。执行保存、同批快照、拆旧合并、清区间、写入及读回，复用本批备份恢复。金额列用会计格式，表头行高下限 26，结果表按 `#4874CB`、白字和宋体 11 写样式。`EliminationProbe` 检查金额、平衡、日期及范围锚点。宿主回归同时验证两表和恢复。

## 6. 批量迁移数据

1. 用户确认源工作簿、源区域和入选目标。`MigrationConfiguration` 生成独立配置快照，`MigrationPlanner` 据此规划，`MigrationNameIndex` 对有序目标建名称索引。预检结果按原行身份回填。单批源区域至多 100000 格。空白源需要显式允许及确认。预检记录范围、目标数、源非空格数和公式格数。
2. 用统一下拉保存六种方式「只写数值」「数值+格式」「只写公式」「公式+格式」「只写格式」「粘贴链接」。逐行重算仅限两种数值方式。用对象身份绑定下拉与目标筛选草稿，确定后批量应用并令预检失效。Esc、外部点击和执行锁定关闭编辑器。
3. 普通公式按行列差平移，保留美元符和第三方引用地址，映射源表显式限定并核对同名目标表。递归核对定义名称的作用域和依赖。对无法解释、循环、数组、溢出及结构化引用给出预检地址。粘贴链接用原始 `SourcePath`、源表及双轴锁定地址，转义单引号，合并区仅链接主格。保存重开核实际外链与计划。预设 v2 保存 `pasteLinks`，读取 v1 时沿用原 `keepFormat` 及默认值。
4. 默认由 `MigrationBackgroundSession` 启动已登记 WPS 进程，进程加入本次 Windows 作业并在私有桌面运行。加载项核对桌面与作业后，经本地管道传递本实例自动化入口。记录启动阶段和 Win32 错误。用户确认改用前台时，先检查占用及原文件摘要。窗体收尾发送一次通知。
5. `MigrationRequest.ReplaceTargets` 控制原文件替换或另存副本。替换目标要求原文件关闭、可写且摘要匹配。暂存件保存、重开及包级核验通过后，`File.Replace` 在同目录切换并留下「迁移备份/批次/原文件名」。另存使用独立名称并防覆盖。后续发生并发修改时按提交摘要保护用户版本。恢复只覆盖仍匹配本次提交的字节。
6. 按宿主 STA 管理网关和计时器。Tick 入口检查 `ticking` 并暂停计时器，退出时按会话状态重启。用 `SourceChangeWatch` 按对象识别 `SheetChange`、关闭和增删表，迁移以公式文本和常量作快照。自建读簿按完整路径映射确认归属，原文件继续执行保存要求。日志写事件、线程、匹配状态及失败阶段。只读包按 1900/1904 基准把 `t=d` 日期规范到测试副本，保留数字格式、合并各格样式、行高列宽与 VBA 载荷。
7. 运行 `DataMigrationProbe` 与 `DataMigrationLiveProbe --background-buttons`，用正式窗体消息泵处理多个目标，核对写入结果、包级保真、原件摘要和恢复。日期样本同时对比日历值和数值序列，公式对比统一逆序矩形坐标并保留美元符。

## 7. 批量生成过程表与单体汇总表

1. 用 `SummaryWorkbookBinding` 显式绑定选中的工作簿，核对 FullName、IUnknown 和绑定版本。独立选择汇总表、Lead 表头、汇总表头及模板列。更换工作簿清空预检。按表头归一化匹配「编码&简称」「全称」「选用」，以行号、编码和全称复查来源记录。
2. 按公司顺序生成 `SummaryYearColumn`，同公司下年度默认降序。单年度一公司一列，多年度一公司占年度数对应的列。以完整 `YYYYMMDD` 目录段应用 `{yyyy}`，保留月日。默认检查来源是否就位，生成逐公司逐年度清单，校验有效行列及合并区。
3. 备份汇总表，按第一行至模板列末行确定写入区间。先拆旧合并和清理，再复制值、样式、合并、批注及列宽，最后按快照恢复原公式并精确替换方括号中的模板公司文件名。跳过字符串及其他来源。含括号文件名时给完整引用加规范单引号，并登记提示。年度表头对齐值为 `-4108`，追加日志，按备份还原本批。
4. `CompanyWorkbookRules.NamingStem` 先按 `StringInfo` 从原名移除指定尾部字符，再清理识别到的 Excel 后缀并追加模板后缀。按完整路径核对保留名、碰撞及现存文件，预览与执行复用同一命名逻辑。读取用独立副本，生成用原始模板快照。只有保存重开并核对指定值和表清单后才发布。
5. 按已完成、成功、失败和剩余任务记录进度，百分比向下保留一位，末项结束后显示 100%。用 `SingleSummaryProbe`、`SummaryAnnualChecks` 和 `CompanyWorkbookProbe` 核对预检失效、列序、路径、取消和保存冲突。

## 8. 文本转换与同名工作表操作

1. 用 `ConversionRangeAddress` 构造地址并集，选区上限 100000 格。通过宿主 STA 的 `InputBox(Type=8)` 读取范围。空选区、跨簿非法输入及超限按原身份提示。预览以 250 毫秒去抖调度，按会话拒绝过期更新。
2. `BatchTextEdit` 按 Unicode 字符位置处理插入、截取及删除，跟踪公式原首字符。保留原等号标记时按公式写入，删除该标记时按文本处理。宿主以最多 1024 格的连续矩形读写，按类型与原样式分组。普通值经 Value2 写入并保留前导零。读回发现 WPS 消费既有单引号时仅做一次传输转义，再核字符串内容。
3. 公式赋值遇 `0x8FE3000B`、`0x800A03EC`、`0x80020009` 时报告实测阶段和位置，按事务恢复。保留其他保护与保存错误的原始分类。用原类型及内容逐格读回检查。
4. 同名工作表批量操作冻结对应工作簿、表和地址。预览分文字和底层数值，确认后可同批写格式。用 `WorksheetEditBatch` 记录原内容、格式及行高，恢复最近对应批次后读回。结果工作簿由用户保存。用 `CellConversionProbe`、`ConversionInteractionProbe` 与 `BatchTextEditLiveProbe` 核对预览、撤销和宿主边界。

# 三. WordUpdater 专属开发经验

本章规定逐格 LINK 的创建、来源实例发现、更新、定位与样本回归。按原目标选区、来源地址及真实字段结果验收。

## 1. LINK 字段与表格选区

1. 用 `LinkFieldCode` 解析和生成 LINK/DDE。`LinkPasteEngine` 在原文档尾部临时区域用 `PasteSpecial` 获取原生 LINK，清除临时区并恢复选区。原文档探测失败时使用自有临时文档，完成后放弃保存关闭。数据类型使用 RTF 1、文本 2、OLE 0。字段引用用 R1C1，如 `LINK Excel.Sheet.12 "C:\样本\Book.xlsx" "Sheet1!R1C1" \t \a`。
2. 源按 Areas 顺序逐区域先行后列，目标矩形先行后列。选区上限 500 格，源与目标格数相等。先核所属表及实际范围，再从 Selection.Cells 或 Range.Cells 取得选中单元格，按物理身份去重。无法读取单格、跨表、缺格、选区内合并或嵌套时结束预检。
3. 读取父表 WordOpenXML，定位所选 `w:tc` 的 `gridSpan`、`vMerge`、`hMerge` 及子表格。累加同一行前置格的 gridSpan，将物理 ColumnIndex 换算为网格坐标。允许选区外的合并或拆分。剪贴板探测后复核同一组真实 Cell，写入前逐格重取 Range，按 `End-1` 保留单元格结束符。
4. 当剪贴板只有包围盒且格数与目标有差异时，从已打开来源读取 Ctrl 多区域。按 Application.Selection、窗口 RangeSelection、窗口 Selection 收集候选。要求 FullName、工作表名、格数与包围盒吻合，分片互不重叠。对多个不同且同样有效的候选报告歧义。
5. 全选正文顶层表格时清理旧 EditableRanges，逐表 `Editors.Add(-1)`，调用 `SelectAllEditableRanges(-1)`，最后删除临时标记并恢复 ScreenUpdating。表单保护值为 2 时结束操作。用 `LinkMapProbe`、`LinkPasteTest` 覆盖几何、重建、来源值、更新及合并表头。`CtrlSelectionLiveProbe` 通过文件关联打开工作簿副本，在默认实例为空的条件下验证六格多区域链接。

## 2. 更新、查看、定位与高亮

1. 用 `WordHostStateScope` 完整读取 DisplayAlerts 和 ScreenUpdating，再启用静默操作。任何捕获失败时结束提交，部分设置失败时恢复已尝试字段并记录。全篇更新顺序处理，断开时逆序处理。
2. `UpdateEngine` 扫描选区，浮动及内嵌对象按锚点去重。文本字段更新成功后比较结果，以 Ordinal 比较去除边缘空格、制表、CR、LF、`\a` 和 `\0` 的字符串。单列图片和 OLE 的结果类型，记录更新成功数、失败数、文本变化数与位置。
3. 查找时按稳定特征重取对象，字段定位 Code 或刷新后的 Result，内嵌对象定位 Range，浮动对象定位 Anchor。窗口依次尝试 ScrollIntoView 兼容路径，成功后交还 WPS 焦点。
4. 连续定位先清旧闪烁，再读当前原色和标黄，由 `FlashHighlighter` 约 2.5 秒后还原。关窗恢复原色，覆盖无颜色 0、其他颜色和原本黄色 7。用 `LinkViewFormProbe` 校验窗体调用顺序，独立 WPS 样本验证字段重取和颜色恢复，实际功能区操作记录计时及焦点。

## 3. 源文件、ROT 与隐藏 ET

1. 通过 `EtSelectionReader` 枚举 ROT，按注册的 Application CLSID 和剪贴板中的来源完整路径筛选。绑定匹配应用和来源 Workbook，复核其 FullName 后取 Workbook.Application。将进入同一实例的多个路径去重。比较路径时折叠连续反斜杠并保留 UNC 前缀。
2. 仅附加已运行实例。读 Workbooks、Windows、选区及 Areas 时登记每次自取 COM 引用，逆序释放。用 `EtSelectionReaderProbe` 覆盖隐藏默认实例、文件入口、同名不同路径、属性读取错误和引用平衡。
3. `EtCleanup` 从进程快照记录新出现的隐藏或状态未知 ET PID，输出进程诊断及 `CleanedHiddenHosts=0`。保留进程供用户保存和关闭。用 `EtCleanupProbe` 覆盖已有 PID、并发实例、窗口状态缺失和枚举异常。

# 四. XyCabin 专属开发经验

本章规定单页离线站点的配置、交互、下载和声音验收。将页面、嵌入式快照和两个安装包作为同一交付批次核对。

## 1. 站点配置与接线

1. 在 `C3-XycabinSite_Configure.py` 集中维护 `SITE`、`TEXT`、`PRODUCTS`、`THEME`、`WORKFLOW`、`FUTURE`、`DOWNLOADS`、`DOWNLOAD_GUIDE` 与 `FEEDBACK`。版本从新到旧，摘要与 B1 对应版本表逐字一致，每版最多 200 字，各项连续编号。
2. 按页面区域在 TEXT 登记键，向模板的 `data-text` 或 JS 文字对象接线。文本转义，JS 配置用 JSON 序列化，动态文案用 `textContent`。主题色校验 `#RRGGBB` 后写 CSS 变量。反馈地址对应官网 HTTPS 或 `mailto`，外链使用 `noopener noreferrer`。
3. 下载链接接受无账号口令的 HTTPS。空配置禁用对应项。下载指南使用有替代文本的站内相对 PNG 和明确操作步骤。运行 `site_probe` 核对主题、导航、项目切换、指南搜索、版本筛选及空配置。

## 2. 布局、按钮与弹窗

1. 下载、版本筛选和局部操作使用对应主题配色。下载主按钮与三角按钮等高，按钮最小 45 px。反馈文本在宽屏呈一行，窄屏按容器宽度换行。
2. 下载主按钮先开指引，继续执行下载，取消关闭对话框。缺指引或浏览器缺对话框能力时走直接下载路径。下拉用 Esc、外部点击和切换项目关闭，初次焦点用 `preventScroll`。
3. 菜单常规宽 245 px，限制到可见视口。显示后清旧位移，用 `getBoundingClientRect`、`visualViewport.width` 和 `offsetLeft` 测边界，回退 `clientWidth`。左右各留 8 px，用 `translateX` 修正溢出，在 resize、scroll 重新测量。验收两种主题、两个产品、宽窄屏、打开后改宽、键盘及外部关闭。

## 3. 下载字节与离线验收

1. 构建空安装器快照供两 DLL 共同内嵌。`attach_installers` 仅写唯一 `installer-assets` JSON 区。`embedded_snapshot` 清空下载区后对照两 DLL 的网页内容。配置器单独预览时，仅使用成功交付回执中版本、大小和摘要匹配的包。
2. 按当前产品取 Base64，每 87380 字符解码为 Uint8Array，核配置摘要字段格式和实际解码字节数。创建 Blob 和临时 `a.download` 提交给浏览器。锁住重复请求，成功请求 60 秒后释放 ObjectURL，异常和 pagehide 立即释放。DLL 空数据页按百度、夸克顺序回退，均为空时显示提示。
3. `verify_installers` 在构建端核 SHA-256 并逐字节对照本轮 EXE。`download_probe.js` 对正式 HTML 的 Blob 做比对。现场以 `file://` 打开网页，各下载一包，用 `Get-FileHash -Algorithm SHA256` 对照独立包，记录浏览器版本、实际保存字节数、摘要和浏览器提示。
4. 用 try/catch 处理 localStorage。存储不可用时保留本次页面状态。剪贴板失败时提供可选文字。浏览器策略拒绝 file 访问时记录策略与实际入口。

## 4. 声音参数与异步播放

1. 点击音效为 640 Hz 正弦、70 ms、默认增益 0.40、上限 1.0，可信按钮点击限频 100 ms。背景默认 tranquility、音量 0.50、上限 0.90。保留 afternoon 与 tranquility 场景 ID，音效和背景分别调节后通过总静音和压缩器。
2. 项目内 `tranquility.mp3` 使用原音频 60 至 111 秒片段，128 kbps、44.1 kHz 立体声、首尾 200 ms 淡化。普通构建读取项目片段。Base64 转 ArrayBuffer 并经 `decodeAudioData` 解码，缓存后循环，暂停保留进度，换曲从头开始。
3. `xy-home-audio` v2 对旧音量百分比除以 2，一次迁移后保存 v2 和原音效开关。以用户选择的静音、零音量和曲目为准，失效曲目回到默认。
4. 自动播放和首次交互补播均开启时，在首次可信导航、指南或主题点击里先 resume，再异步解码。跳过搜索输入、音频控件及合成事件。用代次令牌处理暂停、静音、零音量、换曲和 pagehide，释放旧源和待调度任务。手动播放可重试。运行音频状态机探针，再在目标浏览器验证首次交互、暂停和换曲。
