**共享开发经验与三个项目的专属经验**

# 一. Word与Excel共享开发经验

本章只保留一份共享经验，两个加载项不再各自维护副本。共享的是加载、界面、配置和交付方法；Excel 单元格与工作簿算法见第二章，Word 字段与表格见第三章，网站前端方法见第四章与[B1 的 XyCabin 档案](B1-Project-Files.md#四-xycabin-项目档案)。

## 1. 从零构建加载项

### (1). 环境与必要输入

底座为 Windows、WPS VSTO、.NET Framework 4.8，Python 至少3.10驱动 `%SystemRoot%/Microsoft.NET/Framework64/v4.0.30319/csc.exe`，语法上限 C# 5。先实际执行 Python `--version`；子进程使用 `sys.executable`、参数列表和 UTF-8，路径由脚本位置定位，允许中文、空格和中点。

| 输入                                                                                                               | 最小职责                               |
| ---------------------------------------------------------------------------------------------------------------- | ---------------------------------- |
| project.json、AA-Program/C1-Exceliterate_Configure.py                                                             | 宿主/身份/sources/tools与唯一界面、版本、固定证书配置 |
| src/addin/ThisAddIn.cs、Connect.cs、InteropDefs.cs、AssemblyInfo.cs                                                 | VSTO入口、专属回调、COM声明、程序集GUID          |
| src/addin/UiKit.cs、HostDialogService.cs及所依赖公共源码                                                                  | 界面、宿主确认隔离；依赖以实际sources和工具清单为准      |
| src/configuration、project_support.py、build/rebuild/resign/gen_manifest/build_installer/verify_package/validate脚本 | 生成、编译、签名、打包与验证；文件名中的斜线表示各独立脚本      |
| src/installer、src/tools及业务探针                                                                                     | 安装流程、证书/清单工具、产品身份与行为检查             |
| lib、icons、三个AP输出目录                                                                                               | 六个编译引用、源图标、程序/美工/说明分区              |
| 完整同级Website、适用的父目录verify_shared.py                                                                               | 当前网站快照和公共副本检查；不使用旧build或历史范例补齐输入   |

六个lib为 Tools、Common、Common.v4.0.Utilities、对应Word/Excel宿主库、VSTO Runtime、Office.Core。Utilities与Office.Core随包分发；用户机仍须安装完整VSTO Runtime。公共文件选择以 `verify_shared.py` 的 SHARED_FILES 和当前编译依赖闭包为准，不按历史数量复制全部业务。

### (2). 先隔离产品身份

| 项目项                    | WordUpdater                                              | Excel-Iterate                                                |
| ---------------------- | -------------------------------------------------------- | ------------------------------------------------------------ |
| 程序集/命名空间/addin_key     | WordUpdater                                              | Excel_Iterate                                                |
| 现有兼容探活链（非WPS验收启动策略）    | WPS.Application → KWPS.Application → Word.Application    | KET.Application → et.Application → Excel.Application         |
| 探活对象/工厂/编译宿主库          | Documents / Word.ApplicationFactory / Word.dll           | Workbooks / Excel.ApplicationFactory / Excel.dll             |
| 清单host / HKCU Office挂载 | Word / Software\Microsoft\Office\Word\Addins\WordUpdater | Excel / Software\Microsoft\Office\Excel\Addins\Excel_Iterate |
| HKCU Kingsoft白名单       | Software\Kingsoft\Office\WPS\AddinsWL                    | Software\Kingsoft\Office\ET\AddinsWL                         |
| legacy_name            | WpsLinkUpdater（仅自身旧名）                                    | 空                                                            |

新产品先确定合法程序集/命名空间、显示名、host/whitelist_host、ProgID、GUID、版本及独立发布者，再同步 project.json、配置、手写入口/安装器/日志、窗口登记和探针断言。新产品 legacy_name 留空，存储/挂载/安装目录按程序集隔离；网站产品映射按产品身份区分，不能因宿主同为Excel而套用本产品版本。

ThisAddIn使用 `[StartupObject(0)]`、FullTrust，继承 `Microsoft.Office.Tools.AddInBase`，接收对应宿主ApplicationFactory与IServiceProvider，重写CreateRibbonExtensibilityObject。主DLL不做强名称签名，应用/部署清单另签。AddinsWL登记“值名=addin_key、数据=空字符串”，只登记对应宿主，不建立同名子键。

### (3). 用最小功能建立闭环

按“环境可执行→身份/必要文件齐全→一个信息按钮与只读信息窗→纯模型/回调检查→编译签名打包→安装与宿主点击”推进，随后增加菜单、Agent和业务。Engine负责运算/提交，Form只收集显示；先测正常、空态及异常，再增加下一层功能。

每次新增源码/探针同步 project.json 的 sources/tools、退出清理及产品断言。生成文件由配置/构建产生，不手改 `.g.cs`。发现旧命名空间、缺窗口映射、未知目录或错网站产品时按身份表修正，不以旧产品PASS证明新产品已受检。

首次全新产品发布者由 src/tools/MakeCertCapi.cs独立编译生成，所需引用取project_support.CSC/COMMON_REF和write_project_info；先确认唯一配置可导入且固定指纹为空，再显式运行工具，核对CurrentUser/My中的主题、有效期和HasPrivateKey，回填准确指纹并保留私钥。失败先查是否已创建，避免重复建证书。已有产品只导入原证书及私钥，绝不静默换发布者；公开.cer不能用于签名。

### (4). 公共源码单源与产品上下文

公共实现只在C-00-Shared维护，产品project.json通过@shared/引用。source_layout.resolve_source限制访问产品目录或公共目录，拒绝缺失及越界引用；共享Python通过产品薄入口显式取得PROJECT_ROOT，不能按共享文件位置推导产品。产品身份、生成文件、库依赖和输出保持产品隔离；不新增运行时公共DLL。源码探针读取与编译相同的物理实现；缓存和构建期指纹包含公共输入。SourceLayoutProbe用摘要注入验证公共改动使两产品缓存同时失效，不在构建中修改真实源文件。

### (5). 功能目录与代码索引

两个加载项的功能实现位于 AC-Original Code，层级为选项卡→分组→菜单→末级功能。末级 Connect 分部文件保存真实按钮回调，算法和弹窗按专属归叶子、共用归就近 _Shared 的原则维护；跨产品业务仍只在 C-00-Shared。加载框架、生成文件、配置器、安装器及校验保留原工程目录。

完整导航仅维护在 [B4-Code Index](<B4-Code Index.md>)。新增功能时同步配置器、project.json 的 sources/features，再运行 python C-00-Shared/feature_index.py --write；默认不带 --write 为只读一致性检查。code_layout 按稳定菜单节点 ID 登记英文目录名，features 按叶子 ID 记录目录、入口、实现依赖与 tests；门禁按真实菜单树推导路径，显示名称和同层排序变化不触发目录重命名。

编译清单使用物理相对路径。历史源码守卫通过 source_aliases 显式连接拆分前文件名与实际编译片段；Python 读聚合文本用 source_text，读取单一物理文件用 source_path；C# SourcePaths 的目录扫描枚举实际编译文件，读取拆分别名组合对应片段。缺失、歧义或未参与编译的依赖必须失败，不因旧路径消失而跳过。缓存与清理盘点将 AC-Original Code 视为必要源码输入。 R-30 另核对源码完整归属、基础设施白名单、父菜单接线和共用代码最近公共父级；单项目与联合入口在完整、增量模式均执行只读核验，隔离正反例通过 --self-test 检查。

## 2. 功能区按钮、图标与菜单

### (1). 配置与接线

普通按钮使用全局唯一稳定ID、专属public回调和 `GetImage_<ID>`；网址叶子由生成器建立 `OnOpenUrl_<ID>`。不要按显示文案或WPS代理control.Id分派。配置器检查Connect.cs及Connect.*.cs；业务源文件必须在sources内。

最小配置如下，保留产品自身的版本、证书和生成器导入/bind/main部分；两个回调在Connect分部类实现，图片包装调用本产品现有图片入口：

```python
BUTTONS = [{
    "id": "btnInfo", "group": ("grpInfo", "信息"),
    "label": "项目信息", "size": "normal", "builtin_icon": "document",
    "action": "feature", "callback": "OnShowInfo"
}]
```

```csharp
public void OnShowInfo(object control)
{ UiDialog.Show("加载项已连接。", "项目信息", MessageBoxButtons.OK, MessageBoxIcon.Information); }
public IPictureDisp GetImage_btnInfo(object control)
{ return GetButtonImageCore("btnInfo"); }
```

父菜单仅展开，使用type=menu与children，没有action/callback/url；group和size只在根层，后代继承分组且不生成size。同层按配置顺序，最多8层；隐藏父项连同后代隐藏，空菜单不生成。root normal优先用原生三行；column仅将同组连续normal根项装入vertical box，不能跨组、断续或用于large/后代。

新增功能完成顺序：定义输入输出及取消/失败边界→Engine/Form→唯一配置→Connect接线→sources/tools/退出清理→配置与XML/回调/资源检查→重建→用户安装点击。展开、悬停、Esc不执行叶子业务。

### (2). 图标与绘制

自定义图标优先，空值或字符串None回退内置；编译时Base64内嵌，运行不读取影子缓存旁文件。根large为32、normal与后代16逻辑像素；large上图下文、normal横排。不用空格、无效Ribbon属性或伪造文字颜色调整布局。

IPicture宽高是HIMETRIC，乘96/2540才得逻辑像素；DPI/OLE首次映射可能缓存。RibbonImages用临时32位图实测，并按目标/实测比例校准，内置/自定义共用，误差最多1逻辑像素、最多重校准一次，再失败记录后回退，不改变宿主DPI。成功转交HBITMAP由OLE管理，失败DeleteObject；不释放宿主已取得或缓存的图片。

32画布主体约24～28并保留安全边距；内置垫白、自定义按配置底色合成24bpp兼容WPS alpha。窗口图标单独生成16/32透明ICO并由窗口释放。悬停文案按中文2/半角1折行、XML转义；结果模板按ID，仍补充真实失败和清理结果。

## 3. 弹窗美工、控件与网格


### (1). 先确定容器，再确定数据控件

本节是常见对象的实用参考，不穷举全部控件。编号1—27以项目实际窗口和源码为依据，28—36为“当前项目未使用”的通用示意；图3—5的局部来自不同窗口。按窗口名称、区域用途、对象及修改范围描述需求，图中轮廓仅用于定位。

**图1：结构、容器与文字。** 关注容器层级及内容区域，特别区分外层日志卡片（3）和内部日志文本框（6）。[可编辑SVG](../C-01-ExcelIterate/AB-Graphic/ui-object-map.svg)。

![跨工作簿表头匹配合并：结构、容器与文字](../C-01-ExcelIterate/AB-Graphic/界面对象-表头合并-编号示意.png)

| 准确称呼（图1编号）     | 本项目中的具体对象                                          | 实现及修改范围                                                                                |
| -------------- | -------------------------------------------------- | -------------------------------------------------------------------------------------- |
| 窗口／弹窗（1）       | 整个带标题栏的“跨工作簿表头匹配合并”窗口，类MergeHeadersForm            | Form；标题、整体大小、打开与关闭行为                                                                   |
| 布局区域／布局容器（2）   | 蓝色虚线范围对应主布局root；顶部字段组targetRow、底部按钮栏actions各有自己的布局 | TableLayoutPanel／FlowLayoutPanel；排列、间距、布局行列及区域尺寸；布局容器通常没有可见外框                          |
| 卡片／卡片容器（3）     | 底部包住日志的白色圆角框logCard；顶部目标区和三张网格也各有卡片                | UiCard；背景、圆角、边框、内边距，无阴影；实际高度需核对父布局                                                     |
| 标签／说明文字（4）     | 顶部“目标工作簿”等字段名，以及来源网格上方的操作说明hint                    | Label；文案、字号、颜色和换行                                                                      |
| 输入框／文本框（5）     | 顶部目标路径targetPath及“表头行”targetHeaderRow；后者没有数值调节按钮   | TextBox；输入内容、校验规则、宽高及可编辑状态                                                             |
| 日志文本框／日志文本区（6） | 日志卡片内部的logView，显示“22:41:45 当前工作簿不是.xlsx…”等提示       | 只读RichTextBox；文案、字号、换行、滚动、选择复制及只读状态                                                    |
| 数据网格（9）        | 上方来源网格Sources、中部匹配结果网格Map、下方公式样本网格Samples          | DataGridView；数据展示、编辑、选择及列宽、行高                                                          |
| 网格列标题／行头（11）   | 网格顶部“来源工作簿”“工作表”“表头行”等列名栏；本窗三张网格均隐藏左侧行头            | DataGridViewColumnHeaderCell／DataGridViewRowHeaderCell；标题文案、字体及尺寸；数据列中的“表头行”或“序号”不等于行头 |

**图2：选择、按钮、拖动与滚动。** 沿用图1的窗口布局；网格滚动部件另作局部放大，使用构造样例解释当前网格支持的双向滚动。数值输入框（8）来自“自定排序”另窗例子；空来源网格的下拉列（10）只标列范围。[可编辑SVG](../C-01-ExcelIterate/AB-Graphic/ui-interaction-map.svg)。

![跨工作簿表头匹配合并：选择、按钮、拖动与滚动](../C-01-ExcelIterate/AB-Graphic/界面对象-表头合并-交互操作示意.png)

| 准确称呼（图2编号）      | 本项目中的具体对象                                                              | 实现及修改范围                                                    |
| --------------- | ---------------------------------------------------------------------- | ---------------------------------------------------------- |
| 独立下拉框（7）        | 顶部目标“工作表”targetSheet，Control.Name为targetSheetBox                       | ComboBox；选项、当前值、可否输入及宽度                                    |
| 数值输入框（8，另窗例子）   | “自定排序”的“同一菜单中的顺序号”number，Name为orderNumber；“自定格式”的字号size，Name为fontSize  | NumericUpDown；带数值调节按钮，可改范围、步长、精度及尺寸；合并窗的“表头行”不是此类控件        |
| 网格下拉列／下拉单元格（10） | 来源网格“工作表”列sheet；某条来源记录的row.Cells["sheet"]是下拉单元格。图中为空表，只标出列范围，尚无可见下拉单元格 | DataGridViewComboBoxColumn／Cell；明确整列或指定记录；各来源工作簿可有不同的工作表选项 |
| 按钮／命令项（12）      | 底部“复制日志”按钮CopyLog；功能区命令另写完整菜单路径，如“审计工具箱→表头匹配”                        | 窗内RoundedButton或Ribbon命令；文字、图标、状态、尺寸及点击动作                  |
| 宽度拖柄（13）        | 顶部目标路径输入框右侧的细竖柄；部分按钮旁也有同类竖柄                                            | WidthGrip；鼠标左右拖动，调整对应输入框或按钮宽度；竖柄的外观方向不等于拖动方向               |
| 高度拖柄（14）        | 日志卡片上方gripSamples、下方gripReport两条细横条                                    | RowGrip；鼠标上下拖动，调整父布局行高；上方同时调整公式样本区与日志区，下方只调整日志区、固定按钮栏      |
| 垂直／纵向滚动条（15）    | 图2网格放大示例右侧的竖向滚动条；日志文本框也用同类纵向滚动条浏览多行文字                                  | 改变内容的垂直可见位置，不改变卡片或文本框尺寸                                    |
| 水平／横向滚动条（16）    | 图2网格放大示例底部的横向滚动条；当前合并窗三张网格支持双向滚动，日志未启用横向滚动                             | 改变内容的水平可见位置；若为日志新增横向滚动，须说明自动换行策略                           |
| 滚动滑块（17）        | 图2放大示例中，滚动条上可拖动的小块；水平与垂直滚动条各有滑块                                        | 拖动滑块查看对应方向的内容；描述时指明所属滚动条，可提出尺寸、颜色、可见性或拖动反馈需求，并核对原生控件支持范围   |
| 滚动轨道（18）        | 图2放大示例中，供滑块移动的区域，包括滑块上下或左右两侧的空余部分                                      | 描述轨道颜色或点击轨道的滚动行为；须核对原生控件是否支持所需定制                           |

**图3：选项与网格操作。** [可编辑SVG](../C-01-ExcelIterate/AB-Graphic/ui-options-map.svg)。

![图3：选项与网格操作](../C-01-ExcelIterate/AB-Graphic/界面对象-选项与网格操作.png)

| 准确称呼（对应图号）      | 实例或使用状态                               | 实现、修改范围及源码入口                                                                                                                                                                                        |
| --------------- | ------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| 独立复选框（19）       | “自定格式”的粗体／斜体；“文件批量整理”的替换、分类等选项        | CheckBox；勾选状态、独立多选、三态和事件；[DialogFormatting.cs](../C-00-Shared/src/addin/DialogFormatting.cs)、[FileToolForms.cs](<../C-01-ExcelIterate/AC-Original Code/ExcelIterate/Operations/FileOperations/BatchProcessing/_Shared/FileToolForms.cs>)                              |
| 网格复选列／复选单元格（20） | Excel“按列筛选”的“显示”列；明确整列或某条记录           | DataGridViewCheckBoxColumn／Cell；复选值、提交时机、只读状态；[ColumnFilterForm.cs](<../C-01-ExcelIterate/AC-Original Code/ExcelIterate/Data/FilterEnhanced/FilterColumns/ColumnFilterForm.cs>)                                                                             |
| 单选按钮／单选按钮组（21）  | Excel“链接助手”的当前工作簿、其他工作簿、其他文件夹         | RadioButton；同组互斥、默认选项、分组容器及切换事件；[LinkAssistant.cs](<../C-01-ExcelIterate/AC-Original Code/ExcelIterate/Operations/LinkAssistant/LinkAssistant.cs>)                                                                                                 |
| 网格按钮列／按钮单元格（27） | Excel“文件提取名称”的“候选”列；共享“交互弹窗”的“可视化调整”列 | DataGridViewButtonColumn／Cell；单元格内的命令，不等于窗外按钮；[Connect.FileNames.cs](<../C-01-ExcelIterate/AC-Original Code/ExcelIterate/Operations/FileOperations/BatchProcessing/FileNames/Connect.FileNames.cs>)、[DialogSizeEditor.cs](../C-00-Shared/src/addin/DialogSizeEditor.cs) |

**图4：树与分区。** [可编辑SVG](../C-01-ExcelIterate/AB-Graphic/ui-tree-map.svg)。

![图4：树与分区](../C-01-ExcelIterate/AB-Graphic/界面对象-树与分区.png)

| 准确称呼（对应图号）   | 实例或使用状态                        | 实现、修改范围及源码入口                                                                                                                                                                                               |
| ------------ | ------------------------------ | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| 树形控件／树节点（22） | “自定格式”的控件树、“自定名称／自定排序”的命令树     | TreeView／TreeNode；层级、展开折叠、选中、节点身份和同父排序；六点排序柄是节点子部件；[DialogFormatting.cs](../C-00-Shared/src/addin/DialogFormatting.cs)、[AgentCustomization.cs](../C-00-Shared/src/addin/AgentCustomization.cs) |
| 分区容器／分隔条（23） | “自定格式”中formatSplit的左侧控件区和右侧预览区 | SplitContainer及其分隔条；SplitterDistance、分区最小宽度、拖动后布局；与13号独立宽度拖柄分别命名；[DialogFormatting.cs](../C-00-Shared/src/addin/DialogFormatting.cs)                                                                 |

**图5：反馈与临时菜单。** [可编辑SVG](../C-01-ExcelIterate/AB-Graphic/ui-feedback-map.svg)。

![图5：反馈与临时菜单](../C-01-ExcelIterate/AB-Graphic/界面对象-反馈与临时菜单.png)

| 准确称呼（对应图号）        | 实例或使用状态                      | 实现、修改范围及源码入口                                                                                                                                                                                                           |
| ----------------- | ---------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| 进度条（24）           | Excel“追踪从属”的扫描进度；“追踪引用”未创建此条 | UiProgressBar／ProgressBar；进度范围、当前值、连续或不确定状态，不用于滚动内容；[TrackerForm.cs](<../C-01-ExcelIterate/AC-Original Code/ExcelIterate/Data/_Shared/TrackerForm.cs>)                                                                                                    |
| 上下文菜单／菜单项／分隔线（25） | 追踪网格的右键复制菜单，及OCR候选单元格打开的菜单   | ContextMenuStrip／ToolStripMenuItem／ToolStripSeparator；出现位置、命令、启用状态、分组与关闭释放；[TrackerForm.cs](<../C-01-ExcelIterate/AC-Original Code/ExcelIterate/Data/_Shared/TrackerForm.cs>)、[Connect.FileNames.cs](<../C-01-ExcelIterate/AC-Original Code/ExcelIterate/Operations/FileOperations/BatchProcessing/FileNames/Connect.FileNames.cs>) |
| 悬浮提示（26）          | 合并窗按钮的说明；Word“查看链接”的摘要全文提示   | ToolTip或网格单元格提示；提示文案、关联控件、显示时机及释放；[MergeHeadersForm.cs](<../C-01-ExcelIterate/AC-Original Code/ExcelIterate/Data/AuditToolbox/MergeHeaders/MergeHeadersForm.cs>)、[LinkViewForm.cs](<../C-02-WordUpdater/AC-Original Code/WordUpdater/Links/ViewLinks/LinkViewForm.cs>)                                           |

**图6：通用选择与输入参考。** [可编辑SVG](../C-01-ExcelIterate/AB-Graphic/ui-common-input-map.svg)。

![图6：通用选择与输入参考](../C-01-ExcelIterate/AB-Graphic/界面对象-通用选择与输入.png)

| 准确称呼（对应图号）   | 实例或使用状态         | 实现、修改范围及源码入口                                                |
| ------------ | --------------- | ----------------------------------------------------------- |
| 选项卡／选项页（28）  | 当前项目未使用；图6为通用示意 | TabControl／TabPage；页签、选中页、切换与页面内容；不同于WPS功能区选项卡或网格中的“选项卡归属”列 |
| 分组框（29）      | 当前项目未使用；图6为通用示意 | GroupBox；带标题的分组边框及内部控件；不同于3号UiCard                          |
| 列表框（30）      | 当前项目未使用；图6为通用示意 | ListBox；选项集合、选中项、多选和滚动；不同于数据网格                              |
| 勾选列表框（31）    | 当前项目未使用；图6为通用示意 | CheckedListBox；列表项的勾选与选中分别说明；不同于20号网格复选列                    |
| 日期／时间选择器（33） | 当前项目未使用；图6为通用示意 | DateTimePicker；日期／时间值、范围、显示格式及日历选择；不同于普通文本框                 |

**图7：通用列表与展示参考。** [可编辑SVG](../C-01-ExcelIterate/AB-Graphic/ui-common-display-map.svg)。

![图7：通用列表与展示参考](../C-01-ExcelIterate/AB-Graphic/界面对象-通用列表与展示.png)

| 准确称呼（对应图号） | 实例或使用状态         | 实现、修改范围及源码入口                                       |
| ---------- | --------------- | -------------------------------------------------- |
| 列表视图（32）   | 当前项目未使用；图7为通用示意 | ListView；列表／图标／详细信息视图、项与子项、选择和排序；不同于9号DataGridView |
| 数值滑动条（34）  | 当前项目未使用；图7为通用示意 | TrackBar；范围、刻度、步长及数值；其滑块调整数值，17号滚动滑块改变可见内容位置       |
| 图片框（35）    | 当前项目未使用；图7为通用示意 | PictureBox；图像来源、缩放、居中及留白；更换图像时处理资源释放               |
| 链接标签（36）   | 当前项目未使用；图7为通用示意 | LinkLabel；可点击的文字链接、目标、访问状态和打开动作；不同于4号普通Label       |

5号文本框的变体也要按用途命名：**密码输入框**使用TextBox的UseSystemPasswordChar／PasswordChar掩码显示（例如文件提取名称的secretKey），**多行编辑框**使用Multiline并说明是否允许换行、Tab和编辑（例如需求记录的notesBody）；只读日志与可编辑需求记录分别描述。系统文件夹／文件／颜色选择器属于1号窗口的不同类型，说明其用途与返回结果，不把它们当普通输入框。

**滚动与尺寸调整要分别描述。** 滚动条用于查看更多内容，13／14号拖柄用于调整控件或区域尺寸。合并窗日志logView当前为只读RichTextBox，WordWrap=true（自动换行）、ScrollBars=Vertical（仅启用纵向滚动）；长行通过换行显示，未启用横向滚动。内容未溢出时可能看不到滚动条；图2的滚动部件用独立网格构造样例展示，未给日志增加横向滚动条。若要求左右查看完整长行，应明确提出关闭自动换行、启用双向滚动，这是新增需求，不能描述为修复已有横向滚动条。该窗口三张数据网格已配置ScrollBars.Both，网格的滚动能力不能当作日志的既有能力。

两图中的合并窗对象集中在[MergeHeadersForm.cs](<../C-01-ExcelIterate/AC-Original Code/ExcelIterate/Data/AuditToolbox/MergeHeaders/MergeHeadersForm.cs>)：构造函数组装布局和卡片，BuildSourcesGrid配置来源列，AppendLog添加时间戳。卡片与控件外观见[UiKit.cs](../C-00-Shared/src/addin/UiKit.cs)；另窗数值控件分别见[AgentCustomization.cs](../C-00-Shared/src/addin/AgentCustomization.cs)的CommandCustomizationForm和[DialogFormatting.cs](../C-00-Shared/src/addin/DialogFormatting.cs)的DialogFormatEditor。这些源码入口供两项目共享参考，不表示Word加载项具有Excel合并业务。

典型关系为“窗口→布局容器→卡片→输入控件／日志文本框／数据网格”；仅数据网格继续区分数据行、列和单元格。布局容器也可直接放控件。TableLayoutPanel的布局行列用于排列，DataGridView的行列用于展示业务数据；界面网格单元格与宿主工作表单元格、Word文档表格单元格分别命名。“来源网格中的‘表头行’列”是输入工作表行号的数据列，“来源网格的列标题”才是顶部列名栏。涉及单条记录时用文件名、工作表名等业务身份定位，不能只靠排序后的可视行号。

宽度拖柄（图2·13）与高度拖柄（图2·14）的定位、拖动方向和作用范围见表；其他拖动对象按用途命名：**格式预览缩放手柄**位于“自定格式”右侧预览的选中控件右边、下边和右下角，调整尺寸草稿；**命令树排序拖柄**是“自定排序”树节点的六点标记，拖动节点调整同父命令顺序。这两种预览／排序手柄未出现在合并窗截图中。需求应写明位置、调整目标和是否保存重开后的结果；机制和源码入口见[四种拖动及配置生效](B2-DevExperience.md#5-四种拖动及配置生效)。

需求句式为：**窗口名称／菜单路径→区域位置和用途→对象及作用范围→要改的属性或动作→期望效果→生效与保存方式（需要时）**。例如：

- 外层容器：“修改‘跨工作簿表头匹配合并’弹窗底部日志卡片（图1·3）的边框颜色和内边距。”
- 内部文字：“修改同一弹窗底部日志文本框（图1·6）中‘当前工作簿不是.xlsx’这条提示的文案和字号。”时间戳格式与业务提示文案分别说明。
- 网格控件：“修改同一弹窗来源网格‘工作表’下拉列（图2·10）的选项，作用范围为所有来源记录。”若只改一条，补充对应文件名和工作表名。

**完整需求表达示例（拟新增双向滚动的需求，不代表当前已实现）：**

> 我需要你帮我修改ExcelIterate中“审计批量匹配→表头匹配增加”打开的“跨工作簿表头匹配合并”窗口。修改对象是底部日志卡片内部的日志文本框（图1编号6，源码字段logView）。
>
> 请关闭自动换行，保留日志已有的换行符，让较长的一行保持原来的行结构。文字宽度超过可视区域时，显示底部水平滚动条；日志高度超过可视区域时，显示右侧垂直滚动条。用户可以拖动对应的滚动滑块，左右查看完整长行、上下查看更多日志。
>
> 文本框保持只读，继续支持选择、复制文字及“复制日志”按钮。完成后，请分别用一条包含长路径的超长日志、多行日志和少量短提示验证：长行能左右查看，多行能上下查看，未溢出时不显示多余滚动条；并确认通过上下高度拖柄调整日志区高度后，滚动条仍正常工作。

底部日志卡片logCard是局部变量，内部logView是字段。外层处理边框、圆角和内边距；日志区高度由父TableLayoutPanel对应行及上下RowGrip调整；内层处理文字显示、换行、滚动和复制。源码变量名用于定位，不直接等同Control.Name或稳定配置ID，配置身份须核对窗口ID和注册规则。输入控件提供参数，网格展示数据与编辑状态，Engine负责业务校验和写入。

### (2). 统一视觉与布局

UiKit直接编译进加载项/安装器，无第三方UI运行库。背景#F5F7FB、白卡、边框#E4E9F2、主蓝#2563EB、从属绿#16876F及危险红；卡片圆角12、按钮8，外边距16、区间12。正文Microsoft YaHei UI 10pt、标题16pt、辅助9pt，缺字体回退。

RoundedButton不透明父色覆盖整区，禁系统默认按钮绘制，保留正常/悬停/按下/禁用/键盘焦点状态，避免黑色虚影。Ui.PrepareForm接UiWindowChrome，保留原生标题栏；Windows11支持时DWM配色，高对比度优先系统，失败不阻止开窗。标题栏字体由系统管理。

同一控件不混用手工坐标和Dock控制；短标签按真实字体测宽，长说明按实际可用宽度测高，加粗后重测。状态、说明与动作区独立，窄窗换行或滚动，主区域可选择复制；长公式放列表单元格而非无空格Label。固定坐标窗自行Arrange并按真实Bottom推进后续区域，不能布局后强盖尺寸。

抵消窗口初始状态由DialogSizing在Load内、可见之前应用；Shown只定位分隔条，不重复修改ClientSize或Location。尺寸和格式预览均保持普通窗口；不通过改变TopMost或工具窗样式处理绘制问题。根布局、ResultGrid、WidthGrip及RowGrip局部缓冲并完整铺底；Opaque控件的OnPaint必须覆盖背景。来源和目标字段组按实际可用宽度换行，保留各输入框拖柄和配置身份。默认布局按窗口、DPI和字体条件延迟测量，尺寸编辑器显式请求完整目录。离线静态渲染只能证明该控件树能绘制，短暂黑影与真实跨屏仍须宿主验收。

### (3). 列表填充与身份

搜索只改可见行，不改勾选；全选只作用当前结果。行Tag绑定原对象/列身份，排序后提交、恢复仍按身份，特殊自身行不混入业务计数。Rows.Add首行可能在Tag赋值前触发选中事件，填完显式同步明细/按钮；右键先选鼠标所指行。

行高默认约34、表头约38，关闭EnableHeadersVisualStyles；添加行前设RowTemplate，句柄创建及格式应用后复核DPI。表头完整单行、居中且默认加粗，显式自定格式可关闭粗体；实际字体、字号、Padding和排序图标决定最小列宽与单行表头高度，不启用表头换行。短列AllCells、长列Tag存逻辑上限、末列Fill，均须先满足完整表头宽度；手动缩窄不低于该下限，窄窗横滚。文件提取名称另按正文30个“0”限宽，表头所需宽度优先，不能被局部策略或后续格式应用覆盖。DataGridViewColumn没有MaximumWidth。特殊行同时设普通/选中背景；全文提示用未受EndEllipsis限制的真实测量。规范与维护范围见\[格式维护规范\](B3-Gate Control.md#5-格式维护规范)。

保留原生Ctrl+C，扩展快捷键采用Ctrl+Shift。Clipboard.SetDataObject(text,true)支持宿主退出后粘贴，保留业务已有占用重试。界面或日志显示失败不改变已提交业务事实；诊断与用户结果分流。

结果填充先批量准备行；抵消结果达到5000行时用矩阵作为稳定数据源，CellValueNeeded按需提供所有单元格值，原生复制仍读取完整值。列宽继续按全部数据和完整表头计算，数据、字体、表头或样式变化使缓存失效；同批无变化时跳过重复测宽并保留手调列宽。不能用只测可见行代替完整宽度要求。

## 4. 宿主线程、窗口寿命与安全预览

### (1). COM与窗口分工

宿主UI STA捕获稳定目标/快照并提交COM操作，窗口线程只收纯数据；临时COM引用用finally逐次ReleaseComObject，GetIUnknownForObject增加的引用逐次Marshal.Release，宿主传入的Application为借用引用不释放；工具自己GetActiveObject取得的引用逐次Release，也不将FinalReleaseComObject推广到共享RCW。对象持有期间按RCW/IUnknown加路径名称核对，关闭重开的同名对象不能冒充原目标。

信息/浏览/结果窗通常非模态Show，并按有效宿主同进程HWND指定Owner；同类窗单实例，重复命令刷新/激活，FormClosed只清对应实例。纯参数窗使用独立STA、Application.Run和非模态Show；宿主登记会话、明确应用后复查原目标，等待期间不占全局忙碌。卸载在所属线程关闭自有窗/字体并清引用，重复或迟到回调不写入。

Owner候选需HostDialogService.ValidateOwner核对同进程、调用线程、可见顶层窗，不能借其他应用前台句柄。保留各类窗口的合理原生样式；某个样式差异只能作为排查线索，不能断言ShowInTaskbar/MinimizeBox设置必然造成菜单故障。

同一宿主的从属扫描与抵消预检由HostBatchRunner驱动：宿主STA每批读取最多4096行及65536个单元格，待解析队列最多两批，后台只接收数据块和复制请求；进度合并通知。SourceChangeWatch通过应用事件监听相关来源修改、重算、关闭及删表，块间及提交前复查原对象；关闭重开同名工作簿不能冒充原对象。监听不可用时回退原线程同步读取、后台计算并记录。纯运算取消检查不改变公开引擎签名；关窗、卸载取消会话并拒绝迟到回调。

### (2). 同步确认隔离

UiDialog.Show保留DialogResult/默认值与样式；宿主提示、确认经HostDialogService独立STA显示，以CoWaitForMultipleHandles原生COM感知等待，宿主不建WinForms模态循环，不调用ShowDialog/DoEvents。重入确认返回取消；只恢复本次确实禁用的Owner，预禁用保持。文件夹/OpenFiles/SaveWorkbook只传路径字符串；颜色选择留独立格式线程，安装器独立进程另处理。

取消/Esc/关闭零写入，破坏性确认不绑Enter；中断全部链接保留两次确认。若应用和取消均导致Ribbon立即收起/无法键入，查消息循环、Owner启用、焦点/捕获与释放，保留首异常证据，不直接认定算法或窗口样式是唯一根因。

### (3). 尺寸登记与真实安全预览

新增自有Form接入五处：DialogCatalog稳定ID/默认与最小96 DPI客户区→DialogSizing.Prepare→DialogPreviewFactory真实Form+固定样例→DialogTree实际Ribbon叶子映射→ThisAddIn退出清理。系统选择器、浏览器、安装器不登记。构造/Load不能自动读真实宿主或执行业务，预览注入样例并禁业务/跳转/高亮清理。

DialogSizes.xml按程序集隔离，禁止DTD、版本/范围校验、锁内摘要冲突检查和原子替换；无效值默认、未来版本只读。尺寸草稿、恢复、预览拖动关闭都需显式保存；ResizeBegin/End只捕获用户拖动，普通业务临时窗口缩放、屏幕裁剪、程序缩放不持久化。新预览沿用当前草稿。目录命令列按实际文字/缩进测宽而非Fill，点击/绘制共用坐标，窄窗横滚。

格式编辑器独立STA，ThreadLocal集合管理窗口/字体；隐藏源Form只初始化一次，停在隐藏父容器，画布仅显示不透明选择Overlay。刷新执行Apply→布局→捕获→矩形同步，不Show/Focus/反复挂接。DrawToBitmap只取源，RichTextBox用EM_FORMATRANGE补绘并释放资源，隐藏TextBox/ComboBox按真实文字/样式补绘、密码遮罩；命中与补绘共用裁剪后的同次快照。不能用空白捕获宣称用户可见。

### (4). PDF 当前文档与后台进程

PDF 功能共享 Connect.Pdf、PdfToolsForm、PdfCurrentDocument 及 src/pdf 下固定版本引擎，使用 WPS_WORD 条件编译区分宿主。Excel 的“文件操作”分“批量处理”和“PDF 工具”；Word 只增加 PDF 工具。菜单仅显示 btnCurrentPdf（转为 PDF）和 btnWatermarkPdf（编辑 PDF），旧五项回调保留并映射到两个窗口的对应页签。file.pdf.current/convert/watermark/split/merge 的旧尺寸与预览键保留：新开窗口使用入口的旧尺寸键，已开窗口只切页签，不重设大小。

当前文档首次进入时在宿主 STA 捕获文档及 Excel 工作表/选区，核对 ActiveWindow 所属进程；保存与导出仍在原线程执行，后台只接收临时 PDF 路径。点击执行时先选择输出目录，再验证原文档身份、路径、只读状态、活动目标及选区有效性，再由宿主 Save。首次保存使用宿主原生保存界面；Saved 为假、路径未落盘、取消或异常均停止。成功后 Excel 按整簿/原工作表/原连续选区导出；选区导出忽略已有打印区域以限定选区，不改原表打印设置。Word 先整篇导出，再按页码范围提取为单个结果；不支持选中文字导出。窗口关闭释放自取 COM 引用，不释放借用 Application。

普通批量只枚举来源对应的 Word 或 Excel 宿主，避免无关宿主阻塞；RPC_E_CALL_REJECTED 或 RPC_E_SERVERCALL_RETRYLATER 转为可读暂停结果，保留失败关闭策略，不自动关闭用户窗口。

批量导出在 COM 创建后核对 HWND 对应 PID，已知为已有进程或已存在文档时停止，不在该宿主打开来源、修改宏设置或导出。WPS 文字可能在首篇文档打开前无法取得 Application.Hwnd，PID 为零不等于已确认复用；只在 Documents 为空时允许打开本批次来源，仅持有该文档、不取得进程退出权。打开后再核对文档窗口 PID，发现已有进程立即停止；无法取 HWND 的异常写入日志。清理只关闭本批次文档，具有退出权时也须再确认无剩余文档才 Quit。PID 新增不单独证明独占归属。PdfWorkerHostProbe 离线覆盖宿主类型、PID、进程归属与文档数的组合，不启动 Office。每次任务在临时目录 ExcelIterate-PDF-* 保留 request.xml、桥接启动日志、worker 逐步日志及结果，记录占用预检、创建宿主、打开、导出与清理阶段；日志可能含本地文件路径。

选目录后整个 WPS 无提示退出：Windows Application 的 .NET Runtime 1026 事件确认 Scan 的 BackgroundWorker 完成回调在线程池中执行，RebuildExtensions 向 UI 容器添加控件触发未处理 ArgumentException。原生 WPS 消息循环可能没有 WindowsFormsSynchronizationContext，不能假设 RunWorkerCompleted 自动回到窗体线程。Scan 和转换启动前在窗体线程建立句柄；两种完成回调均显式 BeginInvoke 回窗体线程，窗体销毁后的回调丢弃。PdfAsyncUiProbe 清空同步上下文，真实扫描自建文件并核对扩展控件与完成回调线程，另测销毁后的迟到回调；不启动 WPS，真实宿主复测仍单列记录。

普通批量在任何输出前预检 Office 来源占用；已打开/被占用时返回暂停并要求保存关闭后重试。混合转换后合并在自建临时目录转换全部输入，全部成功才发布一个 PDF；失败或取消清理中间产物，不输出缺页结果。每次执行都选择输出目录，同名自动编号。文字水印支持 RGB 颜色和居中/平铺。PDF 请求及结果 XML 禁止 DTD，不上传文件。

PdfFileSelection 将完整来源与筛选任务分离；后缀匹配忽略大小写，过滤不删除来源、不改变清单顺序。目录扫描在后台运行，默认第一层，取消不提交半成品；目录联接跳过，无权路径进入日志。重叠输入按完整路径去重，首次发现的来源根归属保持稳定；同名根目录分配不冲突别名。手动移除使用墓碑，重扫不自动恢复，显式重新添加文件可恢复。输出目录子树与本窗口记录的已生成结果不再入队；扫描和处理期间锁定页签及输入。

请求 XML 每个 File 可带 folder 相对目录；后台拒绝绝对目录、点段和越界，逐项创建结果子目录。输出扫描排除仅适用于来源内部的输出子目录；输出在来源上层时，不能排除整个来源树，PdfSelectionProbe 覆盖该边界。合并忽略目录结构，仅原子发布完整结果；集中输出沿用无覆盖编号发布。两个窗口使用独立偏好 XML（宿主命名空间隔离），只保存后缀、递归和保留结构选项，不持久化来源与结果。偏好临时文件原子替换；尺寸预览不保存偏好。

清理临时文件及自建 Office 会话时逐项记录警告，不覆盖原始处理异常或已成功发布结果；失败信息保留操作阶段、来源、输出路径、异常类型与 HResult。仅在逐项处理失败后检测输出目录可写性，不提权、不改权限、不换目录重试。宏设置恢复失败属于清理警告，不能据此声称设置已恢复。桥接器在受限环境无法写入 LocalAppData 缓存时，原样报错；离线工作流可直接运行已编译后台进程，不改变正式缓存位置。

离线入口 `python src/pdf/verify_workflow.py --render` 先构建诊断 DLL，运行 PdfSelectionProbe、生产页签及 800/960 宽预览、真实后台 PDF 请求。它明确跳过桥接器缓存提取；真实桥接器仍由 PdfToolsProbe 默认路径验证。新增选择探针已登记 project.json 和后台验证名单。诊断 DLL 不替代正式 VSTO 产物，后台通过不代表真实 WPS 菜单、焦点、分页、COM 导出或签名验收。

src/pdf/build_trial.py 仅生成独立诊断验证包：使用编译接口夹具、排除 ThisAddIn，启动器附加已打开的对应宿主以验证当前文档功能。它不是 VSTO 安装器，也不能证明真实加载项菜单、安装签名和焦点验收通过。正式 build.py 继续要求真实 Microsoft 引用和原固定证书，不替换发布身份。测试入口为 PdfToolsProbe 与 PdfCurrentDocumentProbe；后者 --live 先核对自取 Application 的 HWND/PID 不属于原有进程，且文档集合为空，才在该独立实例创建测试文档；归属不明则跳过，不操作用户文档。

清单签名校验兼容性：不同 .NET Framework 维护版本的 SignedCmiManifest2 私有校验方法可能只有命名空间参数，也可能增加哈希算法枚举参数。ValidateXmlSig 按实际方法签名调用，仍检查身份、许可证摘要及两段 SHA-256 签名；无兼容签名时失败退出，不跳过校验。本机测试副本用独立安装标识和测试证书构建，不修改本源码的原发布身份或固定证书。

已知问题 PDF-001（待复现、未修复）：用户反馈 XLSX 行列较多或文本较长时，导出可能出现跨页拆分、换行或内容显示不完整。尚无最小复现文件，不能认定存在实际内容丢失。当前沿用宿主打印/分页设置，未实现自动适配一页宽、打印预览或重复表头设置；后续需结合具体文件和宿主版本验证。

## 5. 四种拖动及配置生效

### (1). 输入与按钮拖宽：WidthGrip

Ui.Grip将标签、控件和竖柄放入不换行FlowLayoutPanel；其他布局的floating柄跟随目标Parent/Bounds。MouseDown记录屏幕起点及原宽并Capture，移动按增量调整且限制最小宽；正常MouseUp只提交一次。捕获丢失、控件隐藏/释放取消当次拖动并恢复原宽，不保存。

提交经DialogFormatting.SaveControlWidth将宽度除DPI比例存为96 DPI逻辑像素，使用稳定窗口/控件ID写入同一Customization.xml；预览不落盘，写入失败保留本次可见宽度并记录诊断。下拉框自动测宽先尊重显式宽度，避免Load/FontChanged/选项变化覆盖用户值。Dock拉伸路径框是否加柄由实际布局决定，不写“所有窗口都有柄”。

### (2). 卡片区块改高：RowGrip

横柄调整TableLayoutPanel的上/下内容行，非固定下区时保持两行总高，上区夹在minAbove至total-minBelow、下区取剩余；总高不足两个最小值则不启动Capture或保存。固定按钮栏时只改上区。这是改区域高度，不是拖卡片换位置。起点、Capture、取消和提交纪律与WidthGrip一致；SaveRowHeight以 `窗口ID::row.xxx` 合成键记录上区逻辑高度，具体Form加载时将已存逻辑高度经Ui.Px换算到当前DPI，再检查对应空间门限并按可用空间重排；RowGrip最小高度按96逻辑像素声明，Begin与移动时经Ui.Px换成当前DPI物理高度；若移动时两个最小高度已超过原两行总高，则取消并恢复起始尺寸。真实跨屏DPI验收仍单列，不由UiCard或通用控件尺寸应用器替代。

### (3). 格式预览控件缩放：PreviewSelectionOverlay

先Register稳定Name/确定性ID及ResizeWidth/ResizeHeight能力；右、下、右下手柄分别改宽、高或二者，屏幕增量除DPI比例后更新草稿，不联动字号。MouseUp完成当次草稿，Esc/捕获丢失恢复；整个编辑器保存后才落盘。

FormatRule Width96/Height96为1～4096、0默认；ComboBox/UpDown仅宽可调，标题/表头/数字框内部子控件不另缩放。每次Apply还原原Size/MinimumSize/Dock/Anchor/AutoSize→应用文字字体→Resolve尺寸→重排测量→捕获快照。单行TextBox关闭AutoSize才能改高；超出内容由外层滚动，恢复默认尺寸只清宽高。

### (4). 命令顺序拖动：SortableCommandTree

OwnerDrawAll绘蓝色+/−与六点柄，共用Expander/Grip/TextArea/RowAt坐标，按Level/缩进/横滚/DPI计算；叶子和分支保留相同展开槽，槽→柄→文字各隔4逻辑像素，NoPadding绘文字。不用空格改变层级。

超过系统DragSize才ItemDrag，展开槽不拖；DoDragDrop的目标只限同父菜单子项，顶层不可移动。自定名称树复用布局但不显示六点。保存同父顺序而非拖动Ribbon实体按钮，旧ID保持回调/图标对应。

### (5). 名称、排序、格式、尺寸各有生效时机

RibbonCustomization从默认XML登记稳定命令，根菜单转dynamicMenu；专属GetContent/GetLabel读取最新名称/顺序，不依赖control.Id。保存先核对冲突、原子XML写入，再CustomizationSaving.AfterSave通知宿主dispatcher；宿主STA Invalidate已登记Ribbon，刷新失败不能误报未保存。

名称/顺序当前会话生效；格式、控件宽高及弹窗尺寸只影响下次打开的新窗口。FormatSession.Capture按实例固定快照，Data按摘要读取最近有效配置；已开窗不受后续保存影响。LiveCustomization监听Changed/Created/Deleted/Renamed并去抖BeginInvoke，仅名称/顺序变化刷新菜单。损坏读取保留最近有效数据，未来格式只读；两产品隔离，不重启WPS。

默认Font在登记时独立克隆，由所属窗释放；字号6～36、缺字体/样式回退，表头变化重测。取消零落盘，恢复不丢其余配置；名称/顺序保存不能覆盖格式/尺寸。

## 6. Agent-Evo、内嵌文档与网站

EXTRA_TABS预声明选项卡；主toggle的get_pressed与附加tab的get_visible读同一Visible，onLoad登记Ribbon，点击修改/刷新/原子存AgentExtension.state，失败刷新回旧状态、保存失败如实提示。默认隐藏，显示后由用户选择；新增第二附加tab须独立状态/键/刷新集合，不能只追加配置。卸载清窗与引用。

概览取首个一级章节后的首段（跳过标题与空行，至空行或下一标题止）；内嵌文档不再分章拼接，A-UserGuide 与 B2 各自整份内嵌，只去 BOM、统一换行，并在内嵌前整节剔除「本轮交付证据」，章节范围仍由文档门禁按一级章节标题核对。修改后必须重建才能更新DLL。全文每次经SHOpenWithDialog选择应用，取消保留概览、不写默认关联。AgentLocalFiles按程序集与文档／资产联合摘要释放内容包：包内逻辑路径固定为C-04-Devlog/A-UserGuide.md与C-04-Devlog/B2-DevExperience.md，图片按原文相对链接保留，指向各项目AB-Graphic下的PNG与SVG。WpsShared.Agent.Assets.xml记录当前文档的资源映射、大小及摘要；只导出本次请求的文档和资产，缺失／损坏缓存修复后才打开应用，保留可能在用的旧包。构建、增量缓存、文件盘点和图片门禁复用Website的document_assets收集函数，只把实际引用资源列为必要输入。

开发预览的“全文”每次由AgentDocumentAccess发起口令验证，提示“请输入开发者的微信名称”；掩码TextBox仅在当前窗口中保存输入，产品只存校验摘要，不写输入日志、不持久化授权。错误可重试，取消／Esc／关闭／异常均不导出、不选应用；用户全文、概览和复制不验证。DeveloperPasswordForm通过HostDialogService的结果型隔离接口显示，登记agent.developerpassword及其预览、尺寸和退出清理；预览构造不验证、不导出。该验证保护入口，不加密已导出的MD。正文里的业务源码、姊妹项目与历史证据链接仍是工作区定位，离线内容包只随附文档配图及示意SVG。AgentDocumentCoreProbe为无窗体后台回归入口，不把后台结果登记为真实查看器或WPS验收。

网站资源名WpsShared.Agent.Home.html固定，两DLL内嵌同一无EXE快照，独立HTML附加本轮两包；归一installer-assets后比较网页资源，不比较两个宿主DLL字节。缓存释放到LocalAppData/Xy/AgentEvo/Home/<摘要>/index.html，按http/https关联找默认浏览器，不用.html关联替代。网站配置与下载细节见Website日志。

RequirementsForm使用Requirements.txt UTF-8原子保存并查外部修改；介绍Label独立。普通关闭保存/放弃/取消，宿主退出原子保存唯一unsaved草稿，下一次选择恢复；正式保存成功后清草稿失败不能误报保存失败。预览只用固定样例，不读写用户记录。

# 二. Excel-Iterate 专属开发经验

## 1. 单元格、选区与公式

### (1). 引用/从属追踪与高亮

FormulaRefParser纯字符串解析：去前导等号，运算符拆token，单引号整体保留、双引号文字跳过，冒号/感叹号不当普通分隔符；由具体到一般匹配同表/跨表/跨簿。表名/路径排除式字符类容纳中文/全角/空格，美元符逐行列标绝对引用。TrackerProbe保留引用解析回归。

DependentTracker遍历已开Workbooks/Worksheets的UsedRange.Formula，按真实数组下界与范围起点计算地址；反向比对簿/表/单格或区域。独立STA重连ET，不跨线程复用UI COM，BeginInvoke进度/取消安全步。JumpToCellEx按条目重取目标，跨簿缺失核路径再打开，已有扩展名不重复补；连续跳转不绑定后来活动格。

自身“所选单元格”仅UI插行，Tag仍为PrecedentItem，不计Engine结果。ET黄色Interior.Color=65535，原无填充检查ColorIndex=-4142并恢复哨兵，不还原成白填充；约2.5秒/再次跳转/关闭恢复最早状态。

### (2). 稳定身份与整批事务

WorksheetIdentity共用于公式、数值和列筛选，内容/Hidden/恢复记录分开。持有原表/父簿，以RCW/IUnknown加FullName/Name复查；CodeName是可选VBA字段，曾空值误拒绝，不能强制非空。改名/另存/关闭替换对象、成组选择拒绝；窗口确认仍处理开窗原选区，后来选区不换目标。

Area与UsedRange相交，最多4096格分块，重叠去重；HasFormula区分公式与等号文本，合并格只处理选中左上。数组/溢出/锁定或必要属性读取失败整批停止；Formula2仅明确属性不存在才回退Formula，HasArray/HasSpill须明确false。

统一顺序：原对象/快照→整批预检→二次核对→提交读回→失败逆序还原→恢复EnableEvents/ScreenUpdating。setter写后抛也计尝试；未还原项输出地址/原值/错误并存FormulaRecovery或ColumnRecovery，日志失败可复制，不保证WPS Ctrl+Z。只释放自有Range/集合；约5秒恢复状态，仅恢复自己提示。

### (3). 封装、自定义函数与数值计算

IFERROR默认空文字、ROUND两位、ABS一参，每次加一层；拆解仅完整包住整式的最外函数并保留第一参，SUM适用，SUM(A1,B1)+2/无参不变。FormulaWrapper独立解析引号双写/表名/数组/结构化引用，不能Split逗号或复用丢嵌套的追踪Tokenize。

文字加引号且内引号双写；数值有限/不变区域性，ROUND32位整数可负；表达式须完整单参数，结构校验不代替宿主语义。引用原文插入各格不平移/不在参数窗求值；自定义不改快捷默认。

FormulaInteraction宿主STA捕获选区/登记会话，FormulaWindowThread独立STA仅传会话号、操作、Owner和FormulaRequest；隐藏Control投递提交，重复/迟到不写，等待释放_busy。取消/Esc/关闭/初始化失败零写入，卸载异步清会话与引用；切换表提示回原表，内容/保护变化拒绝。FormulaWindowState只记安全阶段/线程/Owner/程序集摘要，不记公式或输入。

CustomFunctionForm空位选择器常显，一空位置灰、多空位可选，模板变化RefreshSlots；校验提示恒定占位，成功绿色正向提示、失败红色异常，高度不跳。标签按最长实测。自定义模板对公式填公式体、数值/文本填值并转义，空/布尔/错误跳过。

ARITHMETIC_DEFAULTS同源生成默认/文案：加减乘100、除10000；公式整体括号，如=(A1+B1)*100，真实数值变=20*100保留格式。文本数字/布尔/空/VT_ERROR跳过，不能将盒装整数全当数值；直接零除拒绝。IFormulaValueCell存HasFormula/Value2类型内容，提交复查，公式按Formula2/Formula、数值按Value2还原并核原类型。

### (4). 列筛选与快捷高筛

活动格依次定位ListObject、包含它的AutoFilter、CurrentRegion，不用UsedRange猜列；正式表HeaderRowRange，其余首行，范围内空/重复/合并标题逐列留。搜索只改显示，全选当前结果，允许零勾选待选但应用至少一列；Tag存列身份。提交复查对象/范围/表头/Hidden/保护，失败逆序还原并核对。

高筛条件来自同一连续表数据行；首行非空/唯一/未合并，同一行AND、不同OR、重叠去重、全空OR忽略，按选中列补条件表头。临时条件表AdvancedFilter支持比较/通配，新结果同名编号，输出值/数值格式不复制外部公式；失败只清本次未完结果。历史A2:B2/A12:B12逐格参考差异未确认，不写已通过。

## 2. 工作簿、目录与表头匹配合并

### (1). 目录与隐藏表跳转

新目录置前、同名编号，列原表/生成时状态和“跳转至本表”；仅目录区域加粗居中，保留超链接颜色/下划线，不改WorkbookUi.StyleTable影响其他清单。

隐藏表使用目录内锚点+SheetFollowHyperlink，结构保护拒绝；隐蔽名称绑定Range，表名引号双写、改名跟随。事件有限重试订阅/退出解绑，目标删除提示重建；加载项禁用不保证跳转。非常隐藏/含单引号表名的既有真机证据保留。

### (2). 合并模型、预检与公式重写

Connect.MergeHeaders编排；MergeHeadersEngine纯模型/匹配/规划，MergeHeadersHost做COM/快照/备份，MergeHeadersForm显示目标、来源、列映射、样本与报告。只处理.xlsx，目标默认活动已保存簿/表，否则说明并留空；选路径/加入来源即时读表名，最多20来源、50万格，超限分批。

表头Normalize忽略大小写与首尾空格，以名称而非物理位置匹配；缺失列红色列出，用户核对后确认继续，不自动建列。每列各自追加到该目标列实际数据末尾下一行，不保证同来源各列起始行相同；预检后执行再读目标表头核对。

FormulaRefRewriter跳过文字/函数名/科学记数/错误/定义名。当前规则：同表相对列按映射，绝对列保持；相对行按目标起始行减源首行偏移，绝对行保持。跨表/跨簿限定前缀与列保持，但相对行仍偏移，如Sheet2!A1在delta=2时成Sheet2!A3。未映射列保留列字母并报告，相对行仍偏移；越界报告并保留相应原行/列，不概括为整条引用完全不变。

### (3). 合并提交、还原与界面边界

执行确认后保存相关目标、同目录“ExcelIterate合并备份”创建副本，捕获本批写入区快照，提交后逐格读回核对；失败尝试快照还原。取消恢复本批快照，确认仅结束可还原状态、文件备份保留，不能套用文件整理的回收纪律。关闭仍可还原批次时明确提醒；源/目标引用与窗口在卸载释放。

来源网格每行Tag存MergeWorkbookSpec；工作表使用DataGridViewComboBoxCell的每行选项，不能拿表单ComboBox代替。UiCard只是布局容器，RowGrip调整根TableLayout行，WidthGrip调整独立输入/按钮；网格列宽拖动属DataGridView自身机制，不等于按钮排序或自由移动卡片。

MergeHeadersProbe只验证重写、匹配、规划及数组边界；MergeLiveProbe属于单独实机检查。后台结果不证明WPS菜单、键入、焦点/拖动或真实合并全过程；这些需工作区副本验证，并核原文件/用户已有实例不变。

## 3. 外部链接与文件批次

### (1). 链接扫描、映射与保存

ZipArchive读取xlsx关系，清单绑原接收簿；复用固定“外部链接”，只删隐藏名称证明程序归属的旧编号表。范围/行身份持久化，重开复扫前两列集合相同才执行，允许排序、第三列保留；增删/重复/身份变要求重扫。扫描/替换保存本次相关簿含既有编辑，先恢复相关高亮；未命名询xlsx路径，取消/只读/保存失败停止，不设Saved=true绕过。

普通替换的已开簿用ChangeLink不关闭，Name取LinkSources原始名；身份另用完整路径。Windows字面%不解码、URI单独解析，数组真实下界；保存后核实际XLSX关系再唯一匹配原始名。WPS特定版本可能保留旧名称并追加新名，不能以旧名仍在LinkSources判失败；URI仅解UTF-8非ASCII，ASCII空格/#/%转义保留，单盘根用原文件盘符，FileShare.ReadWrite读取，其他版本实测。

链式末端先改/逆序还原；多对一、未改来源占用/环形暂停。不支持则停止，不强制重开/刷新值或批改公式文本。普通替换的关闭簿只改externalLinkPath，其他ZIP部件不变，暂存校验+复核原摘要+File.Replace。MissingTarget用GetAttributes仅FileNotFound/DirectoryNotFound跳过，权限/目录/坏文件仍错；每次处理复查，状态“新链接不存在，暂已跳过，请排查原因”，取消只撤成功项。

C列Trim后仅完整匹配“断开链接”生成LinkAction.Break；旧文本“请断开该链接”明确提示改填，空白跳过、其他文本继续路径校验。不能因没有单元格公式就隐藏来源：外链可能保存在名称定义。LinkNameAnalysis按workbook/externalReferences顺序及关系路径解析数字编号，不把externalLink文件编号当来源编号；区分工作簿/工作表名称作用域，扫描在现有状态列说明名称引用及数量，分析失败仍保留来源并提示待核对。

名称外链的扫描文案统一为“仅名称管理器定义引用（X 个名称）”，X按定义条数、含作用域区分计数。只改展示文本时，同步精确文案断言、实机探针及用户指南；无需新增重复测试。旧清单保存的状态不靠安装自动覆盖，需重新扫描；此类文案调整的构建证据与此前断链逻辑的实机证据分别记录。

含断链的工作簿先对全部所选来源做名称依赖预检，再建立整簿备份。名称必须仅指向该来源且未被公式、其他名称或对象使用；混合来源、同作用域身份不唯一、动态引用、嵌入/宏对象等无法证明安全时停止。执行时重新读取保存包，核对宿主Name的作用域与RefersTo，再删除精确匹配的名称、重取真实来源名并调用BreakLink(…,1)，不直接删ZIP关系或全表转值。用户样本中两个来源共49个名称，实机原生BreakLink后仍保留；清理这些名称后再断链才能移除来源，不能用只含单元格外链的假对象冒充此回归。

WPS保存可把缺失来源标为xlExternalLinkPath/xlPathMissing：它仍占用externalReferences编号并可能被其他名称引用，名称分析须识别其Target，不因其不是普通externalLinkPath而中止整个分析，也不能按扫描显示条数重新编号。其余单元格公式按部件和地址建索引、按实际来源路径核对，避免外链编号变化误判及逐格遍历整个结果造成平方复杂度。名称依赖检测跳过字符串、表格结构化列名及其他工作簿的名称，并考虑工作表局部名称对全局名称的遮蔽；未知依赖保守拒绝。

名称依赖不能只靠正则词边界：Y$5、$Y$5、Y:Y与筛选名称中的$A$8:$Y$271都含独立列字母Y，却不是定义名称Y。用户样本因此产生局部Y的360处假依赖和全局Y的1处假依赖，预检在修改前停止。UsesName先复用FormulaRefRewriter.Scan取得A1地址跨度，再核行列有效边界，只跳过完全落在有效地址内的候选；裸Y、SUM(Y)、Sheet!Y及名称与地址混合公式仍须拒绝。Y0、Y1048577、XFE5等超出有效地址边界的名称不能被忽略。回归须同时保留旧DLL失败、地址通过、真实名称拒绝、筛选名称保留，以及完整会话保存/重开和逐字节取消恢复证据，不能通过取消依赖保护来修复误报。

实机格式回归应核对单元格实际使用的cellXf及其字体、填充、边框、数值格式、基样式和主题，而非直接要求styles.xml所有字节相同。样本在关闭后重开暂存副本时，WPS会重排未使用字体记录；数据、几何和已使用样式不变不应判为格式损坏。名称断链须同时验证打开目标、关闭目标及清单接收簿就是目标的路径，保存/重开核对后再关闭自有样本并验证整簿恢复。宿主名称身份解析仅去掉作用域外围引号、解码成对单引号，不能把全部单引号删除而混淆不同工作表；名称句柄读取后、首次删除前再次核Saved和文件摘要，读取期间出现的新编辑不自动保存或覆盖。

首次名称删除或断链前沿用Intent/Uncertain记录；保存后核指定来源与名称已移除、其余名称及公式映射不变，数字外链编号以路径归一后比较。核对失败不得声称原链接未改变，区分未执行、已保存待核对与已还原并保留恢复记录。来源文件缺失不妨碍按当前值断链。关闭目标用同目录副本在现有App中打开，核FullName归属、UpdateLinks=0，保存验证后原子替换，仅关闭自有副本；副本出现新增编辑则保留、不关闭或提交；不启动或Quit宿主。四列清单、LinkAction和旧批次XML保持兼容，执行前重新分析实际文件，不依赖扫描时的内存统计。

### (2). 结果、恢复与清理

LinkBatch存提交意图/实际成功、用户记录与摘要，取消反向改本批并保存核对，保留其他编辑；未完禁止确认。单簿日志其目录、文件夹一汇总、多目录分别；COM按宿主STA调度、关闭包独立STA，取消等当前调用结束。恢复记录写失败禁止继续提交，清单回写失败不改变业务事实。

LinkOperationLog.Record输出无时间用户结果，Trace保留参数/HRESULT等诊断；窗口按事件而非中文过滤，新扫描清用户记录、按偏移追加诊断、重开不重复回放。成功只在保存/关系或提交摘要核对后报告，统计单位链接，失败核提交意图再反向还原，未改不写已还原；清理停止诊断追加避免重建已回收日志。确认严格回收本批记录，不删原簿。

含断链的Owner创建LinkWorkbookCheckpoint：精确整簿备份、BeforeHash/LastHash、PreparedHash及提交意图在原文件提交前持久化。备份后新增未保存编辑或文件摘要变化即停止，不再次RequireSaved吸收用户新编辑。取消必须关闭相关簿且提交摘要一致，整簿恢复同Owner本批全部操作；普通替换Owner仍逆向换链保留后续编辑。程序报告回写独立更新提交摘要，失败保留Uncertain；未知提交按PreparedHash核对，不猜状态、不关闭用户簿。未恢复或未核对完禁止确认；确认仅清本批已核验备份和记录。旧XML缺Action默认为Replace，缺Checkpoints为空；备份名纳入ManagedFile保护。

### (3). 文件整理与批量命名

整理先只读计划/再执行；确认显示数量、回收风险和逐行原目录，无空行/逐文件细节，完整计划存日志。递归不跟随联接、排除备份/程序文件，提交复查摘要时间；不分类归集本层，分类含根文件。全批预留编号，替换保留最新、同时间优先已有目标，淘汰项/未选类型存“备份”批次相对路径，不覆盖。

OperationBatch保存意图/实际路径/摘要时间/目录身份/清理进度，重开继续；旧记录缺快照不猜范围。取消核对还原文件、仅删本批新建且空目录；身份变/外加内容/失败保留可重试，未完禁确认。确认严格回收列出的原目录含后来新增内容，但保护分类结果/原备份/执行后新目录，不跟联接，不降级永久删除；原备份仅清本批子目录。

重命名只扫第一层、绑打开时簿、新清单同名编号；路径/原名/新名/状态四列，身份可排序不可改增删，执行保存清单。主体保扩展、完整名可改后缀不转内容，空白跳过/同名不改；非法保留名/分隔/越界/重复/占用/源变化整批暂停。交换/循环/大小写用唯一临时名两阶段提交及反向还原，逐步记录实际身份，冲突不覆外部修改；状态回写失败不改变文件事实，确认只清本批记录不回收原目录。

## 4. 文件提取名称与云端OCR

### (1). 扫描、线程与工作表写入

FileNameEngine根深度1不跟联接，字符提取只改名称文字不筛文件。扫描/图片/HTTP后台，输入先STA捕获，控件/网格/确认/工作表预检经OnUi BeginInvoke，不依赖SynchronizationContext；关闭先取消并等待释放请求/临图再关闭，不同时Dispose请求中的HttpClient。

FileNameWorksheet开窗身份只定位原工作簿，起始列默认1；EnsureResultSheet确保“OCR提取结果”，存在先清空/否则末尾新建，重捕身份激活，逐格预检保护/合并/数组与快照，失败尝试还原、不自动保存；不写开窗原表。名称/后缀/路径三列，OCR增加第四文本列保前导零。保留已有筛选/行显隐，隐藏行逐格补写并读回核对，错误按快照恢复；无筛选才开启。数据与文本格式写完后，对实际输出的表头及全部行执行AutoFit，表头居中加粗；测量时临时显露输出范围内隐藏行并在finally恢复原Hidden，不改筛选条件。超宿主列宽的正文换行及适配行高，列宽／换行／相关行高纳入失败还原，不能用预览网格30字符上限限制工作表。

### (2). 服务商能力、凭据与错误

| 服务                                            | 当前实现                                                                                                                      |
| --------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------- |
| 百度General/Accurate/GeneralBasic/AccurateBasic | general/accurate/general_basic/accurate_basic；标准最长4096且编码后8M、高精8192/10M、最短15；Base64表单，适用接口才发粒度参数                          |
| 百度PDF                                         | 整份pdf_file+pdf_file_num只识别统一页码（默认1），本机不渲染；伪PDF文件头上传前拒绝                                                                    |
| 阿里RecognizeAdvanced                           | 杭州端点/`2021-07-07`，ACS3-HMAC-SHA256规范query/headers/原始body摘要，body仅图片；NeedRotate/NeedSortPage/Row=true，尺寸16～8191、比例<50、≤10MB |
| 结果解析                                          | 阿里双层JSON优先prism_rowsInfo，缺行用四点矩形，不拼content；百度含位置/无位置分开                                                                    |

OcrImages多页TIFF按页白底PNG，按尺寸/实际编码或二进制大小等比缩小不改原图；太小/超限/不合法比例停该页。接受PNG/JPG/JPEG/BMP/TIF/TIFF/PDF不接Office，阿里不接PDF，标状态并跳过，图片批次继续。

IFileOcr带页/位置/行标识等信息，每页只识别一次。FileOcrSelection的OcrCombined按前缀保存候选与选定值：位置信息按垂直重叠及1.5倍行高水平间隙成行，无位置逐条处理，不跨页拼接；多字段遇下一已配置标记截断，重叠优先长标记，仅去字段边界末尾的、，,；;。未勾选规则可作边界；单前缀仍沿用OcrFeatures.Extract。提取后再按每个前缀的字符类别筛选OCR候选、去重及判断唯一：中文沿用项目汉字范围，数字和字母分别含半角／全角0—9及英文大小写，其余非空白字符归其他符号；按完整Unicode字符处理，不拆代理对。前缀文字不筛选，前导零、单位及值内符号按所选类别保留，组合分隔符“-”始终保留。

全部启用字段确定后按规则顺序组合到原第四列；缺失/歧义/筛选后为空留空，通过OcrFieldsForm选择或补填，也可整行手改或明确留空；筛选后为空提示“筛选后无可用字符”。缓存OcrSegment只在本窗口内保留，修改前缀或字符规则复算不上传；仍启用且前缀未变的规则按稳定ID保留人工字段选择，人工选定/补填/整行手改不再次过滤，停用后重启则从缓存重新核对；整行手工覆盖和明确留空在关闭全部前缀时仍保留。

OcrPrefixesForm提供启用、前缀文本、中文/数字/字母/其他符号六列，以及增删/上下排序，主窗只读摘要与设置按钮。四类逐前缀独立，新增规则默认全选，启用项至少选一类。OcrPrefixStore将规则ID、前缀、启用状态、顺序和字符类别存本机Excel_Iterate/OcrPrefixRules.xml，独立于凭据；格式版本2兼容读版本1并补全四类，保存才升级。禁止DTD、版本检查、锁内旧摘要比对和原子写入；坏/未知版本配置保留原文件，不静默覆盖。预览只注入样例，测试RootOverride隔离。两个子窗登记file.ocrprefixes/file.ocrfields、默认测量接口、尺寸树子节点及安全预览，Esc取消不提交。

百度OcrCredentials.bin、阿里OcrCredentials.Aliyun.bin独立，默认不记，勾选DPAPI CurrentUser加密原子存，清除只当前家；用户提供权限/配额。上传前确认服务商/版本/页数/计费，串行每页120秒，取消停后续。Token按expires_in缓存，110/111仅重新鉴权重试一次，网络/超时不自动重复，权限/欠费/配额/限流暂停，不自动换服务。

### (3). 界面、诊断与证据

六行布局依次路径、配置标题、扫描参数、前缀摘要与设置/PDF页码/API平台、服务商/版本/凭据、红说明/目标；文字加粗重测，MeasureChoice计字体/DPI/箭头，字段组整组换行。配置容器及字段行自动高度可增长也可回缩；红说明按实际可用宽度、字体、文字与DPI测高后重排父布局，窗口变宽或说明缩短时网格紧接上方并占满余量，防重入而不保留过期高度。保留显式控件尺寸和输入宽度，不清空用户配置。仅本预览隐藏RowHeaders，不改变其他网格的行头设置；Tag仍FileNameEntry。九配置输入/下拉WidthGrip存同一Customization.xml，文件夹Dock拉伸无柄，预览不存；API平台来自API_PLATFORMS生成，空URL禁用、独立小窗320等宽关闭。默认1200×760最小900×620。

结果网格初始常显名称/后缀/路径/OCR识别/候选/识别状态六列，未运行OCR的批次显示“未启用OCR”、禁用候选操作；工作表输出仍为非OCR三列、OCR四列。此网格单独关闭持续Fill列宽，局部策略在共享样式及自定格式后仍保持；其他业务网格仅统一表头规范，保留其正文列宽策略。每次新扫描开始适配周期，目录扫描结束和OCR结束各测一次真实文字宽度，正文上限为当前字体30个“0”加内边距、通常下限50逻辑像素；完整单行表头所需宽度优先，居中且默认加粗，显式自定格式可关闭粗体。最小列宽包含实际表头字体、内边距和排序图标，手动缩窄不能截断表头，按钮列按完整文字定宽。用户手动调整列宽后，本周期的识别/重试/核对/规则调整/缩放不再覆盖；自动上限不限制用户手调、不截断内容，横向滚动保留。重新扫描重启适配，不新增列宽跨窗口持久化。

排障：服务/版本→图片预处理→鉴权/签名→权限配额→解析/候选。OcrDiagnostics仅安全阶段/服务/接口/码/RequestId，UiDiagnostics仅类型/HRESULT/线程/摘要/堆栈，不含URL/凭据/Token/签名/原图/识别文字或异常Message，写盘失败仍可复制。ResultGrid OnPaint捕异常画稳定提示，不冒成功，字体独立克隆/释放；单图失败留行，重试未完成不重跑已确定结果。

Feature099Probe覆盖百度接口/Token、无位置隔离、阿里固定签名与双JSON、取消/凭据/TIFF、PDF请求/拒绝、结果表/版式、无同步上下文填表绘制关闭。字符筛选检查单类/组合、每前缀差异、全角及代理对、空值/候选去重、人工覆盖、缓存复算与版本1迁移；界面检查初始六列、限宽及手调优先、说明换行增高/回缩，检查当前显示DPI及100/125/150/200%字体尺度模拟。既有两图四接口实调用核编号，第三原图缺失仅文本提取；PDF指定页、阿里及复杂图片仍缺真实账号验收，隐藏WinForms/假表不代表WPS进程通过。实际列宽鼠标拖动、各档真实显示缩放、跨屏DPI及云端识别另记待验收。旧离线OCR不参与构建，历史资料不自动删除。

接口依据：[百度标准版](https://cloud.baidu.com/doc/OCR/s/zk3h7xz52)、[高精度](https://cloud.baidu.com/doc/OCR/s/1k3h7y3db)、[含位置标准](https://cloud.baidu.com/doc/OCR/s/vk3h7y58v)、[含位置高精](https://cloud.baidu.com/doc/OCR/s/tk3h7y2aq)、[百度鉴权](https://ai.baidu.com/ai-doc/REFERENCE/Ck3dwjhhu)、[阿里全文高精](https://help.aliyun.com/zh/ocr/developer-reference/api-ocr-api-2021-07-07-recognizeadvanced)、[ACS3签名](https://help.aliyun.com/zh/sdk/product-overview/v3-request-structure-and-signature)。

## 5. 合并抵消

交互弹窗的宿主锁死根因与构建规范属 Word/Excel 共享门禁，见\[弹窗维护规范\](B3-Gate Control.md#2-弹窗维护规范)；本章只记 Excel 侧合并抵消的分层、会计模型与写入还原边界。

### (1). 合并抵消的分层与会计模型

合并抵消（审计工具箱子项）由置灰占位转为可用，一个窗体类服务两个按钮（elimination.receivable / elimination.revenuecost），用 EliminationKind 区分标题、提示与业务。三层分工同表头匹配合并：EliminationEngine 纯核只吃文本行、吐分录，零 COM 零 WinForms，可注入假数据离线自测；EliminationHost 复用 MergeWorkbookGateway（定位/只读开簿/Hold/释放）、MergeSheetReader（表头快照、一维/二维数组兼容的 At）、MergeSnapshot（写入前快照）、MergeTargetWriter（批量写+隐藏行补写+读回核对+Restore）、MergeBackup（同目录备份+SHA256）、WorkbookUi（NewSheet 重名编号/DeleteSheet/StyleHeader），只新增三件事：按角色读源数据（合并单元格按锚点值向下填充——往来余额表的单位列常见合并且必须填充后聚合才有意义）、目标表定位/新建（「（新建工作表）」→「抵消分录」表，重名自动编号，登记实际表名供还原删表）、分录→6 列写入计划（分录号/摘要/方向/科目/金额/计算说明，表头行+分录行一个块写走）。

会计模型与用户确认一致：往来抵消按单位聚合，对冲额=min(应收,应付)，分录借 应付账款/贷 应收账款；应收>应付且映射了坏账准备列时，分录借 坏账准备/贷 信用减值损失=坏账准备×对冲额÷应收。收入成本抵消逐行 U=(R−C)×未对外销售比例，借 营业收入 R、贷 营业成本(R−U)、贷 存货 U；借贷始终平衡，收入≤0 或无法解析的行跳过并列入报告。列角色按表头文本猜（应收/应付/坏账/收入/成本/未对外销售比例/单位），用户可在列映射网格覆盖；同一来源（路径|工作表|表头行）重复预检保留已改角色，换来源才按新表头重建。

### (2). 写入、还原与验收边界

交互：来源区（工作簿/工作表/表头行）→「读取表头并预检」灌列映射网格 → 再点预检读数据算分录 → 分录预览 →「执行写入」前二次确认+备份+快照 → 写入后「取消操作」还原（新建表整表删除 / 已存在表快照回写并自动保存）、「确认操作」结束本批（备份文件保留）。目标=来源工作簿内的指定工作表或新建表；起始单元格越界即拒；确认对话框列出来源/目标/分录数/借贷合计/跳过行数。两个按钮共用一个窗体槽位：同 kind 激活、异 kind 且旧窗空闲才换窗（旧窗网关先释放，FormClosed 令牌判断避免误清新窗资源），旧窗忙或本批未确认则提示先处理。关机与宿主切换时关窗并同步释放 COM 引用。

EliminationProbe 92 项锁定金额/比例解析、角色猜测、映射校验、两种抵消的极端场景（比例 0/100%/中间值、坏账冲回 5000×40000÷100000=2000、p=100% 借100000/贷成本60000/贷存货40000）、6 列写入计划与起始单元格解析。弹窗登记 6 处（DialogSizes/DialogPreviewFactory/DialogSizeEditor/dialog_catalog_probe/GridHeadersProbe/DialogSizesProbe 计数 29→31）并同步 C-02。待验收：真实 WPS 内点击、跨簿只读打开、分录写入/还原与联机尺寸拖动。

### (3). 收入成本抵消改矩阵式审计底稿

收入成本抵消的输出由「凭证式分录」改为「矩阵式底稿」：每个（核算主体→客商）交易对占固定的一组行，行内容由 `EliminationMatrix.cs` 的行模板决定——借 营业收入 R=Σ收入；贷 营业成本 C=Σ成本；贷 营业成本 (R−C)−U（成本还原）；贷 存货 U（未实现内部交易损益）。**恒等式 R ≡ C+[(R−C)−U]+U 无条件成立**，U 为负（收入<成本）时写在贷方列即实质借方，科目方向固定、金额带符号，因此不存在做不平的可能，也不必靠 if 守卫取舍存货行——这同时根治了旧分录模型在收入<成本时借贷不平衡的缺陷。

三维取值与动态列：筛选条件列（如「是否为合并内关联方」取「是」），按 **(核算主体, 客商, 年度)** 聚合。客商维度是本次补上的——样张的 SUMIFS 只按「主体+年度」取值，一主体多客商时各行会重复计入该主体总额。年度集合按源数据实际出现值**降序**生成，不再写死固定年度；列布局=固定五列（核算主体/明细客商/备注/调整事项/会计科目）+ 每年度三列（分隔列/借/贷）。

未实现损益 U 有两个来源：映射了「未实现内部交易损益」列且本组该列有值 → Σ该列（写 SUMIFS）；否则逐行 Σ(收入−成本)×未对外销售比例 现算（写 SUMPRODUCT）。现算必须逐行加权——比例可能逐行不同，先汇总再乘会失真。口径铁律：源表损益列的公式是 (未审收入−未审成本)×比例，故 Σ收入/Σ成本 必须同取未审列；调整后列因调整列取负而恒为 0，选错会让 U 与 R 错配。

输出为活公式（可勾选降级写数值）。SUMIFS 用整列引用 `$O:$O`（自带条件跳过，安全）；SUMPRODUCT 现算分支**不能用整列**——整列会把第 1/2 行的汇总行与表头文本卷入乘法产生 #VALUE!，且 WPS 对 SUMPRODUCT 内 IFERROR 的数组语义支持不完整、拦不住传染。因此改用**从数据首行到工作表末尾**（`$O$3:$O$1048576`）：既避开表头行，又随源表向下扩展自动纳入新增行——早期版本限定到 UsedRange 底行（`$O$3:$O$2695`）会在源表新增行后静默漏数据。**条件数组与金额因子必须同区间**（整列×区间因维度不一致返回 #N/A）。该分支扫描行数接近整表，界面「引用整列」复选框统一表达这一权衡：勾选=自动扩展（默认，保证不漏），取消=限定当前数据区（重算更快，新增行需重新执行）。根本最优解仍是补全源表「未实现内部交易损益」列公式——那样走 SUMIFS，整列、自动扩展、性能最好。公式锚点：主体/客商 `$A{行}`、`$B{行}`（列绝对行相对），年度表头 `G$2`（列相对行绝对），组内互引用完全相对 `=(G4-H5)-H7`。表头样式复刻用户样张：深蓝底（RGB 1F4E78）白字加粗 10 号、四边细框线、居中，年度合并格数字格式 General"年""度"；不再沿用插件统一的浅紫表头（WorkbookUi.StyleHeader）。另修实机验收四项：①来源工作表下拉空白（SetSource 直接赋 Text 掩盖了 Items 为空，读取失败的日志又因句柄未建被静默吞掉——AppendLog 改为先缓存 OnHandleCreated 再刷出，OnLoad 时 Items 为空自动补读一次）；②「写数值不写公式」等复选框高 DPI 截断——AutoSize 自适应；③列映射配置按「业务+表头文本序列」持久化到 %LocalAppData%\Excel_Iterate\EliminationColumnMap.xml，相同表头自动带出上次角色，校验失败时明确提示修改后才能执行写入。

分层与可插拔：新建 `EliminationMatrix.cs` 作第二条流水线，**旧 `EliminationResult`/`BuildWritePlan`/`DefaultSheetName="抵消分录"` 全部保留给往来抵消**，因此探针原有 92 项断言一条未改；组高=行模板条数，行号锚点全部推导得出，代码里不出现硬编码的 4，第二阶段接入往来抵消只需在 `MatrixTemplates` 增一个方法。写入侧新增 `BuildMatrixPlan`/`PrepareTargetRegion`/`MergeMatrixHeader`/`StyleMatrixHeader`，其中 `PrepareTargetRegion` 在写入前先拆旧合并再清内容（目标若已写过矩阵，残留合并会让批量赋值失败）；快照在清理之前取，保证「取消操作」能还原到写入前原样。顺带修正 `StyleEntryHeader` 命中 `WorkbookUi.StyleHeader` 三参重载（行号硬写 1）导致表头样式错位的既有缺陷，改走四参重载。

探针扩到 144 项（92 项原样保留 + §10 三组真实数据验算 + §11 写入计划与锚点平移）。真实底稿的两点提醒：源表第 1 行常是 SUBTOTAL 汇总行、真表头在第 2 行，故表头行默认改为 2；年份列须能被 `(19|20)\d{2}` 识别，不接受 Excel 日期序列号，避免误当年度。待验收：真实 WPS 内选列→预检→写入整表（含第 1 行 SUBTOTAL、第 2/3 行合并表头）、公式重算正确、跨簿引用与大数据量时的写入耗时。

EliminationSession集中持有网关、预检、目标、双表快照和备份路径，状态区分读取、计算、可执行、写入和可还原。预检与执行前重算使用复制请求及后台纯计算；提交前重新读取并核对来源，确认取消不创建输出表，进入写入后沿用原子写入与快照还原边界。取消操作逆序恢复结果表和差异表；部分还原成功后保留剩余目标与快照以供重试，失败不误报已完成。差异表和结果表不允许选为同一个现有工作表。

### (4). 写入空行与往来抵消

「写入空行」：目标区新增输入框，每个交易对（核算主体→客商）之间空 N 行，0 不空。空行留在写入块内写 null，与「覆盖写入」语义一致（`PrepareTargetRegion` 本就 `ClearContents` 整个矩形）。行号推导统一走 `TemplateCount + GroupGap`，三处必须同步：`Emit`（组首行）、`VerifyIdentity`（恒等式校验）、`Balanced`（借贷平衡）——漏改任一处会把空行后的一组误判为前一组。组内 RowRef 引用 `Top+index` 自动跟随。

**往来抵消**产出两张表，逻辑落在新建的 `EliminationDiff.cs`，与收入成本的矩阵核并列而非交织——往来的组内行数随科目对数量变化，若塞进固定模板机制会破坏 `TemplateCount/GroupStride` 的既有假设。它复用 `MatrixColumn/MatrixRow/MatrixAmount`、`BuildColumns` 与 Host 的写入/快照/备份链路，但构建逻辑独立。

- 选列：截止日（取前 4 位作年度）、核算主体、科目、对方单位、是否合并内关联方、审定数
- 公司对按**数据驱动**配对：只列示真实出现过往来的组合（核算主体×对方单位，仅保留实际存在的组合，不做 C(n,2) 全枚举）；对方单位**精确匹配**（不再前缀匹配，避免「食锦/食锦北站分公司」被误并）；**严格按实际名称匹配、不做别名归一化**（原笔误别名表已删除，源表写成几种不同名称就按几个不同主体处理）
- 取数**净额法**（样张反推：资产端=债权方资产−负债、负债端=债务方负债−资产；65,950,000 例中 netA=65,950,000、netB=−65,950,000 → 资产端=负债端=65,950,000，抵消 65,950,000、差额 0、备注「以小抵大：双方余额相等，全额抵消」）；**债权方=净额(资产−负债)较大的一方，动态判定**，净额保留真实符号（允许为负、不再钳制为 0），抵消=MAX(0,MIN(资产端,负债端))（任一侧为负记 0）；旧的「双向合计」口径已废弃
- 科目对别名：源表「预付款项/预收款项」按输出名「预付账款/合同负债」处理
- 差异表 3 行/年度（三个科目对），抵消 `=MIN(I?,L?)`、差额 `=I?-L?`（资产端−负债端）引用本行；结果表只输出抵消不为 0 的科目对，且**同一（核算主体、对方单位、会计科目）跨多年度合并为一行**（金额分列各年度列，借/贷由键中 A/B 互换、科目不同天然区分）
- **坏账准备抵消**：列映射新增可选「坏账准备」（AM 列，`EliminationRole.BadDebt` 加入往来抵消角色清单）；对被抵消的其他应收款所计提坏账准备做**全额抵消**（联动跟随抵消，`BadDebt=badDebtSums[债权方|债务方|年度|资产科目]`），在债权方行下紧跟追加「借 坏账准备 / 贷 管理费用（对方单位=—）」两行，调整事项标注「坏账准备抵消」、金额写静态值；`BuildResult` 改为按（公司对,债权方,债务方,资产科目）分组以支持紧跟债权方行
- **角色自动匹配收紧**：`GuessRole` 收入成本「收入→单体收入」「成本→单体成本」、两业务「关联方/是否→合并内关联方」，避免「收入调整」「合并外关联方」误判
- **列映射临时记忆**：`EliminationMapStore` 由 XML 持久化改为**内存字典**（键=文件路径|业务|表名），打开弹窗时懒清理已关闭工作簿的条目（`CleanupExcept` 对照 `Application.Workbooks`），避免 XML 持续膨胀
- 两张表金额均写活公式：差异表资产端/负债端用**净额公式** `SUMPRODUCT((主体="…")*(对方="…")*(科目="…")*(LEFT(截止日,4)="年度")*审定数)` 相减（=资产科目合计 − 负债科目合计；对方单位严格按实际名称精确匹配，区间为数据首行→工作表末尾，避开源表头文本，条件数组与因子同区间——维度不一致会 #N/A）；**结果表金额改跨表引用差异表的抵消金额格** `=差异表!M<行号>`——早先复用差异表行内 `=MIN(I?,L?)` 会因列序/行号不同而引用错乱（实机验收：未勾「写数值」时结果表金额全错）。执行顺序「先建差异表取真实表名 → 设 options.DiffSheetName → 重 Build → 写两张表」，避免引用悬空
- **括号必须成对**：曾出现提前闭合导致语法错误，探针已加 `Balanced()` 括号配对断言
- 目标：差异表与结果表各一个下拉（往来抵消时目标区显示「差异表」「结果表」两个下拉，消除「一张下拉管两张表」的歧义；`EliminationRequest.DiffSheetItem`）；备份一次覆盖两张表，「取消操作」两张表一起还原
- **差异表口径**：表头「资产账面余额(审定数)」改「资产端账面净额(审定数)」、「负债账面余额(审定数)」改「负债端账面净额(审定数)」；「差额(资产-抵消)」改为「差额(资产-负债)」；差值由 `资产−抵消` 改为 `资产端−负债端`（净额）；备注**按方向区分建议**——资产方大=对方账上少记，建议核查**对方**未达账项（单据在途、入账时间差）、对方漏记入账及预收账款未重分类；负债方大=我方账上少记，建议核查**本单位**未达账项、暂估未入账及预付账款未重分类
- **差异表新增四列展示净额过程**（11→15 列）：在「资产端账面净额」前插**债权方资产余额/债权方负债余额**、在「负债端账面净额」前插**债务方资产余额/债务方负债余额**，四列即双向四口径 `aAsset/aLiab/bAsset/bLiab` 按债权方/债务方映射（`EliminationDiffRow` 加 `CreditorAsset/CreditorLiability/DebtorAsset/DebtorLiability`）；净额列改引用本行四列做减法——资产端 `=G?-H?`、负债端 `=K?-J?`（删合并净额式 `NetFormula`，改 `DirectionalSumFormula` 方向合计 + `NetSubFormula`）；抵消金额列 I→M、差额 J→N、备注→O，结果表跨表引用随改 `=差异表!M`。生产 `StaticValues` 恒 true 仍写静态值，四列在生产也是数值（公式仅探针/活公式模式生效）
- **格式接入 (2) 规范**：新增 `EliminationHost.FinishMatrixOutput`——金额列会计格式（含千位分隔）、数据区垂直居中、指定列水平居中（差异表序号/年度、结果表调整事项/会计科目）、列宽 AutoFit（含表头+数据）、表头行高 26 下限、表头不换行。差异表/往来结果表/收入成本结果表三处共用。**两个结果表另按用户「格式参考」样张**：表头填充 `#4874CB`（主题色 accent1，theme1.xml 实测值）+ 白字、表头与数据区宋体 11、边框 left=medium + 其余 slantDashDot（xlSlantDashDot=13，左缘/顶缘 xlMedium=-4138）、**会计格式从 R1 SUBTOTAL 行起铺设**（fmtTop=headerRow−1）
- **「写数值不写公式」复选框删除**：写公式多次在宿主核对环节报错（跨表引用 `$` 规范化等），`request.StaticValues` 恒 true、金额恒写静态数值；确认文案同步
- **拖柄加固 + 留白根除**：`WidthGrip.OnPaintBackground` 由空改为铺底色（双保险）；删除 `gripSource/gripTarget` 两个行拖柄；`Ui.Grip` 包按钮时清其底边距，前导 Label 上边距 8→6、非按钮输入框上下边距归零、hint 上下边距 8→2、来源/目标间隔行 8→2、两卡片纵向 Padding 8→4。**多轮压缩无效的真根因（新增 LayoutProbe 离线测量证实）**：AutoSize 行里放 `Dock=Fill` 的子控件时，TLP 读的是子控件「当前实际高度」而非期望高度（循环依赖），卡片渲染高 100px 而内容仅需 46/38px——改 `AutoSize=true + Dock=Top` 后卡片高度由内容决定（46/38px），留白彻底消失。`UiKit.cs` 现由 C-00-Shared 单源维护
- **公式写入核对失败（G4）修复**：结果表跨表引用 `='差异表(2)'!I3` 写入后，宿主把相对引用规范化为 `$I$3` 回读，`NormalizeFormula` 不去 `$` 致逐字符比对失败。修复：NormalizeFormula 追加 `Replace("$","")`；核对失败消息附带期望/实际公式便于诊断；写入顺序改为**先写差异表、后写结果表**（引用目标先有数据）。`MergeHeadersHost.cs` 的宿主专属差异按两产品契约分别验证

# 三. WordUpdater 专属开发经验

## 1. LINK字段与表格选区

### (1). 建立逐格链接

LinkPasteEngine优先在原文档文末临时区PasteSpecial读取WPS生成的LINK代码，删临区并恢复选区；该类型原文档探测失败后回退自建并激活临时文档，finally按不保存关闭，不能保证无新标签闪现。DataType依文本2/RTF1/OLE0，每种先命名再位置参数。Fields.Add逐格重建并沿用剪贴板表名前缀：WPS LINK用R1C1，写A1可能Update成功但结果空，须内容核对。字段形状例如 `LINK Excel.Sheet.12 "C:\路径\Book.xlsx" "Sheet1!R1C1:R4C4" \t \a`。

源按Areas顺序、区域先行后列，目标所选矩形先行后列，格数严格相等、不扩大整表；Selection.Cells失败再Range.Cells，任何单格读取失败都停止，不能跳过后缩小选区。仅校验所选单元格的归属、实际边界、物理唯一性与网格矩形；选区外合并、拆分、嵌套和整表Rows/Columns不可读不阻止链接。选区内合并/嵌套、缺格、列边界错位、跨表/越界、超过500格预检拒绝；清内容用End-1保留单元格结束符。

WPS单格Range.WordOpenXML会重建网格并丢失gridSpan，不能用单格XML证明未合并；读取父表Range.WordOpenXML，仅定位用户实际选中的w:tc并判断其gridSpan/vMerge/hMerge及嵌套，不扫描其他单元格作整表拒绝。WPS横向合并后的ColumnIndex是该行物理单元格序号，须累加前面单元格的gridSpan换算成网格列；否则选区左侧合并会造成误拒或错列。内部入口返回排序后的实际Cell，公开ResolveTargetRect签名保留；剪贴板探测后重新验证同一批Cell，逐格即时获取Range，不再用Table.Cell/Rows.Cells兜底猜目标。

WordUpdater同版本修复也要同步B1对应版本摘要与网站配置PRODUCTS中word的版本摘要、使用边界，否则site_probe会因网站/档案摘要不一致而拒绝交付。先固定正文，再重建Word内嵌网站快照与安装包；单项目交付只更新Word安装包，网站联合正式页面与Excel安装包保持原交付，最终证据在B1审计节回填。

### (2). Ctrl多区域来源与边界

剪贴板可能只有包围盒；只有解析一片且格数不等于目标，才回读仍打开ET的多区域明细。先读Application.Selection；在WPS文字功能区回调内，该属性可能为空或不可用，此时遍历已打开Workbooks的Windows，优先读RangeSelection，不可用再读窗口Selection，保留全部纯数据候选，不能只尝试全局选区一次便返回。全过程不Activate/Select、不新建或关闭来源宿主。

候选必须工作簿FullName、表名匹配，格数等于目标，各片在剪贴板包围盒内且互不重叠，整体外边界与包围盒完全一致；跳过错误格数后继续找下一窗口，多个不同但同样可用的选区拒绝猜测。全部失败时保留原剪贴板模板及数量拒绝，不扩大目标。若包围盒格数恰好等目标仍不触发回读，需核映射或分块复制。真机多选先Range.Select再Copy，单Copy不改变Selection。

“WPS表格无可用选区”只能证明发现的实例没有选区，不能直接归因为文字回调限制。新样本实测默认KET/Excel.Application指向Workbooks.Count=0的隐藏ET，而正常打开的来源在另一实例，文件ROT仍可取得真实Workbook.Application。只补窗口回读会继续失败；不得因已找到一个Application就跳过来源文件登记。

回归必须覆盖实例发现过程：若先通过默认Application创建来源，测试会把来源放在旧代码本来就能找到的实例内，从而漏掉真实故障。CtrlSelectionLiveProbe经LiveProbeSafety复制输入文件，通过正常文件关联打开副本并从文件ROT取得测试源，断言默认实例仍为空；再调用完整LinkPasteEngine验证第1/4行与第1/3行各6条LINK、显示值、更新、合并表头内容/格式与原件摘要。旧DLL在此条件下稳定报源12、目标6，区别于LinkPasteTest的112项正常路径与空Selection注入回归。实际按钮、实机COM和假对象证据分开记录。

预检拒绝与写入失败分开报告；UndoRecord及一次Ctrl+Z效果依宿主，不能承诺所有环境整批撤销，写入前保存用户文档。这些几何与LINK字段规则不能套用Excel Range事务。

### (3). 正文顶层表格

Selection.Information(12)判断表内；Table.Uniform与整表Columns.Count不能作为局部选区合法性的门槛，整表尺寸仅作可选参考。先核选区与表格的Range边界再读Cells，越表时WPS曾返Cells.Count=-1；cell.Range.Tables.Count>1及所选w:tc内的子表格识别选区内嵌套。父表XML无法解析或必要属性无法读取时明确失败，所有目标验证完成后才清除内容。LinkMapProbe保留修复前失败回归并覆盖写入前坐标变化、物理重复、结构变化和Range偏移漂移；LinkPasteTest以LiveProbeSafety自建文档验证合并表头下四格、左侧合并右侧四格的实际选区、字段引用、显示值、更新和选区外内容/格式。离线通过、实机COM与真实按钮点击分别记录。

全选表格仅正文顶层，不含页眉页脚/嵌套：清旧EditableRanges→逐表Editors.Add(-1)→SelectAllEditableRanges(-1)→删临时标记，保留多段选区。表单保护2提前拒绝，异常也清标记并恢复ScreenUpdating。

## 2. 更新、查看、定位与高亮

### (1). 更新与变化比对

UpdateEngine扫描当前选区，ShapeRange/InlineShapes按锚点去重；只有成功更新的LINK文本参与变化比对，图片/OLE单列无文本可比，失败不算变化/不标黄。Ordinal字符串比较，不解析数字；先Trim空格/制表/CR/LF/单元格结束符\a及\0。位置依表格行列、段落、页码、兜底描述，合并格异常可降级。

WordHostStateScope先完整捕获DisplayAlerts和ScreenUpdating，再记录每项触碰和恢复；任一捕获失败零修改、零链接更新，设置部分失败也恢复已触碰状态，恢复错误进入结果及诊断。更新链接与链接查看共用这一边界。EtCleanup只诊断非自有实例，移除两处固定500ms等待，不关闭用户进程。

### (2). 重新定位对象

字段定位Code、内嵌对象Range、浮动Anchor；操作前按稳定序号/特征重取，文档改变要求刷新，定位后交还WPS焦点。Update后旧Result可能失效，必须该次重新取Result。

FlashHighlighter存原HighlightColorIndex、黄色7；通常用Result，ShowCodes用Code。LinkViewForm.JumpToSelectedItem必须先ClearPendingFlash，再由JumpToItemEx读原色并标黄，最后SharedFlash.Add/Restart；同一字段可能重取不同Range包装，不能靠引用去重补救顺序错误。若先标黄后清旧状态，第二次会读入临时黄、清掉新高亮，最终残留黄。还原保留0、其他颜色及原本黄色7，不统一清零；关窗同样恢复。更新结果提示关闭后启动约2.5秒计时。Window.ScrollIntoView逐级兼容，全部失败保留视图，不用Select/Activate替代而改选区。

验收分两层：[LinkViewFormProbe](../C-02-WordUpdater/src/tools/LinkViewFormProbe.cs)调用实际私有定位方法，以私有desktop隐藏窗体/假宿主验证清理顺序；真实KWPS核心序列覆盖自建LINK字段的Result/Code、0/2/7原色、连续及换字段恢复和旧明细拒绝。LINK结果文字是自建DOCX显式缓存样本，不代表来源刷新、真实按钮点击或2.5秒计时器验收。各批断言、旧顺序失败及正式DLL复验记录集中放在交付审计区。

## 3. 源文件、ROT与隐藏ET

### (1). 路径与来源实例

LINK代码反斜杠可能双写，FullName单写；比较折连续反斜杠保留UNC前缀，File.Exists成功不能证明原字符串相等。默认关联打开失败可Explorer选中供用户处理，不改关联。

GetActiveObject只有一个实例可能为隐藏ET。EtSelectionReader每次枚举ROT时同时匹配已注册Application的CLSID及剪贴板来源的完整路径；无论已发现多少Application，都不能漏掉来源工作簿登记。只绑定匹配的Application或来源文件，不再无差别绑定其他ROT对象；来源Workbook.FullName须再次精确核对，然后取其Application，不把Workbook当成Application读取Selection。文件名相同而路径不同、路径缺失、FullName或Application读取失败均不加入候选；LINK双反斜杠与ROT路径共用NormPath，保留UNC前缀。

读取器仅附加已运行实例；来源Workbook由当次ROT条目作用域释放，取得的Application转交外层候选作用域，其他ROT枚举、Workbooks/Windows、选区/Areas/表/簿引用按每次取得登记，逆序逐次ReleaseComObject，重复发现、异常与提前返回同样释放。不使用FinalReleaseComObject，不启动、Activate、Quit或关闭宿主。EtSelectionReaderProbe以假引用覆盖来源工作簿归一、隐藏实例已发现、别名去重、错路径/接口失败及全局/窗口读取与引用平衡；不枚举原生ROT或附加WPS。

### (2). 隐藏ET诊断与进程归属

LINK粘贴/更新可能留下隐藏et.exe。EtCleanup前记PID，快照不完整返回null；后续只检查并记录新增隐藏/窗口状态未知的候选，不发WM_CLOSE、不Kill、不等待退出。调用方仅看到PID差集，并未明确创建或持有实例，不能证明独占归属；既有/并发/未知用户实例保留，CleanedHiddenHosts为0。源文件可能仍被占用，用户保存工作后自行确认关闭对应WPS。EtCleanupProbe以假候选覆盖旧PID、并发、未知窗口与异常边界，不操作真实进程。

LinkMap/Update/LinkView/LinkViewForm等隔离探针保留映射、更新、列表身份与窗口回归；真实映射/字段内容/焦点/源占用与进程归属诊断另记，不以假对象或旧截图当当前宿主验收。

# 四. XyCabin 专属开发经验

## 1. 从零网站与配置接线

从最小“模板+配置→校验/转义→单文件HTML→浏览器交互→行为探针”开始，再接下载/音频/联合交付。site_core.render生成公共正文，template负责结构，控制器处理状态，不手改生成页；普通构建只依项目内资产，无外部CDN/fetch邻近文件。新配置先定义用户可见行为/空态与失败，再生成器校验、模板DOM接线、隔离探针、指南和正式重建。

| 配置                 | 接线与边界                                                                                                                                       |
| ------------------ | ------------------------------------------------------------------------------------------------------------------------------------------- |
| SITE/TEXT/PRODUCTS | 品牌/文案/指南/版本；TEXT按页序拆NAV/AUDIO/HERO/PROJECTS/GUIDE/VERSIONS/DOWNLOAD_GUIDE/WORKFLOW/FUTURE/FEEDBACK/FOOTER小字典末尾合并，键对应data-text或T.xxx，改名同步DOM |
| THEME              | 校验#RRGGBB后CSS变量，action/copy/version/selected/text控制对应局部区域，default_mode仅首次默认                                                                 |
| WORKFLOW/FUTURE    | 已实现流程与讨论规划分开                                                                                                                                |
| DOWNLOADS          | word/excel各baidu/quark，空项禁用、仅无账号密码HTTPS，可含查询提取码                                                                                             |
| DOWNLOAD_GUIDE     | 站内相对PNG内嵌、alt非空、steps非空，四段标题/说明/继续/取消文案仍TEXT，缺项停止生成                                                                                         |
| FEEDBACK           | 邮箱→官网HTTPS映射、新标签noopener noreferrer，无映射mailto；复制仅地址，登录发送由用户完成                                                                               |

文本转义、JS以JSON序列化，DOM写textContent不拼执行HTML。保持导航、六分区、指南项目切换/搜索/空态/展开、版本筛选、复制/打印/返回顶部可操作；指南用户步骤不要夹构建细节。

## 2. 布局、按钮与弹窗

网页卡片/网格是HTML/CSS布局，不能套用加载项WinForms类或宿主表格语义。局部按钮配色限定立即使用、下载两半、版本复制/筛选，悬停保留对应填充；版本选中背景/边框/加粗显示状态，按钮统一45px最小高/字号行高/内边距，下载三角等高，不改全站按钮。max-width会限制width，反馈区用.wrap宽度/13px字体，宽屏单行、窄屏自然换行。

下载主按钮先展示dialog指引，明确继续才发下载、暂不关闭；dialog不支持或缺指引数据则退回直接下载，不阻断功能。下拉未配禁用，支持键盘/Esc/外部点击/项目切换关闭，初焦点preventScroll。菜单CSS常规245px、min-width:0和视口上限；展开清旧位移后getBoundingClientRect测边界，visualViewport width/offsetLeft优先、clientWidth回退，两边留8px，translateX纠溢；resize/scroll重测，长文换行。只改left/right不足以覆盖两侧按钮和缩放。

最小真实验收：深浅主题/两项目/宽窄屏/展开后改变宽度，菜单左右不出可见区；再测键盘/外部关闭/切换与默认/云下载。download_probe几何替身不等于实际移动浏览器或安装验收。

## 3. 下载、字节校验与能力降级

独立HTML Base64生成Blob+临时a.download，只请求不运行EXE；DLL空数据页按百度→夸克降级，皆空提示，无配置不能冒充有包。按当前项目即时读取，不捕获旧选择。Blob Base64按87380字符（4倍数）分块atob→Uint8Array，避免整份二进制字符串；防重复，异常恢复控件可用，成功请求60秒后revoke，异常/pagehide立即释放ObjectURL。

校验分两层：构建侧site_core.verify_installers真正计算SHA-256并与本轮EXE逐字节核对；浏览器site_downloads只核摘要字段格式、解码总字节数并生成Blob，没有自行计算SHA-256。不得把字段存在写成浏览器已验内容摘要。download_probe正式HTML路径验证Blob字节与安装包；真实保存位置/下载许可由浏览器，提示“已请求下载”不等于已保存。

单配置预览仅附成功交付记录且版本/大小/摘要匹配的包；正式交付从空快照编译，attach_installers只改唯一installer-assets JSON区，embedded_snapshot归一空数据后与DLL网页资源比较。网页正文还可能有生成元数据，按当前snapshot规范比较，不扩大为所有原始HTML字符串都相同。

localStorage try/catch，禁存不影响当前操作；复制失败兼容再显示可选文本，旧剪贴板工具值不能单独证明写入失败。浏览器工具拒file时不得绕策略或用HTTP结果代替；file测试列待验收。

## 4. 声音资产与异步播放

SOUND默认正弦640Hz/70ms、gain0.40/max1.0，仅可信主动按钮click、100ms限频，滚动/悬停/输入/滑块无音效。BACKGROUND默认tranquility、volume0.50/max0.90，音效与背景独立增益再经总静音/压缩器。场景id保afternoon/tranquility，名称描述可改；相对MP3越界/缺失构建失败。

宁静平原是用户原MP3的60～111秒、128kbps/44.1kHz立体声、首尾200ms淡化、仅音频流51秒项目片段；原文件摘要保留，普通重建不访问开发机原盘。午后暖阳合成琴音。Base64转ArrayBuffer后decodeAudioData，不截编码字节冒可解码片段；MP3解码缓存循环，暂停保当前进度、换曲从头，位置不持久化。

xy-home-audio version2将旧百分比除2仅迁移一次、支持小数，旧sound开关保留；用户静音/0/曲目优先，删曲回默认、畸形或禁存储不阻本次。加载试自动播放，suspended静默；若autoplay与play_on_first_interaction都开，首次可信普通导航/指南/主题click补一次，键盘可信click也可，模拟click/音频控件/搜索输入/音量面板排除。音效开关不阻补播，可信事件内先resume、再异步解码，复用未完resume避免多源。

暂停/静音/背景0/换曲/pagehide以代次令牌取消未完播放和补播资格，不能由后续普通点击恢复；手动播放可重试，旧源/调度器释放，手动解码失败显示普通状态仍可用其他网页功能。AudioProbe是隔离DOM/AudioContext状态回归、不发声；真实效果依系统音量/压缩器/录音，不能将增益倍数承诺为听感倍数。
