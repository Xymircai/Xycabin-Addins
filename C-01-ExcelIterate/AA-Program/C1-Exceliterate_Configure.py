# -*- coding: utf-8 -*-
"""
C1-Exceliterate_Configure.py —— 「Excel-iterate」按钮集中配置生成器

一、完整流程（就三步）
==============================================================================

  ① 在本文件顶部「按钮配置区」改好内容，然后运行本脚本
         python AA-Program\\C1-Exceliterate_Configure.py
     想换自定义图标的话，把图丢进项目根的 icons\\ 目录（可选）。
     注意：这一步只是"改配置 + 生成源码"，做完 WPS 里不会有任何变化。

  ② 双击 AA-Program\\一键重建并打包.bat
     编译 DLL、给两份清单签名、重新打包安装器，全自动。
     它自己也会跑一次第 ① 步，所以第 ① 步其实是可选的 —— 先单独跑的
     好处是能立刻发现配置写没写错。批处理完成后提示你自行保存文档并完全退出 WPS，
     再询问是否打开安装器；它不检测或关闭 WPS，也不自动安装。

  ③ 让它在 WPS 里生效（这步手动执行）
       1) 保存文档并完全退出 WPS 表格，避免已加载的 DLL 占用安装目录
       2) 运行 AA-Program\\Excel_Iterate_Setup.exe，点「安装」；安装目录会回填上次的位置
       3) 安装器先检查 VSTO Runtime；缺少运行时时停止并显示解决说明，
          安装官方运行时后重新打开安装器，不以写入白名单替代运行时安装
       4) 安装器在本次安装过程中调用「Microsoft Office 自定义项安装程序」
          并等待结果；若出现信任窗口，点击「安装」。同一证书已有有效信任
          时可能不再询问。取消或失败会明确提示，不能当作安装成功
       5) 安装器报告成功后重开 WPS 表格，顶部出现「Excel-Iterate」选项卡即成功

  签名身份只由下方 SIGNING_CERT_THUMBPRINT 决定；常规重建不创建新证书。
  指定证书、有效期或私钥不可用时停止签名；build 缓存不参与证书选择。

二、会改动哪些文件
==============================================================================
  写入 src/addin/Ribbon.g.cs            —— 自动生成，整文件覆盖，请勿手改
  写入 src/addin/ApiPlatforms.g.cs      —— 自动生成，整文件覆盖，请勿手改
  写入 src/installer/InstallerText.g.cs —— 自动生成，整文件覆盖，请勿手改
  写入 icons/README.txt                 —— 图标规格说明，整文件覆盖
  同步 src/addin/AssemblyInfo.cs       —— 按配置更新程序集版本，值未变化时不写文件
  只读  Connect*.cs                     —— 核对操作、图标及显隐回调的存在性和签名（不修改）

三、为什么只跑本py脚本不行？
==============================================================================
  本脚本生成 Ribbon.g.cs、ApiPlatforms.g.cs 与 InstallerText.g.cs 三份配置源码，而 WPS 真正加载的是安装目录里那个已经编译好、签好名的二进制 DLL。中间隔着四层：Ribbon.g.cs →编译→ Excel_Iterate.dll →签名+打包→ Setup.exe →安装→ 安装目录的 DLL

  按钮的标题 / 提示 / 图标 / 网址全都编译进 DLL，不是运行时读的配置，没有热更新这回事。所以第 ②③ 步一步都不能省。

  下面是第 ② 步那个 .bat 的命令行等价版（必须手动敲，用bat可以进行批处理）：

    1) python build.py
    2) python resign.py        （内部自动完成：生成两份清单 + SignCmi 签
                                SHA-256 双签名；旧工具 SignFx 写死 SHA-1 已废弃）
    3) python build_installer.py
"""

# =============================================================================
#                              按 钮 配 置 区
#                    （只改这一块就够了，下面全是生成逻辑）
# =============================================================================

# --- 选项卡：默认是一个独立的自定义选项卡 ---
#   TAB_ID    自己造的标识，只含字母数字下划线，改了不影响功能
#   TAB_LABEL 选项卡上显示的名字（可以写中文、可以带符号）
TAB_ID = "tabEtTrace"
TAB_LABEL = "Excel-Iterate"

# 独立选项卡插在哪：None = 追加到所有选项卡最右；
# 想插到某个内置选项卡前面就填它的 id，例如 "TabInsert"（插入）、"TabHome"（开始）
TAB_INSERT_BEFORE_MSO = None

# 想改回"塞进某个内置选项卡"，把下面填成内置 id（如 "TabHome"），
# 填了它以后上面三行就忽略。想用独立选项卡就保持 None。
TAB_ID_MSO = None

# --- 自定义图标目录（相对项目根；改标点的话连带把图搬过去）---
ICON_DIR = "AB-Graphic"

# --- 图片透明区域的合成底色 ---
#   WPS 功能区按钮只接收 HBITMAP，不支持真透明通道；带透明像素的 PNG 会先把
#   透明部分合成到这个颜色上。默认取功能区浅灰底 #F5F6F7。
#   若你的图本身不带透明（推荐），这项可以忽略。
ICON_BACKGROUND = "#F5F6F7"

# --- 悬停提示方框（screentip / supertip）的宽度 ---
#   先说清原理：这个方框没有"宽度=xxx"的开关，宽度是由【最长的一行文字】
#   自动撑出来的 —— screentip（悬停时的加粗标题行）和 supertip（详细说明）
#   里的每一行都算。若某个按钮的说明里写了很长的网址/邮箱，方框会
#   被那一行撑得更宽。所以：
#
#     · 想让一句话【整行显示、不折行】：别在句子里插 \r\n，把它写成一行，
#       方框会自己变宽。
#       （WPS 对方框有内部宽度上限，行太长最终仍会被 WPS 自己折行，这层上限属于宿主行为，任何配置都改不了）
#
#     · 想反过来把方框【限宽】（自动按宽度断行），把下面设成整数：
#         TIP_WRAP_WIDTH = 40
#       单位 = 半角字符宽（一个中文/全角字符按 2 计）。
#       None = 不自动换行（默认，保持 WPS 原生行为）。
#
#   单个按钮想单独限宽：在该按钮里加一行 "tip_width": 40（优先于全局值）。
TIP_WRAP_WIDTH = None

# --- 按钮列表：顺序 = 功能区里的排列顺序 ---
#   每一层 children 列表也按书写顺序显示；调整同层完整条目的位置即可排序。
#   例如 menuWrapFormula 的 children 当前依次为 IFERROR、ROUND、ABS。
#   只移动完整条目，保留 id/callback；修改后需重新构建、签名、打包并手动重装。
#
# 字段说明
#   type        默认 button；menu 配 children，不配置执行回调。
#               嵌套项继承根分组，不填 group/size，图标为 16 个逻辑像素。
#   id          唯一标识，只含字母数字下划线，改了不影响功能，但别重名
#   group       二元组 (分组id, 分组显示名)；相同分组 id 的按钮排在同一组
#   label       按钮标题
#   size        "large"(32 逻辑像素，上图下文) / "normal"(16 逻辑像素，图文横排)
#   column      可选容器 id；同组连续的根层 normal 按钮或菜单用同一值，上下排为一列
#   icon        自定义图标文件名（放在 icons\\ 下）；留 None 用下面的 builtin_icon
#               写字符串 "None" 效果一样（很多人习惯这么写）。
#               文件名可以不带扩展名，只写主名如 "更新Xy"，会自动找
#               icons\\更新Xy.png / .jpg / .bmp / .ico
#   builtin_icon None=只有自定义图 / "update"=内置循环箭头 / "paste"=内置网格+链条
#                / "mail"=内置信封 / "view"=内置文档+放大镜
#                / "table"=内置表格网格+对勾 / "formula"=内置 fx / "filter"=内置漏斗
#                （配了 icon 时自定义图优先，图缺失或被删会自动退回内置图）
#   screentip   鼠标悬停时的短标题（加粗，显示在方框最上面一行）
#   supertip    鼠标悬停时的详细说明，可写多行（用三引号），会渲染成多行
#   tip_width   只对本按钮生效的"悬停提示每行宽度"（半角字符数，中文按 2 计），
#               留 None 用全局 TIP_WRAP_WIDTH；两边都没配 = 不自动换行
#   action      "feature"=调用已有功能；"url"=用默认浏览器打开网址
#   callback    action="feature" 时要调用的方法名（必须是 Connect.cs 里的方法，
#               包括追踪、公式及数值计算、按列筛选操作，菜单本身不填 callback）
#   url         action="url" 时的目标地址（http/https/mailto 均可）
#   hidden      True 则整条按钮不写进功能区（临时下线用，留着配置不丢）
#
# 默认数值与“默认加/减/乘/除…”文案从同一设置生成。
# 改默认值后需重建并手动重装；自定义窗口只作用于本次操作。
# default_parameter 只用于数值计算默认叶子；children 列表顺序决定菜单顺序。
ARITHMETIC_DEFAULTS = {"add": 100, "subtract": 100, "multiply": 100, "divide": 10000}

BUTTONS = [
    {
        "id": "btnTracePrecedents",
        "group": ("grpTrace", "数据"),
        "label": "追踪引用",
        "size": "normal",

        "icon": None,
        "builtin_icon": "view",
        "screentip": "解析当前单元格公式，逐项展开被引用单元格的详情",
        "supertip": (
            "选中含公式的单元格后点击，列出该公式直接引用的全部单元格："
            "\r\n工作簿、工作表、地址、引用类型（同表/跨表/跨簿）、绝对/相对。"
            "\r\n\r\n双击列表某条记录可跳转并高亮闪烁；跨簿未打开时会提示后打开再跳。"
        ),
        "action": "feature",
        "callback": "OnTracePrecedents",
        "url": None,
        "hidden": False,
    },
    {
        "id": "btnTraceDependents",
        "group": ("grpTrace", "数据"),
        "label": "追踪从属",
        "size": "normal",

        "icon": None,
        "builtin_icon": "view",
        "screentip": "遍历所有已打开工作簿，找出哪些单元格的公式引用了当前单元格",
        "supertip": (
            "选中单元格后点击，遍历所有已打开工作簿的所有工作表公式，"
            "\r\n反向找出引用了当前单元格的单元格并列出详情。"
            "\r\n\r\n大表扫描时显示进度条并可取消；双击列表记录跳转高亮。"
        ),
        "action": "feature",
        "callback": "OnTraceDependents",
        "url": None,
        "hidden": False,
    },
    {'type': 'menu',
     'id': 'menuFilterEnhanced',
     'group': ('grpTrace', '数据'),
     'label': '筛选增强',
     'size': 'normal',
     'builtin_icon': 'filter',
     'screentip': '按活动单元格所在表格的表头选择显示列',
     'children': [{'id': 'btnFilterColumns',
                   'label': '按列筛选',
                   'builtin_icon': 'filter',
                   'action': 'feature',
                   'callback': 'OnFilterColumns'},
                  {'id': 'btnQuickAdvancedFilter', 'label': '快捷高筛', 'builtin_icon': 'filter', 'action': 'feature', 'callback': 'OnQuickAdvancedFilter'}]},
    {
        "type": "menu",
        "id": "menuAuditToolbox",
        "group": ("grpTrace", "数据"),
        "label": "审计工具箱",
        "size": "large",
        "icon": None,
        "builtin_icon": "table",
        "screentip": "审计辅助工具集合：表头匹配合并、合并抵消",
        "supertip": (
            "表头匹配：按表头列名（忽略大小写与首尾空格）匹配列，"
            "\r\n把来源工作簿的列数据合并到目标工作簿同名列下。"
            "\r\n\r\n合并抵消：往来抵消按单位对冲生成分录；收入成本抵消按（核算主体,客商,年度）三维聚合生成矩阵式底稿。"
            "\r\n两者列映射角色均可改，写入前备份、可一键还原。"
        ),
        "children": [
            {
                "id": "btnMergeHeaders",
                "label": "表头匹配",
                "builtin_icon": "table",
                "screentip": "按表头列名匹配，把多个来源工作簿的列追加到目标工作簿同名列下",
                "supertip": (
                    "选择目标工作簿与若干来源工作簿，按表头列名匹配列，"
                    "\r\n把来源列的数据追加到目标工作簿同名列已有数据的下方。"
                    "\r\n\r\n每一列各自追加到该列已有数据的末尾；目标缺失的列名会先列出并请你确认，"
                    "\r\n不会自动新建列。公式按目标位置重写同表引用；跨表与跨簿前缀及列不变，相对行按行差偏移。"
                    "\r\n写入前自动备份目标工作簿，完成后可点「取消操作」还原。"
                ),
                "action": "feature",
                "callback": "OnMergeHeaders",
            },
            {
                "type": "menu",
                "id": "menuElimination",
                "label": "合并抵消",
                "builtin_icon": "table",
                "screentip": "合并报表抵消分录辅助（往来抵消 / 收入成本抵消）",
                "supertip": (
                    "往来抵消按单位对冲出凭证式分录；收入成本抵消按三维聚合出矩阵式底稿（年度列动态、金额写活公式）。"
                    "\r\n列映射角色可改，写入前备份、可一键还原。"
                ),
                "children": [
                    {
                        "id": "btnEliminationReceivable",
                        "label": "往来抵消",
                        "builtin_icon": "table",
                        "screentip": "内部往来科目抵消",
                        "supertip": "按（核算主体,对方单位,年度）双向合计后以小抵大，生成「往来抵消差异」与「往来抵消结果」两张表；年度取截止日前 4 位；三个科目对：应收应付、其他应收其他应付、预付账款合同负债。",
                        "action": "feature",
                        "callback": "OnEliminationReceivable",
                    },
                    {
                        "id": "btnEliminationRevenueCost",
                        "label": "收入成本抵消",
                        "builtin_icon": "table",
                        "screentip": "内部交易收入成本抵消",
                        "supertip": "按（核算主体,客商,年度）三维聚合生成矩阵式抵消底稿：借 营业收入、贷 营业成本、贷 成本还原、贷 存货；年度列按数据动态生成，金额写活公式，写入前备份、可一键还原。",
                        "action": "feature",
                        "callback": "OnEliminationRevenueCost",
                    },
                ],
            },
        ],
    },
    {
        "type": "menu", "id": "menuOneClickFormula", "group": ("grpAction", "操作"),
        "label": "一键公式", "size": "large", "builtin_icon": "formula",
        "screentip": "封装或移除当前选区公式的最外层函数",
        "supertip": "仅点击最终选项才修改公式；支持连续选区和 Ctrl 多选。",
        "children": [
            {
                "type": "menu", "id": "menuWrapFormula", "label": "进行公式封装",
                "builtin_icon": "formula",
                "children": [
                    {"type": "menu", "id": "menuIFErrorWrap", "label": "IFERROR封装", "builtin_icon": "formula",
                     "children": [
                         {"id": "btnWrapIFError", "label": "默认返回空", "builtin_icon": "formula",
                          "action": "feature", "callback": "OnWrapIFError", "supertip": '增加 IFERROR(原公式,"") 一层。'},
                         {"id": "btnCustomIFError", "label": "自定义封装", "builtin_icon": "formula",
                          "action": "feature", "callback": "OnCustomIFError", "supertip": "设置错误时返回的文字、数值或公式/引用，确认后处理选区；取消不修改。"},
                     ]},
                    {"type": "menu", "id": "menuRoundWrap", "label": "ROUND封装", "builtin_icon": "formula",
                     "children": [
                         {"id": "btnWrapRound", "label": "默认保留2位", "builtin_icon": "formula",
                          "action": "feature", "callback": "OnWrapRound", "supertip": "增加 ROUND(原公式,2) 一层，固定保留两位小数。"},
                         {"id": "btnCustomRound", "label": "自定义封装", "builtin_icon": "formula",
                          "action": "feature", "callback": "OnCustomRound", "supertip": "自定义整数位数，例如 10、0 或 -2，确认后处理选区；取消不修改。"},
                     ]},
                    {"id": "btnWrapAbs", "label": "ABS封装", "builtin_icon": "formula",
                     "action": "feature", "callback": "OnWrapAbs", "supertip": "增加 ABS(原公式) 一层，使用标准单参数语法。"},
                    {"id": "btnCustomFunction", "label": "自定义函数", "builtin_icon": "formula",
                     "action": "feature", "callback": "OnCustomFunction", "supertip": "输入任意函数模板，留一个空参数位表示每个单元格内容，如 =TEXTSPLIT(,\"-\",1) 或 =ISBLANK()。"},
                ],
            },
            {"id": "btnUnwrapFormula", "label": "移除公式封装", "builtin_icon": "formula",
             "action": "feature", "callback": "OnUnwrapFormula",
             "supertip": "只拆一层完整包住整条公式的函数，保留第一个参数，丢弃其他参数。"},
    {'type': 'menu',
     'id': 'menuArithmetic',
     'label': '数值计算',
     'builtin_icon': 'formula',
     'children': [{'type': 'menu',
                   'id': 'menuCalculateAdd',
                   'label': '加',
                   'builtin_icon': 'formula',
                   'children': [{'id': 'btnCalculateAdd',
                                 'label': '默认加',
                                 'default_parameter': 'add',
                                 'builtin_icon': 'formula',
                                 'action': 'feature',
                                 'callback': 'OnCalculateAdd'},
                                {'id': 'btnCustomCalculateAdd',
                                 'label': '自定义计算',
                                 'builtin_icon': 'formula',
                                 'action': 'feature',
                                 'callback': 'OnCustomCalculateAdd'}]},
                  {'type': 'menu',
                   'id': 'menuCalculateSubtract',
                   'label': '减',
                   'builtin_icon': 'formula',
                   'children': [{'id': 'btnCalculateSubtract',
                                 'label': '默认减',
                                 'default_parameter': 'subtract',
                                 'builtin_icon': 'formula',
                                 'action': 'feature',
                                 'callback': 'OnCalculateSubtract'},
                                {'id': 'btnCustomCalculateSubtract',
                                 'label': '自定义计算',
                                 'builtin_icon': 'formula',
                                 'action': 'feature',
                                 'callback': 'OnCustomCalculateSubtract'}]},
                  {'type': 'menu',
                   'id': 'menuCalculateMultiply',
                   'label': '乘',
                   'builtin_icon': 'formula',
                   'children': [{'id': 'btnCalculateMultiply',
                                 'label': '默认乘',
                                 'default_parameter': 'multiply',
                                 'builtin_icon': 'formula',
                                 'action': 'feature',
                                 'callback': 'OnCalculateMultiply'},
                                {'id': 'btnCustomCalculateMultiply',
                                 'label': '自定义计算',
                                 'builtin_icon': 'formula',
                                 'action': 'feature',
                                 'callback': 'OnCustomCalculateMultiply'}]},
                  {'type': 'menu',
                   'id': 'menuCalculateDivide',
                   'label': '除',
                   'builtin_icon': 'formula',
                   'children': [{'id': 'btnCalculateDivide',
                                 'label': '默认除',
                                 'default_parameter': 'divide',
                                 'builtin_icon': 'formula',
                                 'action': 'feature',
                                 'callback': 'OnCalculateDivide'},
                                {'id': 'btnCustomCalculateDivide',
                                 'label': '自定义计算',
                                 'builtin_icon': 'formula',
                                 'action': 'feature',
                                 'callback': 'OnCustomCalculateDivide'}]}]},
        ],
    },
    {'type': 'menu', 'id': 'menuFileOperations', 'group': ('grpAction', '操作'), 'label': '文件操作', 'size': 'large', 'builtin_icon': 'document', 'screentip': '整理文件、批量重命名及 PDF 转换、水印、拆分与合并', 'supertip': '批量处理保留原功能；PDF 支持当前文档、混合批量转换、水印、拆分和合并。', 'children': [{'type': 'menu', 'id': 'menuBatchProcessing', 'label': '批量处理', 'builtin_icon': 'document', 'children': [{'id': 'btnOrganizeFiles', 'label': '文件批量整理', 'builtin_icon': 'document', 'action': 'feature', 'callback': 'OnOrganizeFiles', 'supertip': '把子目录文件移到第一层；可分类、处理同名或按类型送入回收站。'}, {'id': 'btnRenameFiles', 'label': '文件批量命名', 'builtin_icon': 'document', 'action': 'feature', 'callback': 'OnRenameFiles', 'supertip': '扫描文件夹第一层文件，填写清单第三列后执行替换。'}, {'id': 'btnFileNames', 'label': '文件提取名称', 'builtin_icon': 'document', 'action': 'feature', 'callback': 'OnFileNames', 'supertip': '扫描并预览文件清单；OCR可设置多个前缀，逐项核对后按顺序用短横线组合结果。'}]}, {'type': 'menu', 'id': 'menuFileConversion', 'label': 'PDF 工具', 'builtin_icon': 'document', 'children': [{'id': 'btnCurrentPdf', 'label': '转为 PDF', 'builtin_icon': 'document', 'action': 'feature', 'callback': 'OnCurrentPdf', 'supertip': '在窗口中切换当前文档与批量文件，支持文件夹和后缀筛选。'}, {'id': 'btnWatermarkPdf', 'label': '编辑 PDF', 'builtin_icon': 'document', 'action': 'feature', 'callback': 'OnWatermarkPdf', 'supertip': '在窗口中切换添加水印、拆分与合并，共用来源清单及顺序。'}]}]},
    {"id": "btnLinkAssistant", "group": ("grpAction", "操作"), "label": "链接助手", "size": "large", "builtin_icon": "link", "action": "feature", "callback": "OnLinkAssistant", "screentip": "扫描 xlsx 外部引用，C列填新路径或「断开链接」后执行；支持清理所选来源的未使用名称，断链可从整簿备份恢复"},
    {"id": "btnGenerateDirectory", "group": ("grpAction", "操作"), "label": "生成目录", "size": "large", "builtin_icon": "directory", "action": "feature", "callback": "OnGenerateDirectory", "screentip": "链接全部工作表，隐藏表点击后显示并跳转"},
    {
        "id": "btnFeedback",
        "group": ("grpFeedback", "反馈"),
        "label": "问题反馈",
        "size": "large",
        "icon": None,
        "builtin_icon": "mail",
        "screentip": "有BUG或新增功能请发送邮箱至2783935558@qq.com，我会尽快回复（尽量）。",
        "supertip": (
            "点此打开 QQ 邮箱网页版，登录后即可撰写反馈邮件。"
        ),
        "action": "url",
        "callback": None,
        "url": "https://wx.mail.qq.com",
        "hidden": False,
    },
]

# --- 操作完成后弹窗文案的自定义模板 ---
#   本项目（Excel-iterate）的追踪结果在 TrackerForm 弹窗里展示，
#   公式成功使用状态栏，异常/空选区使用 UiDialog；RESULT_TEMPLATES 保持空字典。
# Agent-Evo：网站名称、网址和顺序在此修改；叶子 id 对应专属回调。
CLOUD_APIS = [
    ("btnAgentBaidu", "百度智能云", "https://console.bce.baidu.com"),
    ("btnAgentAliyun", "阿里智能云", "https://home.console.aliyun.com/"),
]
AGENT_SITES_CHINA = [
    ("btnAgentWorkBuddy", "WorkBuddy", "https://copilot.tencent.com/work/"),
    ("btnAgentTrae", "TRAE", "https://www.trae.cn/"),
    ("btnAgentQoder", "Qoder", "https://qoder.com/"),
    ("btnAgentDeepSeek", "DeepSeek", "https://chat.deepseek.com/"),
    ("btnAgentDoubao", "豆包", "https://www.doubao.com/"),
    ("btnAgentKimi", "Kimi", "https://www.kimi.com/"),
]
AGENT_SITES_OVERSEAS = [
    ("btnAgentChatGPT", "ChatGPT", "https://chatgpt.com/"),
    ("btnAgentClaude", "Claude", "https://claude.ai/"),
    ("btnAgentGemini", "Gemini", "https://gemini.google.com/"),
]
RIBBON_ON_LOAD = "OnAgentRibbonLoad"
BUTTONS.append({"id": "btnAgentExtension", "group": ("grpExtension", "扩展"),
    "type": "toggleButton", "label": "Agent扩展", "size": "large",
    "action": "feature", "callback": "OnToggleAgentExtension",
    "get_pressed": "GetAgentExtensionPressed", "builtin_icon": "agent",
    "screentip": "显示或隐藏 Agent-Evo", "supertip": "包含 AI 网站、离线开发文档和弹窗尺寸设置；按用户记住本项目的选择。"})
EXTRA_TABS = [{"id": "tabAgentEvo", "label": "Agent-Evo", "get_visible": "GetAgentTabVisible", "buttons": [
    {"id": "menuAgentApi", "group": ("grpAgentResources", "详情"), "type": "menu",
     "label": "调用API", "size": "large", "builtin_icon": "web", "children": [
        {"id": "menuAgentCloud", "type": "menu", "label": "智能云-API", "builtin_icon": "web", "children": [
            {"id": bid, "label": label, "action": "url", "url": url, "builtin_icon": "web"}
            for bid, label, url in CLOUD_APIS]},
        {"id": "menuAgentSites", "type": "menu", "label": "Agent-API", "builtin_icon": "agent", "children": [
            {"id": "menuAgentChina", "type": "menu", "label": "国内Agent", "builtin_icon": "agent", "children": [
                {"id": bid, "label": label, "action": "url", "url": url, "builtin_icon": "web"}
                for bid, label, url in AGENT_SITES_CHINA]},
            {"id": "menuAgentOverseas", "type": "menu", "label": "国外Agent", "builtin_icon": "agent", "children": [
                {"id": bid, "label": label, "action": "url", "url": url, "builtin_icon": "web"}
                for bid, label, url in AGENT_SITES_OVERSEAS]},
        ]},
    ]},
    {"id": "menuAgentWorkflow", "group": ("grpAgentResources", "详情"), "type": "menu",
     "label": "开发流程", "size": "large", "builtin_icon": "document", "children": [
        {"id": "btnAgentReadme", "label": "开发预览", "action": "feature", "callback": "OnOpenAgentReadme", "builtin_icon": "document"},
        {"id": "btnAgentUserGuide", "label": "用户预览", "action": "feature", "callback": "OnOpenAgentUserGuide", "builtin_icon": "document"},
        {"id": "btnAgentHome", "label": "XyCabin", "action": "feature", "callback": "OnAgentHome", "builtin_icon": "agent"},
    ]},
    {"id":"menuAgentCustomization","group":("grpAgentOperations","操作"),"type":"menu","label":"自定配置","size":"large","builtin_icon":"settings","children":[
        {"id":"btnAgentSort","label":"自定排序","action":"feature","callback":"OnAgentSort","builtin_icon":"settings"},
        {"id":"btnAgentNames","label":"自定名称","action":"feature","callback":"OnAgentNames","builtin_icon":"settings"},
        {"id":"btnAgentFormats","label":"自定格式","action":"feature","callback":"OnAgentFormats","builtin_icon":"settings"},
        {"id":"btnAgentDialogSizes","label":"交互弹窗","action":"feature","callback":"OnOpenAgentDialogSizes","builtin_icon":"settings"},
    ]},
    {"id":"menuAgentExperiments","group":("grpAgentOperations","操作"),"type":"menu","label":"实验计划","size":"large","builtin_icon":"settings","children":[
        {"id":"btnAgentOcrPlaceholder","label":"OCR识别","action":"feature","callback":"OnAgentOcrPlaceholder","builtin_icon":"document","enabled":False},
        {"id":"btnAgentRequirements","label":"需求记录","action":"feature","callback":"OnAgentRequirements","builtin_icon":"document"},
    ]},
]}]

RESULT_TEMPLATES = {}

"""\r\n\r\n 换两次 → 中间空一行,用来分段
    \r\n    换一次 → 紧贴下一行,同一段内继续"""


# =============================================================================
#                        安 装 程 序 界 面 配 置 区
#           （Excel_Iterate_Setup.exe 的界面文字，与按钮配置互不影响）
# =============================================================================
#
# 这里的内容会被生成进 src\\installer\\InstallerText.g.cs，随后打包进安装器。
# 只改这里不用重编加载项 DLL，但安装器 exe 必须重新打包才会变
# （一键重建并打包.bat 会自动做；手动的话最后跑 python build_installer.py）。
#
#   INSTALLER_FEATURE_TEXT  顶部色带里的功能说明文字；
#                           太长会在色带内自动换行，不会冲出窗口；
#                           空字符串 = 不显示这一行
#   INSTALLER_VERSION       版本号，必须是四段式数字，如 "0.9.5.0"
#                           （每段 0~65534；不许留空、不许写成三段或五段）
#                           ★ 全项目唯一的版本准绳：AssemblyInfo.cs 与 VSTO
#                             清单都由它自动跟随，别在别处手写版本号
#                           ★ 版本规则：只按用户明确指定的版本修改，不自动递增。
#   INSTALLER_COPYRIGHT     版本号后面那句自定义文案（底部一行，加粗蓝字）；
#                           空字符串 = 只显示 "Version：0.9.5.0"，
#                           不会留下一个孤零零的逗号
#   INSTALLER_VERSION_LINE_PREFIX  版本行最前面的产品名前缀（可留空 ""）；
#                           非空时显示为：前缀 + 空格 + Version：x.x.x.x
#
#   底部那一整行是脚本拼出来的，不用你自己拼：
#       INSTALLER_VERSION_LINE_PREFIX = "Exceliterate" + INSTALLER_VERSION = "0.9.5.0"
#         + INSTALLER_COPYRIGHT = "Xymircai Initial Development"
#         => 安装器底部居中显示：Exceliterate Version：0.9.5.0，Xymircai Initial Development
#       INSTALLER_COPYRIGHT = "" 时
#         => 安装器底部居中显示：Exceliterate Version：0.9.5.0
#
INSTALLER_FEATURE_TEXT = (
    "扩展：Agent扩展切换 Agent-Evo，提供 AI 网站与离线文档阅读。\n"
    "「追踪引用」解析公式逐项展开被引用单元格详情；\r\n"
    "「追踪从属」反向找出引用当前单元格的全部单元格；"
    "点击记录跳转并高亮闪烁。\r\n"
    "「一键公式」封装与计算；「筛选增强」列显示与快捷高筛；\r\n"
    "「链接助手」批量换链；「生成目录」快速跳转工作表。"
)

# -----------------------------------------------------------------------------
# 「跳转到API平台」弹窗选项（文件提取名称窗口的跳转按钮会读取这一份配置）
# -----------------------------------------------------------------------------
# 每一项 = (显示名称, 网址)；想加几个就加几个，顺序即弹窗里的排列顺序。
#
#   网址填完整地址 -> 点击后用系统默认浏览器打开；
#   网址留空 ""    -> 该选项显示为灰色、不可点击（占位，例如“未来计划”）。
#
#   例：下次想新增一个选项，直接在下面追加一行即可，无须改 C# 代码：
#       ("XX", "https://……"),
#
# 注意：改完要重跑 python C1-Exceliterate_Configure.py，再 python build.py 重新编译，
#       网址会生成进 src/addin/ApiPlatforms.g.cs 并随 DLL 一起签名打包。
API_PLATFORMS = [
    ("百度智能云", "https://cloud.baidu.com/product-s/qianfan_home"),
    ("阿里云平台", "https://www.aliyun.com/?spm=a2c4g.11174283.0.d_logo.5fb0215fSiDrJz"),
    ("未来计划", ""),
]
INSTALLER_VERSION = "0.9.9.3"

# --- 固定签名证书（证书指纹，不含私钥）---
# 常规重建复用此证书；证书或私钥不可用就停止，不按主题猜选，也不自动创建。
# 仅在明确更换发布者证书时修改此值；清空 build 不会改变签名身份。
SIGNING_CERT_THUMBPRINT = "8EA393A55F624AE71F3BD6ADCE2732B84171900B"  # CN=Excel_Iterate Dev，与主题一致并带私钥
INSTALLER_COPYRIGHT = "Xymircai Initial Development"
INSTALLER_VERSION_LINE_PREFIX = "Exceliterate"

# --- 安装器窗口标题栏 / 顶部色带主标题 / 副标题（均可自定义） ---
#   INSTALLER_WINDOW_TITLE  窗口标题栏文字
#   INSTALLER_TITLE         顶部色带里的主标题（大字）
#   INSTALLER_SUBTITLE      顶部色带里的副标题（小字，主标题下方）
INSTALLER_WINDOW_TITLE = "Excel-Iterate - 安装程序"
INSTALLER_TITLE = "Excel-Iterate"
INSTALLER_SUBTITLE = "为数据开发而生"





# 公共生成逻辑使用隔离模块，导入配置时不生成文件、不写字节码缓存。
import importlib.util as _importlib
import os as _os
import sys as _sys
_old_bytecode = _sys.dont_write_bytecode
_sys.dont_write_bytecode = True
try:
    _spec = _importlib.spec_from_file_location("_wps_configure_core", _os.path.join(_os.path.dirname(__file__), "..", "src", "configuration", "configure_core.py"))
    _core = _importlib.module_from_spec(_spec)
    _core.__dict__.update({k: v for k, v in globals().copy().items() if k.isupper()})
    _spec.loader.exec_module(_core)
    _core.bind(globals(), __file__)
finally:
    _sys.dont_write_bytecode = _old_bytecode
get_version = _core.get_version
sync_assembly_info = _core.sync_assembly_info
check_installer_version = _core.check_installer_version
main = _core.main
if __name__ == "__main__":
    main()
