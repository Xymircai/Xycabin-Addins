# -*- coding: utf-8 -*-
"""
C2-Wordupdater_Configure.py —— 「WordUpdater」按钮集中配置生成器

一、完整流程（就三步）
==============================================================================

  ① 在本文件顶部「按钮配置区」改好内容，然后运行本脚本
         python AA-Program\\C2-Wordupdater_Configure.py
     想换自定义图标的话，把图丢进项目根的 icons\\ 目录（可选）。
     注意：这一步只是"改配置 + 生成源码"，做完 WPS 里不会有任何变化。

  ② 双击 AA-Program\\一键重建并打包.bat
     编译 DLL、给两份清单签名、重新打包安装器，全自动。
     它自己也会跑一次第 ① 步，所以第 ① 步其实是可选的 —— 先单独跑的
     好处是能立刻发现配置写没写错。批处理完成后提示你自行保存文档并完全退出 WPS，
     再询问是否打开安装器；它不检测或关闭 WPS，也不自动安装。

  ③ 让它在 WPS 里生效（这步手动执行）
       1) 保存文档并完全退出 WPS 文字，避免已加载的 DLL 占用安装目录
       2) 运行 AA-Program\\WordUpdater_Setup.exe，点「安装」；安装目录会回填上次的位置
       3) 安装器先检查 VSTO Runtime；缺少运行时时停止并显示解决说明，
          安装官方运行时后重新打开安装器，不以写入白名单替代运行时安装
       4) 安装器在本次安装过程中调用「Microsoft Office 自定义项安装程序」
          并等待结果；若出现信任窗口，点击「安装」。同一证书已有有效信任
          时可能不再询问。取消或失败会明确提示，不能当作安装成功
       5) 安装器报告成功后重开 WPS 文字，顶部出现「WordUpdater」选项卡即成功

  签名身份只由下方 SIGNING_CERT_THUMBPRINT 决定；常规重建不创建新证书。
  指定证书、有效期或私钥不可用时停止签名；build 缓存不参与证书选择。

二、会改动哪些文件
==============================================================================
  写入 src/addin/Ribbon.g.cs            —— 自动生成，整文件覆盖，请勿手改
  写入 src/installer/InstallerText.g.cs —— 自动生成，整文件覆盖，请勿手改
  写入 icons/README.txt                 —— 图标规格说明，整文件覆盖
  同步 src/addin/AssemblyInfo.cs       —— 按配置更新程序集版本，值未变化时不写文件
  只读  Connect*.cs                     —— 核对操作、图标及显隐回调的存在性和签名（不修改）

三、为什么只跑本py脚本不行？
==============================================================================
  本脚本生成 Ribbon.g.cs 与 InstallerText.g.cs 两份配置源码，而 WPS 真正加载的是安装目录里那个已经编译好、签好名的二进制 DLL。中间隔着四层：Ribbon.g.cs →编译→ WordUpdater.dll →签名+打包→ Setup.exe →安装→ 安装目录的 DLL

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
TAB_ID = "tabWordUpdater"
TAB_LABEL = "WordUpdater"

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
#   里的每一行都算。「问题反馈」的方框比别的宽，就是因为它那行
#   "…2783935558@qq.com…" 更长。所以：
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
#                / "mail"=内置信封 / "view"=内置文档+放大镜（查看链接）
#                / "table"=内置表格网格+对勾 / "formula"=内置 fx / "filter"=内置漏斗（全选表格）
#                （配了 icon 时自定义图优先，图缺失或被删会自动退回内置图）
#   screentip   鼠标悬停时的短标题（加粗，显示在方框最上面一行）
#   supertip    鼠标悬停时的详细说明，可写多行（用三引号），会渲染成多行
#   tip_width   只对本按钮生效的"悬停提示每行宽度"（半角字符数，中文按 2 计），
#               留 None 用全局 TIP_WRAP_WIDTH；两边都没配 = 不自动换行
#   action      "feature"=调用已有功能；"url"=用默认浏览器打开网址
#   callback    action="feature" 时要调用的方法名（必须是 Connect.cs 里的方法，
#               目前只有 OnUpdateLinks / OnLinkToExcel / OnViewLinks /
#               OnSelectAllTables / OnOpenUrl 五个，别乱填）
#   url         action="url" 时的目标地址（http/https/mailto 均可）
#   hidden      True 则整条按钮不写进功能区（临时下线用，留着配置不丢）
#
BUTTONS = [
    {
        "id": "btnUpdateLinks",
        "group": ("grpLink", "链接"),
        "label": "更新链接",
        "size": "normal",

        "icon": "None",             # 默认状态为None，若想换成自己的图：把相应的PNG放入icons这个文件夹，若png的名称为"更新Xy.png"，则这里将“None”替换为"更新Xy"，即可。
        "builtin_icon": "update",
        "screentip": "一键批量更新所选范围内的全部链接",
        "supertip": (

        ),
        "action": "feature",
        "callback": "OnUpdateLinks",
        "url": None,
        "hidden": False,
    },
    {
        "id": "btnLinkToExcel",
        "group": ("grpLink", "链接"),
        "label": "链接Excel",
        "size": "normal",

        "icon": None,
        "builtin_icon": "paste",
        "screentip": "把复制的单元格区域逐格链接到你选中的表格区域",
        "supertip": (
            "用法：先在 WPS 表格里复制一块区域（可按住 Ctrl 多选几个不连续单元格），"
            "\r\n再回到 WPS 文字，用鼠标拖选一个同样格数的矩形区域，然后点本按钮。"
            "\r\n\r\n只认你选中的这个矩形：格数必须刚好相等 —— 多了判溢出、少了判未填满，都会拒绝。"
            "\r\n源按「子区域顺序 + 区域内先行后列」展开，目标按「先行后列」展开，两边按位一一对应。"
            "\r\n想贴整张表，就先把整张表选上；只点一下光标 = 只链接 1 格。"
            "\r\n（按住 Ctrl 多选复制时，请保持 WPS 表格处于打开状态。）"
            "\r\n\r\n链接建立后，源数据变动时选中该表格「更新链接」"
        ),
        "action": "feature",
        "callback": "OnLinkToExcel",
        "url": None,
        "hidden": False,
    },
    {
        "id": "btnViewLinks",
        "group": ("grpLink", "链接"),
        "label": "查看链接",
        "size": "normal",

        "icon": "None",             # 换成自己的图：把 PNG 放进 icons\\，这里写主文件名即可
        "builtin_icon": "view",
        "screentip": "查看本文档引用了哪些Excel文件及其工作表/区域",
        "supertip": (

        ),
        "action": "feature",
        "callback": "OnViewLinks",
        "url": None,
        "hidden": False,
    },
    {
        "id": "btnSelectAllTables",
        "group": ("grpAction", "操作"),
        "label": "全选表格",
        "size": "large",

        "icon": None,
        "builtin_icon": "table",
        "screentip": "选中本文档中的全部表格（正文不受影响）",
        "supertip": (
            "自动滚动到每个表格并按住 Ctrl 逐个框选，效果等同手工"
            " Ctrl 多选：选区里只有表格、不含正文。\r\n\r\n"
            "选中后即可用功能区的表格工具 / 字体等命令对所有表格统一调整。"
        ),
        "action": "feature",
        "callback": "OnSelectAllTables",
        "url": None,
        "hidden": False,
    },
    {'type': 'menu', 'id': 'menuFileOperations', 'label': '文件操作', 'group': ('grpAction', '操作'), 'size': 'large', 'builtin_icon': 'document', 'children': [{'type': 'menu', 'id': 'menuFileConversion', 'label': 'PDF 工具', 'builtin_icon': 'document', 'children': [{'id': 'btnCurrentPdf', 'label': '转为 PDF', 'builtin_icon': 'document', 'action': 'feature', 'callback': 'OnCurrentPdf', 'supertip': '在窗口中切换当前文档与批量文件，支持文件夹和后缀筛选。'}, {'id': 'btnWatermarkPdf', 'label': '编辑 PDF', 'builtin_icon': 'document', 'action': 'feature', 'callback': 'OnWatermarkPdf', 'supertip': '在窗口中切换添加水印、拆分与合并，共用来源清单及顺序。'}]}]},
    {
        "id": "btnFeedback",
        "group": ("grpFeedback", "反馈"),
        "label": "问题反馈",
        "size": "large",

        "icon": None,
        "builtin_icon": "mail",
        "screentip": "有BUG或新增功能请发送邮箱至2783935558@qq.com，我会尽快回复（或许）。",
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
#   「更新链接」「链接Excel」执行完会弹一个结果对话框，下面可以整体改写
#   那句话。不写（空字典）就用内置默认文案。
#
#   写法：键 = 按钮 id（必须和上面 BUTTONS 里的一致），值 = 文案模板。
#   模板里放占位符，程序执行完会把实际数字填进去；\r\n 照旧表示换行。
#   占位符写错会在运行本脚本时当场报错并提示可用的有哪些。
#
#     更新链接（OnUpdateLinks）可用占位符：
#       {found}    发现的链接数          {success}  成功数
#       {failed}   失败数                {changed}  数值发生变动的条数
#       {cleaned}  自动释放的后台进程数
#     链接Excel（OnLinkToExcel）可用占位符：
#       {rows}     目标表格行数          {cols}     目标表格列数
#       {cells}    建立的链接格数        {refreshed} 已刷新数
#       {failures} 刷新失败数            {source}   源描述（工作簿+区域）
#       {cleaned}  自动释放的后台进程数
#
#   三条安全规则（写自定义模板也甩不掉）：
#     · 有失败时，失败明细会自动附在文案后面（不附你看不到失败原因）；
#     · 有数值变动时，变动明细会自动附在文案后面，变动处会标黄 2.5 秒；
#     · 清理过后台进程、而你没用 {cleaned} 时，会自动附一句释放说明。
#
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

RESULT_TEMPLATES = {
    "btnUpdateLinks":
        "链接更新完成。\r\n"
        "共发现 {found} 个链接：成功 {success} 个，失败 {failed} 个，"
        "数值变动 {changed} 个。\r\n",

    "btnLinkToExcel":
        "链接粘贴完成。\r\n"
        "已按 {rows} 行 × {cols} 列建立 {cells} 个链接，其中 {refreshed} 个已刷新。\r\n"
        "源：{source}",
}

"""\r\n\r\n 换两次 → 中间空一行,用来分段
    \r\n    换一次 → 紧贴下一行,同一段内继续"""


# =============================================================================
#                        安 装 程 序 界 面 配 置 区
#           （WordUpdater_Setup.exe 的界面文字，与按钮配置互不影响）
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
#       INSTALLER_VERSION_LINE_PREFIX = "Wordupdater" + INSTALLER_VERSION = "0.9.5.0"
#         + INSTALLER_COPYRIGHT = "Xymircai Initial Development"
#         => 安装器底部居中显示：Wordupdater Version：0.9.5.0，Xymircai Initial Development
#       INSTALLER_COPYRIGHT = "" 时
#         => 安装器底部居中显示：Wordupdater Version：0.9.5.0
#
INSTALLER_FEATURE_TEXT = (
    "扩展：Agent扩展切换 Agent-Evo，提供 AI 网站与离线文档阅读。\n"
    "「链接Excel」将 WPS 表格区域逐格粘贴为表格链接；\r\n"
    "「更新/查看链接」批量更新、总览全文链接并定位；"
    "「其他」全选表格、问题反馈等。"
)
INSTALLER_VERSION = "0.9.9.3"

# --- 固定签名证书（证书指纹，不含私钥）---
# 常规重建复用此证书；证书或私钥不可用就停止，不按主题猜选，也不自动创建。
# 仅在明确更换发布者证书时修改此值；清空 build 不会改变签名身份。
SIGNING_CERT_THUMBPRINT = "81B72008EC4CDBBBC6CCED3A2EB8378812F5E231"  # CN=WordUpdater Dev，与主题一致并带私钥
INSTALLER_COPYRIGHT = "Xymircai Initial Development"
INSTALLER_VERSION_LINE_PREFIX = "Wordupdater"

# --- 安装器窗口标题栏 / 顶部色带主标题 / 副标题（均可自定义） ---
#   INSTALLER_WINDOW_TITLE  窗口标题栏文字
#   INSTALLER_TITLE         顶部色带里的主标题（大字）
#   INSTALLER_SUBTITLE      顶部色带里的副标题（小字，主标题下方）
INSTALLER_WINDOW_TITLE = "WordUpdater - 安装程序"
INSTALLER_TITLE = "WordUpdater"
INSTALLER_SUBTITLE = "WPS 文字加载项 · WordUpdater"





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
