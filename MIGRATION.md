# 目录重构对照表（2026-10-19）

基线提交：`70f4656`（tag `pre-restructure-20261019`）
整包备份：`../B-1-Addon-Program-backup-20261019.zip`
重构分支：`restructure/20261019`

## 一、目录映射

| 旧路径 | 新路径 | 说明 |
|---|---|---|
| `<项目>/AP1-程序/` | `<项目>/AA-Program/` | 配置器、启动器、正式产物 |
| `<项目>/AP2-美工/` | `<项目>/AB-Graphic/` | 界面截图 |
| `<项目>/icons/` | `<项目>/AB-Graphic/` | 功能区图标与生成的 `README.txt` |
| `<项目>/src/documentation/*.svg` | `<项目>/AB-Graphic/*.svg` | 界面示意 SVG |
| `<项目>/AP3-说明/` | 仓库根 `C-04-Devlog/` | 三项目文档统一维护 |

## 二、配置器改名

| 旧文件 | 新文件 |
|---|---|
| `C-01-ExcelIterate/AP1-程序/Configure_program.py` | `C-01-ExcelIterate/AA-Program/C1-Exceliterate_Configure.py` |
| `C-02-WordUpdater/AP1-程序/Configure_program.py` | `C-02-WordUpdater/AA-Program/C2-Wordupdater_Configure.py` |
| `C-03-Website·Xycabin/AP1-程序/Configure_Site.py` | `C-03-Website·Xycabin/AA-Program/C3-XycabinSite_Configure.py` |

## 三、文档拆分与合并

| 旧文件 | 旧章节 | 新位置 |
|---|---|---|
| 三项目 `AGENTS.md` | 全部 | `C-04-Devlog/A0-Entry.md`（合并重写为四节） |
| 三项目 `使用说明.md` | 全部 | `C-04-Devlog/A-UserGuide.md`（每产品一个一级章节） |
| 三项目 `开发日志.md` | 一. Agent及开发者接手概览 | `C-04-Devlog/B1-Project-Files.md` 第一章 |
| 三项目 `开发日志.md` | 二. 项目档案 | `B1-Project-Files.md` 第二至四章（每项目一章） |
| C-01/C-02 `开发日志.md` | 三. Word与Excel共享开发经验 | `B2-DevExperience.md` 第一章 |
| C-01 `开发日志.md` | 五. Excel专属开发经验 | `B2-DevExperience.md` 第二章 |
| C-02 `开发日志.md` | 五. Word专属开发经验 | `B2-DevExperience.md` 第三章 |
| C-03 `开发日志.md` | 四. XyCabin专属开发经验 | `B2-DevExperience.md` 第四章 |
| C-01/C-02 `开发日志.md` | 四. Word与Excel共享维护纪律 | `B3-Uphold-Discipline.md` 第一章 |
| C-03 `开发日志.md` | 三. XyCabin项目维护纪律 | `B3-Uphold-Discipline.md` 第二章 |

2026-10-20 更新：`B3-Uphold-Discipline.md` 已更名为 `B3-Spec Coding.md`，第一章标题改为「Word与Excel共享门禁」，条款按 R-xx（自动门禁）/ H-xx（人工门禁）分类。

## 四、包内与运行时路径

| 项 | 旧值 | 新值 |
|---|---|---|
| 内嵌文档逻辑路径 | `AP3-说明/开发日志.md`、`AP3-说明/使用说明.md` | `C-04-Devlog/B2-DevExperience.md`、`C-04-Devlog/A-UserGuide.md`（各自整份内嵌） |
| 运行时导出目录 | `<用户目录>/AP3-说明/` | `<用户目录>/C-04-Devlog/` |
| 文档图片资源路径 | `AP2-美工/<文件>`、`src/documentation/<文件>` | `<项目>/AB-Graphic/<文件>` |

## 五、校验入口变化

| 旧入口 | 新入口 |
|---|---|
| `C-0x/src/tools/devlog_probe.py`（五章制、200 字、版本升序） | 薄封装，转调仓库根 `verify_devlog.py` |
| `verify_shared.py` 的共享块逐字比对 | 改为「C-04 恰好五份 + 共享章节只保留一份 + 项目内无副本」 |
| `layout_probe.py` 的 AP1/AP2/AP3 期望 | 改为 AA-Program / AB-Graphic / C-04-Devlog |

## 六、回滚

```bash
git checkout main
git reset --hard pre-restructure-20261019      # 仅丢弃本次重构
# 或从仓库外整包恢复：
#   ../B-1-Addon-Program-backup-20261019.zip
```

重构分支 `restructure/20261019` 保留分阶段提交，可按阶段回退：

1. 目录改名
2. 文档迁移与合并
3. 代码改引用
4. 门禁改写
