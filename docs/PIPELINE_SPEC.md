# 科研台账 (keyantaizhang) - 全流程与跨数据库字段映射规范说明书 (PIPELINE_SPEC)

> **产品全称**：科研台账 (keyantaizhang)  
> **核心定位**：课题组学术成果智能整理与多维统计工作台 (Research Group Scholarly Output & Ledger Hub)  
> **版本**：v4.3 (课题组学术成果台账与双核解耦版)  
> **更新时间**：2026-08-19  
> **核心原则**：
> 1. **全相对路径原则**：所有 VBA 与脚本严禁硬编码绝对路径，必须通过 `ThisWorkbook.Path` 动态获取根目录，保证项目随处拷贝移动 100% 正常运行。
> 2. **双核解耦架构**：采用独立的 [`Mod_Sync.bas`](../vba_modules/Mod_Sync.bas) 作为底座负责秒级热载入，业务全流程调度在 [`Mod0_ControlPanel.bas`](../vba_modules/Mod0_ControlPanel.bas)，支持随时无损更新。
> 3. **7 大标准终稿字段直出**：首列“序号”与末列“收录类型”（支持 `SCI+EI` 复合收录标识），严格过滤非本组人员论文。
> 4. **正式出版年份日期过滤**：严格以 `PY` / `Publication year` / `Year-年` 为准，提供单行极简交互配置，默认填充当年 2026 全年度。

---

## 一、 系统顶层架构与目录设计

```
📁 科研台账系统根目录/              # 项目根目录（动态自适应相对路径 ThisWorkbook.Path）
│
├── 📂 config/                      # 【人员台账与字典层】
│   ├── teachers_profile.xlsx       # 课题组师生档案库（姓名、团队、方向、全格式拼音库）
│   └── journal_if.xlsx             # 期刊影响因子(IF)与分区检索字典（预留拓展）
│
├── 📂 raw_data/                    # 【原始成果数据层】（只读源数据，永不篡改）
│   ├── WOS.txt / WOS.xlsx          # Web of Science 官方导出文件 (首选 Tab delimited .txt)
│   ├── EI.xlsx / EI.xls            # EI Compendex 官方导出文件 (首选 Excel 详细记录)
│   ├── CNKI.xls / CNKI.xlsx        # 中国知网 (CNKI) 官方导出文件 (首选 自定义Excel)
│   └── scopus.csv                  # Scopus 官方导出文件 (备用)
│
├── 🎮 console_dashboard.xlsm       # 【科研台账交互工作台】（日常单工作表操作工作台）
├── 🏆 papers_final_merged.xlsx     # 7大核心字段终稿表（课题组成果交付大表）
│
├── 📂 vba_modules/                 # 【核心 VBA 源代码模块】（纯文本 .bas，Git版本受控）
│   ├── Mod_Sync.bas                # 【⭐ 独立热更底座】专职负责一键秒级无损同步所有模块并重构UI
│   ├── Mod0_ControlPanel.bas       # 控制台交互、状态看板、业务全流程调度与交付报告引擎
│   ├── Mod1_TeacherPinyin.bas      # 师生多格式拼音与检索特征构建子引擎
│   └── Mod2_CleanRawData.bas       # 多源抽取、日期校验、消歧、去重、人员过滤与 7 列直出子引擎
│
├── 📂 docs/                        # 【系统文档与变更规范】
│   ├── PIPELINE_SPEC.md            # 全流程业务与跨数据库字段映射规范说明书（本文件）
│   └── CHANGELOG.md                # 需求变更与版本迭代日志
│
├── 📂 .vscode/                     # VS Code 编辑器配置（自动识别 GBK 编码）
├── 📄 README.md                    # 用户使用操作手册
└── 📄 .gitignore                   # Git 排除规则（忽略数据文件，追踪源码文档）
```

---

## 二、 跨数据库多源字段全景对照映射表 (Cross-Database Field Matrix)

本系统支持同时接入 **Web of Science (WOS)**、**EI Compendex**、**中国知网 (CNKI)** 及 **Scopus** 等多源数据，各数据库字段映射及提取对齐标准如下：

| 交付终稿字段<br>(`papers_final_merged.xlsx`) | Web of Science (WOS)<br>(`WOS.txt` / `WOS.xlsx`) | EI Compendex<br>(`EI.xlsx` / `EI.xls`) | 中国知网 (CNKI)<br>(`CNKI.xls` / `CNKI.xlsx`) | Scopus (Elsevier)<br>(`scopus.csv`) | 核心清洗与提取转换规则 |
| :---: | :---: | :---: | :---: | :---: | :--- |
| **序号** (Col A) | *(自动生成)* | *(自动生成)* | *(自动生成)* | *(自动生成)* | 过滤非本组人员后，自增纯数字编号（1, 2, 3...）。 |
| **论文题目** (Col B) | `TI`<br>`Article Title` | `Title` | `Title-题名`<br>`题名` | `Title` | • 压缩连续多余空格与换行符；<br>• 去除题目末尾多余的英文句点 `.`。 |
| **期刊名称** (Col C) | `SO`<br>`Source Title` | `Source` | `Source-文献来源`<br>`文献来源` | `Source title` | • 英文期刊统一规范为 **Title Case**；<br>• 保护专用缩略词大写（如 IEEE/ASME/CFD）；<br>• 虚词/介词/连词小写（如 of/in/and/for）。 |
| **卷** (Col D) | `VL`<br>`Volume` | `Volume` | `Volume-卷`<br>`卷` | `Volume` | 纯净卷号提取，若原库为空则保留为空。 |
| **期** (Col E) | `IS`<br>`Issue` | `Issue` | `Period-期`<br>`期` | `Issue` | 纯净期号提取，若原库为空则保留为空。 |
| **作者** (Col F) | `AU`<br>`Authors` | `Author` | `Author-作者`<br>`作者` | `Authors` | • **EI 机构角标剥离**：正则去除 `(1,2)` 等数字标签；<br>• **师生档案多维匹配**：根据 `teachers_profile.xlsx` 全拼音别名消歧识别；<br>• **仅保留本组人员**，多位成员分号 `; ` 间隔；<br>• **无本组成员者自动排除出终稿**。 |
| **收录类型** (Col G) | 标记为 `SCI` | 标记为 `EI` | 标记为 `中文核心` | 标记为 `Scopus` | • 跨库合并时，自动组合为复合标签（如 **`SCI+EI`**）；<br>• 排序权重优先级：`SCI` > `EI` > `中文核心`。 |
| *(去重键: DOI)* | `DI`<br>`DOI` | `DOI` | `DOI-DOI`<br>`DOI` | `DOI` | 统一转换为全小写、去空格，作为跨库合并主键。 |
| *(去重键: 归一化题目)* | `TI` (正则纯净串) | `Title` (正则纯净串) | `Title-题名` (纯净串) | `Title` (纯净串) | 仅提取中英文字符与数字作为二级去重备用主键。 |
| *(时间过滤基准: 年份)* | **`PY`**<br>`Publication Year` | **`Publication year`** | **`Year-年`**<br>`年` | `Year` | 严格以**正式出版年份**为核心过滤基准。 |
| *(时间过滤辅助: 日期)* | `PD`<br>`Publication Date` | `Issue date` | `PubTime-发表时间` | `Date` | 辅助提取具体月份与日期（英文月份转数字）。 |

---

## 三、 各核心字段清洗与消歧规则细则

### 1. 论文题目 (Title) 清洗规则
- **去空格与换行**：将 `\r`, `\n`, `\t` 全部转为空格，并消除连续双空格；
- **去末尾句点**：英文数据库（如 WOS / EI）题目结尾常带有多余句号 `.`，算法统一 `RTrim` 清除末尾标点，确保同一篇论文在不同数据库导出的题目完全一致。

### 2. 期刊名称 (Journal) 英文 Title Case 与专用词典规则
- **中文期刊保护**：若包含中文字符（Unicode $\ge$ `0x4E00`），保持原样中文刊名输出；
- **英文期刊大小写规整（Title Case）**：
  - **首字母与主词大写**：每个单词首字母大写，后续字母小写（如 *Journal of Manufacturing Processes*）；
  - **专用学术缩写全大写保护词典**：
    `IEEE`, `ASME`, `ACM`, `CFD`, `AI`, `PINN`, `FEM`, `CNC`, `MEMS`, `NEMS`, `IC`, `LED`, `OLED`, `UV`, `NMR`, `SEM`, `TEM`, `AFM`, `XPS`, `XRD`, `PANI`, `PDMS`, `UHMWPE`, `PBO`, `SOFC`, `DES`, `LES`, `RANS`, `1D`, `2D`, `3D`, `4D`, `5D`, `I`, `II`, `III`, `IV`, `V`, `VI`, `VII`, `VIII`, `IX`, `X`；
  - **从属小写虚词词典**（非句首时小写）：
    `of`, `in`, `and`, `the`, `on`, `for`, `to`, `at`, `by`, `with`, `from`, `into`, `onto`, `via`, `a`, `an`, `as`, `nor`, `but`, `part`, `section`；
  - **连字符复合词规整**：如 `Bio-Inspired`、`Micro-Machining` 连字符前后单词分别智能大写。

### 3. 作者姓名 (Authors) 智能消歧与认领规则
- **EI 机构标签清洗**：EI 导出的作者往往带有机构数字上标，如 `Liu, Huichao(1); Chen, Yan(1,2); Zheng, Quanshui(3,4,5)`。
  - 正则表达式：`\(\s*[\d\s,]+\s*\)` 替换为空白，恢复纯净英文姓名。
- **多格式全拼音消歧字典匹配**：
  - 从 [`config/teachers_profile.xlsx`](../config/teachers_profile.xlsx) 读取每位成员的全部变体（如 `Zhang Jianfu`、`Zhang, J.`、`Zhang, JF`、`Zhang, Jian-Fu` 等）；
  - 算法对作者字符串进行无视大小写、无视标点符号、无视空格的哈希匹配；
  - **去重认领**：同一位成员在一篇论文中多次出现仅保留一次；
  - **非本组排除**：若整篇论文的所有作者均不在课题组师生档案库中，该论文判定为“非本组/未认领”，**绝对不写入终稿表**，计入看板排除统计。

### 4. 跨数据库去重与复合收录合并规则
- **去重主键（优先级 1: DOI）**：
  - 提取原始 DOI，统一转为小写字符串；
  - 若多库记录的 DOI 相同，自动合并为同一条成果。
- **去重备用键（优先级 2: 归一化题目哈希）**：
  - 若部分中文核心或会议论文无 DOI，提取题目并移除非字母数字字符转为纯小写字符串；
  - 当题目完全相同时自动识别合并。
- **复合收录类型合成**：
  - 若 WOS 命中且 EI 命中，收录类型合成并格式化为 **`SCI+EI`**；
  - 字段互补：若 WOS 卷期缺失而 EI 完整，自动用非空卷期补齐。

### 5. 正式出版日期与时间范围过滤规则
- **正式出版年份（`PY` / `Publication year` / `Year-年`）一票准入**：
  - 统计严格以正式刊登的出版年份为准（排除跨年的网络预发表干扰）；
  - 若记录中含有具体月份/日期（如 `MAY 15 2026` 或 `2026-06-20`），转为标准日期对象与设定的起止区间比较；
  - 若记录中**仅有年份**（如 `2026`），只要年份在设定起止区间内，即判定为在期保留。

---

## 四、 控制台交互设计与 680pt 几何对齐规范

控制台工作表 [`console_dashboard.xlsm`](../console_dashboard.xlsm) 严格遵循现代化仪表盘工业设计规范：

```
[ 顶部横幅 Banner: 20 ~ 680pt ] -> [ 重置与同步代码 (325~465pt) ] + [ >>> 一键执行全流程 <<< (475~670pt) ]
[ 数据概览看板: 3张卡片 (20~230pt, 245~455pt, 470~680pt) ]
[ 原始数据清单 (左栏 B~D) ]      [ 交付成果大表与单行排除备注 (右栏 H~J: 470~680pt) ]
[ 成果发表时间工具条 (Row 20) ]  [ 起始: D20 | 截止: F20 | 预设按钮: 440~680pt ]
[ 操作流程与工作原理看板 (Row 23~27) ] [ 3步极简操作指引 (左栏) | 三大底层核心支柱 (右栏) ]
[ 快捷链接与日志 (Row 29~33) ]   [ 4个快捷按钮 (20~680pt) | 状态日志框 (20~680pt) ]
```

1. **绝对对齐基准**：所有横幅、卡片、表格列宽与操作按钮右边缘**统一锁定在 `680.0pt`**；
2. **纯净 GBK 编码**：全模块代码严禁包含任何 Unicode Emoji 图标，彻底根治乱码与问号；
3. **静默执行与唯一终极弹窗**：执行期间全程关闭屏幕刷新与文件闪烁，执行完毕后弹出多维全息交付报告。
