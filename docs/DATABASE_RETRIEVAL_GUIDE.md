# 科研台账：数据库检索与导出指南

> 目标：按“实验室署名 + 统计年份 + 收录口径”从 Web of Science、EI Compendex 和中国知网导出候选记录，再由本地系统按成员档案认领、去重和生成台账。Scopus 仅作为补充核查渠道。
>
> 最近核验：2026-08-20。数据库界面、字段名和导出上限可能调整；正式批量导出前，必须先执行本文的“小样本验证”。

## 1. 适用边界

- 本指南只检索**明确署名实验室现名、历史名称或可核验缩写**的成果；仅署名“清华大学机械工程系”等依托单位而未署名实验室的论文，不在当前自动入库口径内。
- 机构检索无法保证“100% 查全”。作者漏署名、数据库地址解析错误、名称变体未覆盖或记录尚未入库，都会造成漏检。
- 数据库端以查全为先，本地端再用 `config/teachers_profile.xlsx` 认领成员；但本地认领不能补回数据库检索阶段已经漏掉的记录。
- 当前自动流程实际读取 WOS、EI 和 CNKI 的固定文件名；Scopus 尚未接入 `vba_modules/Mod2_CleanRawData.bas`，不能把 `scopus.csv` 当作已支持的自动输入。

## 2. 机构名称与检索原则

| 类型 | 已核验名称 |
| --- | --- |
| 中文现名 | `高端装备界面科学与技术全国重点实验室` |
| 中文历史名称 | `摩擦学国家重点实验室` |
| 英文现名 | `State Key Laboratory of Tribology in Advanced Equipment` |
| 英文历史名称 | `State Key Laboratory of Tribology` |
| 官方缩写 | `SKLT` |

实验室官网确认了现名、历史名称及 `SKLT` 缩写。缩写不是唯一标识，不能脱离 `Tsinghua` 单独检索，否则可能混入其他机构。

检索时遵循以下规则：

1. 现名与历史名称之间使用 `OR`；机构、年份和文献类型之间使用 `AND`。
2. 每个机构词都要绑定到数据库对应的机构/地址字段，不能把一组词先括起来、再只给整组末尾添加字段限定，除非平台明确支持这种写法。
3. 年份尽量写入检索式，并同时核对结果页的出版年筛选；系统最终仍按导出记录中的正式出版年过滤。
4. 不使用未经当前数据库机构检索页核验的组织 ID。数据库可能拆分、合并或重建机构档案。

## 3. Web of Science Core Collection

### 3.1 推荐检索式

在 **Web of Science Core Collection → Advanced Search** 中选择 **Science Citation Index Expanded (SCI-EXPANDED)**，再运行：

```text
AD=(
  "State Key Laboratory of Tribology in Advanced Equipment"
  OR "State Key Laboratory of Tribology"
  OR (SKLT SAME Tsinghua)
)
AND PY=2026
```

说明：

- `AD` 是地址字段，`PY` 是出版年字段；`SAME` 将 `SKLT` 与 `Tsinghua` 约束在同一地址中，降低缩写误命中。
- 不再使用 `"State Key Laboratory of Tribology*"`：它把精确短语与截词混在一起，含义不如显式列出现名和历史名称清楚，也不利于逐项验证。
- `SCI-EXPANDED` 应在数据库/引文索引范围中限定，不能只凭结果来自 Web of Science 就标记为 SCI。
- 当前业务口径没有进一步限定 WOS 文献类型。若以后只统计 Article 或 Review，应先确认口径，再增加 `AND DT=(Article OR Review)`；不要在本指南中悄悄缩小成果范围。

将 `2026` 替换为目标正式出版年。跨多年检索可使用平台支持的年份范围，或按年分别检索以便核对数量。

### 3.2 导出

1. 选择 **Export → Tab delimited file**。
2. `Record Content` 选择 **Full Record**；确认导出字段至少包含 `TI`、`SO`、`VL`、`IS`、`AU`、`DI`、`PY` 和 `PD`。
3. 按当前界面允许的最大批次导出。若结果超过单批上限，合并为一个仅保留一次表头的制表符文件，并核对合并前后记录数。
4. 保存为 `raw_data/WOS.txt`。不要同时保留另一个待导入批次并期待系统自动扫描；当前解析器只读取约定文件名。

## 4. EI Compendex（Engineering Village）

### 4.1 推荐检索式

在 **Engineering Village → Compendex → Expert Search** 中运行：

```text
(
  ({State Key Laboratory of Tribology in Advanced Equipment} WN AF)
  OR ({State Key Laboratory of Tribology} WN AF)
  OR (({SKLT} WN AF) AND ({Tsinghua University} WN AF))
)
AND ({JA} WN DT)
AND ({2026} WN YR)
```

说明：

- `AF`、`DT`、`YR` 分别用于机构、文献类型和年份；每个机构候选词都单独绑定 `WN AF`。
- `JA` 是当前项目批准的 EI 期刊论文口径。不要把会议论文类型混入后再统一标记为 EI (JA)。
- `SKLT` 分支必须和 `Tsinghua` 同时出现。Engineering Village 的字段范围不等同于 WOS 的 `SAME`，因此首次使用时必须人工检查该分支的前 20 条及随机 20 条结果；若仍有明显误命中，删除该分支，只保留完整名称。
- Engineering Village 的可用检索代码可能随所选数据库和机构订阅界面变化。若平台拒绝该式，使用 `Search codes` 面板插入 Affiliation、Document type、Year 条件，以平台自动生成的 Expert Search 式为准，并把生成式记录在当年检索日志中。

### 4.2 导出

1. 选择全部目标记录并点击 **Download / Export**。
2. 格式选择 **Excel**，内容选择 **Detailed record**，勾选 **include columns without data**。
3. 确认首行至少包含 `Title`、`Source title`、`Volume`、`Issue`、`Author`、`DOI`、`Publication year` 和 `Issue date`。
4. 合并多批数据后保存为 `raw_data/EI.xlsx`；当前系统不会自动合并多个任意文件名的批次。

## 5. 中国知网（CNKI）

### 5.1 检索配置

进入 **高级检索**，数据库范围选择 **学术期刊**，按以下条件配置：

| 逻辑 | 字段 | 匹配方式 | 值 |
| --- | --- | --- | --- |
| 第一行 | 作者单位 | 精确短语或平台最接近的短语匹配 | `高端装备界面科学与技术全国重点实验室` |
| 或者 | 作者单位 | 精确短语或平台最接近的短语匹配 | `摩擦学国家重点实验室` |
| 并且 | 发表年度 | 精确 | 目标年份，如 `2026` |

重要说明：

- 使用界面中的“或者/OR”连接两条作者单位条件；不要把 `A + B` 直接输入一个文本框。`+` 在不同版本界面中不一定代表 OR，可能被当作普通字符或其他运算。
- 期刊来源类别选择 **核心期刊（北大核心）** 与 **CSCD** 的并集，并在结果页确认两类筛选是 OR，而不是同时满足的 AND。
- 若当前界面无法明确表达 OR，分别运行现名和历史名称两次检索，利用导出结果的 DOI/题名去重；不要凭猜测改用未核验的“专业检索”字段代码。
- CNKI 的界面和可导出字段受机构订阅及版本影响。本节给出的是字段配置，不声称是一条可跨版本复制的专业检索式。

### 5.2 导出

1. 选择 **导出与分析 → 自定义导出**（或当前界面的 Excel 导出入口）。
2. 至少导出：题名、作者、文献来源、年、卷、期、DOI、作者单位；页码、基金项目可一并保留。
3. 确认表头能对应 `Title-题名`、`Source-文献来源`、`Volume-卷`、`Period-期`、`Author-作者`、`DOI-DOI`、`Year-年`。
4. 保存为 `raw_data/CNKI.xls` 或 `raw_data/CNKI.xlsx`。若分两次检索，合并时只保留一次表头，并核对合并记录数。

## 6. Scopus（补充核查，不是当前自动输入）

Scopus 可用于查漏或交叉核对。不要使用原文档中的 `AF-ID("State Key Laboratory of Tribology" 60021634)`：该 ID 尚未在当前 Scopus 机构选择器中核验，名称与 ID 的组合不能靠猜测固化。

不依赖机构 ID 的推荐式为：

```text
(
  AFFILORG("State Key Laboratory of Tribology in Advanced Equipment")
  OR AFFILORG("State Key Laboratory of Tribology")
  OR AFFIL(SKLT AND Tsinghua)
)
AND SRCTYPE(j)
AND PUBYEAR IS 2026
```

- `AFFILORG` 检索机构部分；`AFFIL(SKLT AND Tsinghua)` 要求两个词出现在同一 affiliation 中。
- `SRCTYPE(j)` 表示期刊来源，不等于只保留 Article。若业务口径需要“仅 Article”，另行增加 `DOCTYPE(ar)`，并明确它会排除 Review 等其他期刊文献类型。
- 若通过 Scopus 的 affiliation selector 找到并核对了目标机构 ID，可用选择器自动插入 `AF-ID(...)`，但应先比较“ID 检索”与“名称检索”的结果差集，再决定是否替换。
- 当前 VBA 流程不读取 `scopus.csv`；导出文件只能人工核查，不能宣称会被一键流程合并。

## 7. 批量导出前的小样本验证

每个数据库和每个目标年份至少完成一次：

1. **逐分支计数**：分别运行现名、历史名称、缩写分支，记录各自数量及 OR 合并后的去重数量。
2. **已知样本召回**：选取至少 3 篇确定署名实验室的论文（包含现名、历史名称；如有则包含缩写署名），确认均被检出。
3. **误命中抽查**：检查每个分支前 20 条和随机 20 条，确认地址确属目标实验室；重点检查 `SKLT` 分支。
4. **年份核对**：抽查 online first 与正式卷期跨年的记录，确认导出字段中的正式出版年符合统计口径。
5. **导出完整性**：导出行数应与平台选中记录数一致；合并批次后仅有一个表头，关键字段不为空或错列。
6. **留痕**：记录数据库、所选子库、完整检索式、检索日期、各分支命中数、总数和人工抽查结论。数据库持续更新，同一检索式在不同日期得到不同数量是正常现象。

## 8. 自动化处理闭环

```text
数据库：实验室署名 + 年份 + 已批准的收录口径
                     │
                     ▼
       raw_data/WOS.txt、EI.xlsx、CNKI.xls(x)
                     │
                     ▼
正式出版年过滤 → 成员别名认领 → DOI/题名去重 → 复合收录标记
                     │
           ┌─────────┴─────────┐
           ▼                   ▼
【课题组入库成果】      【未认领排除成果】
```

## 9. 核验依据

- [Web of Science Core Collection：Advanced Search Field Tags](https://webofscience.help.clarivate.com/en-us/Content/wos-core-collection/woscc-search-field-tags.htm)
- [Web of Science：Search Rules](https://webofscience.help.clarivate.com/en-us/Content/search-rules.htm)
- [Scopus Support：Advanced Search 字段代码与运算规则](https://www.elsevier.support/scopus/answer/how-can-i-best-use-the-advanced-search)
- [清华大学机械工程系：State Key Laboratory of Tribology in Advanced Equipment](https://www.me.tsinghua.edu.cn/en/info/1249/1874.htm)
- [实验室官网：现名、历史名称与 SKLT 缩写](https://sklt.tsinghua.edu.cn/info/1157/2468.htm)
