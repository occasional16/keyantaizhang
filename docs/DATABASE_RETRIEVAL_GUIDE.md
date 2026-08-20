# 科研台账：数据库检索与导出指南

> 目标：按“实验室署名 + 统计年份 + 收录口径”从 Web of Science、EI Compendex 和中国知网导出候选记录，再由本地系统按成员档案认领、去重和生成台账。Scopus 仅作为补充核查渠道。
>
> 文档规则核对：2026-08-20。WOS 检索词已根据当前账号的差分实测修正；EI、CNKI、Scopus 仍须在实际订阅界面逐分支验证。数据库界面、索引方式和导出上限可能调整，正式批量导出前必须执行本文的“小样本验证”。

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
2. 每个机构词都要绑定到数据库对应的机构/地址字段。官网完整名称用于确认机构身份，实际检索词必须适配数据库中的缩写、分词和规范化方式，不能假定地址字段会原样保存完整名称。
3. 年份尽量写入检索式，并同时核对结果页的出版年筛选；系统最终仍按导出记录中的正式出版年过滤。
4. 不使用未经当前数据库机构检索页核验的组织 ID。数据库可能拆分、合并或重建机构档案。

## 3. Web of Science Core Collection

### 3.1 推荐检索式

在 **Web of Science Core Collection → Advanced Search** 中选择 **Science Citation Index Expanded (SCI-EXPANDED)**，再运行：

```text
PY=(2025-2026) AND AD=(("State Key Lab* Tribol*" OR "State Key Laboratory of Tribology*" OR SKLT) SAME (Tsinghua OR 100084))
```

说明：

- `AD` 是地址字段。WOS 会缩写许多地址词，实际记录常见 `Univ`、`Lab`、`Tribol` 等索引形式；`Univ*`、`Lab*`、`Tribol*` 用于同时召回缩写和完整拼写。
- `Tribol*` 不能收窄为 `Tribology`。在当前实测中，仅这一变化就造成结果数大幅下降，说明大量地址记录使用了 `Tribol` 等缩写形式。
- `State Key Lab* SAME Tribol*` 同时覆盖实验室历史名称和重组后的现名；无需把含有 `of`、`in` 的官网完整名称作为精确短语写入 `AD`。WOS 官方也建议地址全名检索时避免这些介词。
- `SAME` 要求各组词位于同一条作者地址中。加入 `Tsinghua Univ*` 是为了避免命中其他单位的同类实验室，`SKLT` 分支同样必须受此约束。
- `SCI-EXPANDED` 应在数据库/引文索引范围中限定，不能只凭结果来自 Web of Science 就标记为 SCI。
- `PY=2026` 用于生成候选集，但 WOS 的 `PY` 检索同时考虑 Early Access Year 和 Final Publication Year，不能单凭结果页年份断言正式出版年；导出后仍须检查 `PY`、`PD`、卷期及正式出版信息。
- 当前业务口径没有进一步限定 WOS 文献类型。若以后只统计 Article 或 Review，应先确认口径，再增加 `AND DT=(Article OR Review)`；不要在本指南中悄悄缩小成果范围。

首次使用时不要只运行总式。应在相同数据库范围和年份条件下分别比较：

```text
AD=(Tsinghua Univ* SAME State Key Lab* SAME Tribol*) AND PY=2026
```

```text
AD=(Tsinghua Univ* SAME State Key Lab* SAME Tribology) AND PY=2026
```

第二式仅用于显示“未截词会漏掉多少记录”，不能作为正式检索式。抽查两式的差集，确认新增记录的地址确实属于目标实验室；若差集中出现系统性误命中，再增加地址特征，而不是直接删除 `*`。

将 `2026` 替换为目标正式出版年。跨多年检索可使用平台支持的年份范围，或按年分别检索以便核对数量。

### 3.2 导出

1. 选择 **Export → Tab delimited file**。
2. `Record Content` 选择 **Full Record**；确认导出字段至少包含 `TI`、`SO`、`VL`、`IS`、`AU`、`DI`、`PY` 和 `PD`。
3. 按当前界面允许的最大批次导出。若结果超过单批上限，合并为一个仅保留一次表头的制表符文件，并核对合并前后记录数。
4. 保存为 `raw_data/WOS.txt`。不要同时保留另一个待导入批次并期待系统自动扫描；当前解析器只读取约定文件名。

## 4. EI Compendex（Engineering Village）

### 4.1 候选检索式（必须在当前账号实测）

在 **Engineering Village → Compendex → Expert Search** 中运行：

```text
(((2025-2026) WN YR AND ((State Key Lab* NEAR/5 Tribol*) OR SKLT) WN AF AND (Tsinghua OR 100084) WN AF) AND ({ja} WN DT))
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

### 5.1 候选检索配置与专业检索式（必须在当前账号实测）

在 **中国知网 (CNKI) → 高级检索 / 专业检索** 中，数据库范围选择 **学术期刊**，期刊来源类别勾选 **【核心期刊】（北大核心）** 与 **【CSCD】**：

- **专业检索式**：
  ```text
  AF % '清华大学' AND (AF % '摩擦' + '高端装备')
  ```
  *(说明：`AF` 为作者单位；`%` 表示模糊匹配/包含；`+` 表示逻辑或 OR。该式匹配单位包含“清华大学”，且包含“摩擦”（摩擦学国家重点实验室）或“高端装备”（高端装备界面科学与技术全国重点实验室）。)*

- **高级检索界面条件配置（等效形式）**：

| 逻辑 | 字段 | 匹配方式 | 值 |
| --- | --- | --- | --- |
| 第一行 | 作者单位 | 包含/精确 | `高端装备界面科学与技术全国重点实验室` |
| 或者 | 作者单位 | 包含/精确 | `摩擦学国家重点实验室` |
| 并且 | 发表年度 | 精确 | 目标年份（如 `2025-2026`） |

重要说明：
- 期刊来源类别务必选择 **核心期刊（北大核心）** 与 **CSCD** 的并集。
- 发表年度根据需要限定目标统计年份（如 `2025-2026`）。

### 5.2 导出

1. 选择 **导出与分析 → 自定义导出**（或当前界面的 Excel 导出入口）。
2. 至少导出：题名、作者、文献来源、年、卷、期、DOI、作者单位；页码、基金项目可一并保留。
3. 确认表头能对应 `Title-题名`、`Source-文献来源`、`Volume-卷`、`Period-期`、`Author-作者`、`DOI-DOI`、`Year-年`。
4. 保存为 `raw_data/CNKI.xls` 或 `raw_data/CNKI.xlsx`。若分两次检索，合并时只保留一次表头，并核对合并记录数。

## 6. Scopus（补充核查，不是当前自动输入）

Scopus 可用于查漏或交叉核对。不要使用原文档中的 `AF-ID("State Key Laboratory of Tribology" 60021634)`：该 ID 尚未在当前 Scopus 机构选择器中核验，名称与 ID 的组合不能靠猜测固化。

不依赖机构 ID 的候选式如下；必须先做逐分支计数和地址抽查，不能仅因语法可执行就视为检索有效：

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
- [Web of Science Core Collection：Address 字段、缩写与机构名规则](https://webofscience.help.clarivate.com/en-us/Content/wos-core-collection/woscc-search-fields.htm)
- [Web of Science：`SAME` 运算符](https://webofscience.help.clarivate.com/Content/search-operators.html)
- [Web of Science：Search Rules](https://webofscience.help.clarivate.com/en-us/Content/search-rules.htm)
- [Scopus Support：Advanced Search 字段代码与运算规则](https://www.elsevier.support/scopus/answer/how-can-i-best-use-the-advanced-search)
- [清华大学机械工程系：State Key Laboratory of Tribology in Advanced Equipment](https://www.me.tsinghua.edu.cn/en/info/1249/1874.htm)
- [实验室官网：现名、历史名称与 SKLT 缩写](https://sklt.tsinghua.edu.cn/info/1157/2468.htm)
