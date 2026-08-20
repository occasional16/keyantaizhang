# 科研台账 (keyantaizhang) - 全流程与跨数据库字段映射规范说明书 (PIPELINE_SPEC)

> **产品全称**：科研台账 (keyantaizhang)  
> **核心定位**：课题组学术成果智能整理与多维统计工作台 (Research Group Scholarly Output & Ledger Hub)  
> **版本**：v0.1.0 (双工作表台账直出与 100% 动态实测统计版)  
> **更新时间**：2026-08-20  
> **核心原则**：
> 1. **全相对路径原则**：所有 VBA 与脚本严禁硬编码绝对路径，必须通过 `ThisWorkbook.Path` 动态获取根目录，保证项目随处拷贝移动 100% 正常运行。
> 2. **双核解耦架构**：采用独立的 [`Mod_Sync.bas`](../vba_modules/Mod_Sync.bas) 作为底座负责秒级热载入，业务全流程调度在 [`Mod0_ControlPanel.bas`](../vba_modules/Mod0_ControlPanel.bas)，支持随时无损更新。
> 3. **双工作表成果直出 (`papers_final_merged.xlsx`)**：
>    - **Sheet 1: `【课题组入库成果】`**：标准包含 8 大核心字段的在期本组正式成果表（含影响因子）；
>    - **Sheet 2: `【未认领排除成果】`**：包含 8 大核心字段的在期非本组/未认领论文明细（含原库作者及影响因子），便于溯源核验与查漏补缺。
> 4. **期刊影响因子智能匹配**：依据 `config/journal_if.xlsx` 的 `Journal Impact Factor` 列进行期刊名称归一化匹配，自动填充最新影响因子。
> 5. **100% 真实动态统计（零硬编码、零比例推算）**：看板统计数据直接逐行真实统计 Sheet 1 与 Sheet 2 的全部条目，彻底杜绝任何估算。
> 6. **正式出版年份日期过滤**：严格以 `PY` / `Publication year` / `Year-年` 为准，提供单行极简交互配置，默认填充当年 2026 全年度。

---

## 一、 系统顶层架构与目录设计

```text
📁 科研台账系统根目录/              # 项目根目录（动态自适应相对路径 ThisWorkbook.Path）
│
├── 📄 PROJECT.md                   # 【项目专属依据】产品边界、兼容范围、技术约束与验证清单
├── 📄 AGENTS.md                    # 【开发与授权规范】通用 Agent 协作与开发行为准则
├── 📄 CHANGELOG.md                 # 【版本变更日志】遵循 Keep a Changelog 规范
├── 📄 .project-template.json       # 【工程元数据】模板管理与哈希校验清单
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
├── 🏆 papers_final_merged.xlsx     # 成果交付大表（Sheet1入库成果 + Sheet2未认领排除成果）
│
├── 📂 vba_modules/                 # 【核心 VBA 源代码模块】（纯文本 .bas，GBK编码版本受控）
│   ├── Mod_Sync.bas                # 【⭐ 独立热更底座】动态扫描并秒级无损同步所有模块
│   ├── Mod0_ControlPanel.bas       # 控制台交互、按钮回调、业务全流程调度与交付报告引擎
│   ├── Mod0_MetricsEngine.bas      # 100% 真实动态逐行扫描、状态看板刷新引擎
│   ├── Mod1_TeacherPinyin.bas      # 师生名单解析、全格式拼音与检索特征库构建引擎
│   ├── Mod2_PipelineMain.bas       # 多源清洗总调度、记录聚合去重、双工作表 8 列直出
│   ├── Mod2_IngestSources.bas      # 多源原始文献解析器 (WOS / EI / CNKI 官方文件高精度抽取)
│   ├── Mod3_Field_Author.bas       # 【字段层 - 作者】别名库加载、机构角标剥离与消歧认领
│   ├── Mod3_Field_JournalIF.bas    # 【字段层 - 期刊/IF】Title Case 规范化与 JIF 字典高速匹配
│   ├── Mod3_Field_Date.bas         # 【字段层 - 时间】出版年与日期跨度解析、区间有效性判定
│   └── Mod3_Field_Deduplication.bas# 【字段层 - 去重键】题目清洗、归一化字符提取与主备键生成
│
├── 📂 docs/                        # 【系统详细规范与开发文档】
│   ├── PIPELINE_SPEC.md            # 全流程业务与跨数据库字段映射规范说明书（本文件）
│   ├── DATABASE_RETRIEVAL_GUIDE.md # 各学术数据库高精度检索与导出实操指南
│   ├── release.md                  # 发布流程与门禁规范
│   └── dev/                        # 研发工作文档与决策记录
│       └── README.md               # 工作文档生命周期与管理规范
│
│
├── 📂 .vscode/                     # VS Code 编辑器配置（自动识别 GBK 编码）
├── 📄 README.md                    # 英文说明与快速入门
├── 📄 README.zh-CN.md              # 详尽中文使用说明手册
└── 📄 .gitignore                   # Git 排除规则（忽略数据文件，追踪源码文档）
```

---

## 二、 跨数据库多源字段全景对照映射表 (Cross-Database Field Matrix)

本系统支持同时接入 **Web of Science (WOS)**、**EI Compendex**、**中国知网 (CNKI)** 及 **Scopus** 等多源数据，各数据库字段映射及提取对齐标准如下：

| 交付终稿字段<br>(`papers_final_merged.xlsx`) | Web of Science (WOS)<br>(`WOS.txt` / `WOS.xlsx`) | EI Compendex<br>(`EI.xlsx` / `EI.xls` / `EI.csv`) | 中国知网 (CNKI)<br>(`CNKI.xls` / `CNKI.xlsx`) | Scopus (Elsevier)<br>(`scopus.csv`) | 核心清洗与提取转换规则 |
| :---: | :---: | :---: | :---: | :---: | :--- |
| **序号** (Col A) | *(自动生成)* | *(自动生成)* | *(自动生成)* | *(自动生成)* | 区分入库与排除后，自增纯数字编号（1, 2, 3...）。 |
| **论文题目** (Col B) | `TI`<br>`Article Title` | `Title` | `Title-题名`<br>`题名` | `Title` | • 压缩连续多余空格与换行符；<br>• 去除题目末尾多余的英文句点 `.`。 |
| **期刊名称** (Col C) | `SO`<br>`Source Title` | `Source` | `Source-文献来源`<br>`文献来源` | `Source title` | • 英文期刊统一规范为 **Title Case**；<br>• 保护专用缩略词大写（如 IEEE/ASME/CFD）；<br>• 虚词/介词/连词小写（如 of/in/and/for）。 |
| **卷** (Col D) | `VL`<br>`Volume` | `Volume` | `Volume-卷`<br>`卷` | `Volume` | 纯净卷号提取，若原库为空则保留为空。 |
| **期** (Col E) | `IS`<br>`Issue` | `Issue` | `Period-期`<br>`期` | `Issue` | 纯净期号提取，若原库为空则保留为空。 |
| **作者** (Col F) | `AU`<br>`Authors` | `Author` | `Author-作者`<br>`作者` | `Authors` | • **EI 机构角标剥离**：正则去除 `(1,2)` 等数字标签；<br>• **师生档案多维匹配**：根据 `teachers_profile.xlsx` 全拼音别名消歧识别；<br>• **入库表 (Sheet 1)** 仅保留本组人员姓名；<br>• **排除表 (Sheet 2)** 保留原库作者全文便于溯源。 |
| **通讯作者** (Col G) | `RP`<br>`Reprint Author` | `Corresponding author(s)` | `Supervisor-导师`<br>`导师` | `Correspondence Address` | • **仅提取本组教师**：非本组成员自动过滤不保留；<br>• **英文来源格式**：`英文名(中文名)`（如 `Tian, Yu(田煜)`）；<br>• **知网来源格式**：`中文名(导师)`（如 `邵天敏(导师)`）；<br>• **跨库去重合流**：多源合流时去重合并并优先保留全拼格式。 |
| **收录类型** (Col H) | 标记为 `SCI` | 标记为 `EI` | 标记为 `中文核心` | 标记为 `Scopus` | • 跨库合并时，自动组合为复合标签（如 **`SCI+EI`**）；<br>• 排序权重优先级：`SCI` > `EI` > `中文核心`。 |
| **影响因子** (Col I) | *(匹配 JIF 字典)* | *(匹配 JIF 字典)* | *(匹配 JIF 字典)* | *(匹配 JIF 字典)* | • 依据 `config/journal_if.xlsx` 的 `Journal Impact Factor` 列；<br>• 期刊名去除符号与空格后全大写归一化匹配；<br>• 命中则填充数值，未收录则留空。 |
| **DOI** (Col J) | `DI`<br>`DOI` | `DOI` | `DOI-DOI`<br>`URL-网址` | `DOI`<br>`Link` | • 写入 Excel 原生**蓝色下划线可点击超链接**；<br>• 点击直接打开 `https://doi.org/<DOI>`；<br>• 无标准 DOI 时自动回退填充知网 `URL-网址` 确保可点击。 |
| *(去重键: DOI)* | `DI`<br>`DOI` | `DOI` | `DOI-DOI`<br>`DOI` | `DOI` | 统一转换为全小写、去空格，作为跨库合并主键。 |
| *(去重键: 归一化题目)* | `TI` (正则纯净串) | `Title` (正则纯净串) | `Title-题名` (纯净串) | `Title` (纯净串) | 仅提取中英文字符与数字作为二级去重备用主键。 |
| *(时间过滤基准: 年份)* | **`PY`**<br>`Publication Year` | **`Publication year`** | **`Year-年`**<br>`年` | `Year` | 严格以**正式出版年份**为核心过滤基准。 |
| *(时间过滤辅助: 日期)* | `PD`<br>`Publication Date` | `Issue date` | `PubTime-发表时间` | `Date` | 辅助提取具体月份与日期（英文月份转数字）。 |

---

## 三、 各核心字段清洗与消歧规则细则

### 1. 双工作表输出与排除条目明细
- **Sheet 1 (`课题组入库成果`)**：
  - 仅输出通过消歧认领、至少包含 1 位课题组人员的在期正式成果；
- **Sheet 2 (`未认领排除成果`)**：
  - 输出在期但未认领到本组任何人员的全部排除论文（包含原始作者、收录类型与影响因子）；
  - **核心价值**：课题组老师核对时，若发现某篇论文遗漏，可直接在 Sheet 2 中查找原因（如新进学生名字未录入 `teachers_profile.xlsx`），补齐后重新一键运行即可。

### 2. 100% 真实逐行动态统计
- 控制台在刷新看板时，分别逐行扫描 `papers_final_merged.xlsx` 的 Sheet 1 与 Sheet 2；
- 真实计算每一张工作表中的 SCI、EI、中文核心、SCI+EI 复合收录篇数；
- **彻底废除任何硬编码分支或比例推算公式**。

### 3. 期刊影响因子 (JIF) 高速匹配引擎
- **字典源**：`config/journal_if.xlsx`（包含 50,000+ 来源期刊与最新影响因子）；
- **匹配机制**：加载进入内存哈希字典，将文献期刊名称（`Journal`）与字典标准名称（`Name`）统一剔除空格及所有标点符号并转大写（Alphanumeric Normalization）；
- **输出格式**：保留数值精度（如 `8.2`, `6.9`, `14.2`），未匹配或非 SCI 期刊保持为空白，支持在 Excel 中直接进行数值降序排序与统计。
