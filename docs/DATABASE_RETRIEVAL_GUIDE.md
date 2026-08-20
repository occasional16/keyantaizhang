# 📚 科研台账 - 数据库高精度检索与导出实操指南 (DATABASE_RETRIEVAL_GUIDE)

> **目标**：指导课题组老师、科研秘书及研究生，在 Web of Science (WOS)、EI Compendex、中国知网 (CNKI) 及 Scopus 官方平台上，以最高“查全率”与“查准率”批量检索本实验室师生学术成果，并精准导出为本系统支持的标准格式。

---

## 🏛️ 一、 实验室中英文官方名称与检索特征库

在各数据库检索机构（Affiliation / Address）时，请使用以下标准别名库进行覆盖：

| 语言 | 机构类型 | 官方标准全称与变体检索词 |
| :--- | :--- | :--- |
| **中文** | **重组后全称 (现名)** | `高端装备界面科学与技术全国重点实验室` |
| | **历史全称 (原名)** | `摩擦学国家重点实验室` |
| | **院系与依托单位** | `清华大学机械工程系` 或 `清华大学机械系` 或 `清华大学` |
| **英文** | **重组后英文 (现名)** | `State Key Laboratory of Tribology in Advanced Equipment` |
| | **历史英文 (原名)** | `State Key Laboratory of Tribology` |
| | **官方缩写** | `SKLT` |
| | **院系与依托单位** | `Department of Mechanical Engineering, Tsinghua University` 或 `Tsinghua University` |

---

## 🔍 二、 各大数据库检索与导出标准操作 SOP

---

### 1. Web of Science (WOS - Core Collection)

#### 1.1 检索策略与推荐检索式
在 WOS 核心合集（推荐限定为 `Science Citation Index Expanded (SCI-EXPANDED)`）中，进入 **“高级检索 (Advanced Search)”**：

- **策略 A（查全推荐：全量教师检索，本地消歧去重）**：
  直接按 `config/teachers_profile.xlsx` 中的教师姓名进行作者检索，配合清华大学机构限定：
  ```text
  OG=(Tsinghua University) AND AU=("Lu, Xinchun" OR "Lu, XC" OR "Zhang, Jianfu" OR "Feng, Pingfa" OR ...)
  ```
- **策略 B（实验室精准检索）**：
  ```text
  (AD=("State Key Laboratory of Tribology*" OR "SKLT" OR "Tribology in Advanced Equipment") OR (AD=("Tsinghua Univ*" SAME "Mech* Eng*"))) AND AU=(教师英文姓名列表...)
  ```
- **年份限定**：在检索页面或结果侧边栏勾选出版年份（如 `2026`）。

#### 1.2 导出规范（至关重要）
1. 检索结果页面点击顶部 **【导出 (Export)】 $\rightarrow$ 【纯文本文件 (Plain text file)】** 或 **【制表符分隔文件 (Tab delimited file)】**（🥇首选 Tab delimited file）；
2. **记录内容 (Record Content)**：务必下拉选择 **`全部记录 (Full Record)`**（必须包含 DOI、作者、出版物名称、出版年份等）；
3. **记录范围**：选择 `1 到 N 条记录`（每次最多导出 1000 条）；
4. **命名与放置**：将导出的 `.txt` 文件重命名为 **`WOS.txt`**，直接放入项目根目录下的 [`raw_data/`](../raw_data/) 目录中。

---

### 2. EI Compendex (Engineering Village)

#### 2.1 检索策略与推荐检索式
进入 Engineering Village 平台，切换至 **Compendex** 数据库，进入 **“Expert Search”**：

- **推荐检索式**：
  ```text
  (({State Key Laboratory of Tribology} OR {SKLT} OR {Tribology in Advanced Equipment} OR ({Tsinghua} WN AF AND {Mechanical} WN AF)) WN AF) AND ({Lu, Xinchun} OR {Zhang, Jianfu} OR {Feng, Pingfa} ...) WN AU
  ```
- **出版年份**：在侧边栏限定目标年份（如 `2026`）。

#### 2.2 导出规范
1. 勾选需要导出的全部记录（或在顶部勾选 **Select All**）；
2. 点击顶部 **【Download / Export】** 按钮；
3. **Format**：选择 **`Excel`**；
4. **Record content**：选择 **`Detailed record`（详细记录）**；
5. **勾选项**：务必勾选 **`include columns without data`**；
6. **命名与放置**：将导出的 `.xlsx` 文件重命名为 **`EI.xlsx`**，放入 [`raw_data/`](../raw_data/) 目录中。

---

### 3. 中国知网 (CNKI - 官方平台)

#### 3.1 检索策略与范围限定
进入中国知网首页，点击 **“高级检索”**：

- **文献分类**：勾选 **“学术期刊”**；
- **核心收录范围限定（重要）**：
  - 在期刊来源类别中，勾选 **【全部期刊】** 或前置限定勾选 **【核心期刊】**（北大核心）、**【CSSCI】**、**【CSCD】**；
- **检索条件配置**：
  - **作者**：填入老师姓名（多位老师用 `+` 或 `OR` 连接，如 `路新春 + 张建富 + 冯平法`）；
  - **作者单位**：`高端装备界面科学与技术全国重点实验室 + 摩擦学国家重点实验室 + 清华大学机械工程系 + 清华大学机械系`；
  - **发表年度**：精确限定为目标统计年份（如 `2026 至 2026`）。

#### 3.2 导出规范
1. 在检索结果列表点击 **“全选”**；
2. 点击 **【导出与分析】 $\rightarrow$ 【自定义导出 (Custom Export)】** 或 **【导出文献】 $\rightarrow$ 【Excel】**；
3. **必选字段**：题名、作者、文献来源（期刊名称）、年、卷、期、页码、DOI、作者单位、基金项目；
4. **命名与放置**：导出的 Excel 文件重命名为 **`CNKI.xls`** 或 **`CNKI.xlsx`**，放入 [`raw_data/`](../raw_data/) 目录中。

---

### 4. Scopus (Elsevier)

#### 4.1 检索策略与检索式
进入 Scopus 数据库高级检索（Advanced Search）：

- **推荐检索式**：
  ```text
  (AF-ID("State Key Laboratory of Tribology" 60021634) OR AFFIL("Tribology in Advanced Equipment") OR AFFIL("State Key Laboratory of Tribology") OR (AFFIL("Tsinghua University") AND AFFIL("Mechanical Engineering"))) AND AUTH(教师姓名列表...) AND PUBYEAR = 2026
  ```

#### 4.2 导出规范
1. 点击 **【Export】**；
2. 格式选择 **`CSV`**；
3. 勾选 **Citation information**（Author, Document title, Year, Source title, Volume, Issue, DOI）与 **Bibliographical information**（Affiliations 等）；
4. **命名与放置**：将导出的 `.csv` 文件命名为 **`scopus.csv`**，放入 [`raw_data/`](../raw_data/) 目录中。

---

## 📊 三、 查全与查准黄金法则

```text
┌────────────────────────────────────────────────────────┐
│  数据库检索端 (查全优先)                                │
│  • 覆盖所有历史实验室别名 (现名/原名/机械系/清华大学)    │
│  • 包含全量师生候选名                                  │
└───────────────────────────┬────────────────────────────┘
                            │ 导出原始文件放入 raw_data/
                            ▼
┌────────────────────────────────────────────────────────┐
│  科研台账系统本地端 (查准与精准消歧)                    │
│  • 1. 严格出版年份校验 (排除网络预发表跨年)            │
│  • 2. 课题组师生档案库深度消歧认领 (入库 vs 排除)       │
│  • 3. DOI + 纯净题目跨库双重去重 (合并 SCI+EI)          │
│  • 4. 直出双工作表台账 (入库大表 + 排除明细带溯源)       │
└────────────────────────────────────────────────────────┘
```

1. **宁多勿漏**：在官方数据库检索时，应采用“宽进严出”策略，尽量扩大检索词覆盖面（如包含清华大学机械工程系）；
2. **本地全自动保底**：多余或非本组人员成果，系统底层算法会自动安全分流至 Sheet 2 (`【未认领排除成果】`)，绝不污染正式台账。
