# 📊 科研台账 (keyantaizhang)

> **课题组学术成果智能整理与多维统计工作台 (Research Group Scholarly Output & Ledger Hub)**  
> **官方标语**：*“课题组学术家底，一账摸清、一键直出。”*  
> 专为高校与科研院所 PI（课题组长/导师）、学术秘书及研究生打造，告别散乱手工 Excel，一键把 WOS、EI、知网等原始记录汇聚为课题组专属的“成果总台账”！

English documentation: [README.md](README.md)

---

## 📁 1. 项目目录结构一览

```text
📁 科研台账系统根目录/
│
├── 📄 PROJECT.md                 # 【项目专属依据】产品边界、兼容范围、技术约束与验证清单
├── 📄 AGENTS.md                  # 【开发与授权规范】通用 Agent 协作与开发行为准则
├── 📄 CHANGELOG.md               # 【版本变更日志】遵循 Keep a Changelog 规范
├── 📄 .project-template.json     # 【工程元数据】模板管理与哈希校验清单
│
├── 📂 config/                    # 【人员台账与字典层】
│   ├── teachers_profile.xlsx     # 课题组师生档案库（姓名、团队、方向、全格式拼音别名库）
│   └── journal_if.xlsx           # 期刊影响因子(IF)与分区字典（预留拓展）
│
├── 📂 raw_data/                  # 【原始成果数据层】（各大数据库导出的原始文件直接放入，只读）
│   ├── WOS.txt / WOS.xlsx        # Web of Science (WOS) 官方导出 (🥇首选 Tab delimited 制表符文本)
│   ├── EI.xlsx / EI.xls          # EI Compendex 官方导出 (🥇首选 Excel 详细记录)
│   ├── CNKI.xls / CNKI.xlsx      # 中国知网 (CNKI) 官方导出 (🥇首选 自定义Excel)
│   └── scopus.csv                # Scopus 官方导出 (备用)
│
├── 🎮 console_dashboard.xlsm     # 【⭐ 科研台账交互工作台】（日常单工作表全局总控面板）
├── 🏆 papers_final_merged.xlsx   # 【⭐ 成果交付大表】（Sheet1入库成果 + Sheet2未认领排除成果）
│
├── 📂 vba_modules/               # 【VBA 源码与算法引擎】（底层纯文本源码，GBK编码受控）
│   ├── Mod_Sync.bas              # 【⭐ 热更底座】专职秒级同步本地最新代码并重置面板
│   ├── Mod0_ControlPanel.bas     # 控制台交互、状态看板、操作指引与业务调度引擎
│   ├── Mod1_TeacherPinyin.bas    # 师生多格式英文拼音与别名特征构建子引擎
│   └── Mod2_CleanRawData.bas     # 多源抽取、时间过滤、消歧认领、跨库去重与直出子引擎
│
└── 📂 docs/                      # 【详细规范与开发文档】
    ├── PIPELINE_SPEC.md          # 全流程业务与跨数据库字段映射规范说明书
    ├── DATABASE_RETRIEVAL_GUIDE.md # 各学术数据库高精度检索与导出实操指南
    ├── release.md                # 发布流程与门禁规范
    └── dev/                      # 开发工作文档与决策记录
        └── README.md             # 工作文档生命周期与管理规范
```

---

## 🚀 2. 极简操作流程与系统工作原理

### 🌟 极简操作流程（3 步搞定）

| 步骤 | 操作动作 | 详细说明 |
| :---: | :--- | :--- |
| **第 1 步** | **准备基础数据** | 将从 WOS、EI、知网导出的原始文件直接放入 [`raw_data/`](raw_data/)，在 [`config/teachers_profile.xlsx`](config/teachers_profile.xlsx) 填入课题组师生姓名。 |
| **第 2 步** | **设定统计时间** | 双击打开 [`console_dashboard.xlsm`](console_dashboard.xlsm)，在面板中直接填入或点击胶囊按钮设定年份范围（默认填充当年 `2026-01-01` ~ `2026-12-31`）。 |
| **第 3 步** | **一键生成台账** | 点击右上角橙色大按钮 **`[ >>> 一键自动化执行全流程 <<< ]`**，系统一次性自动完成消歧认领、跨库去重与复合标记，在根目录直接直出交付大表 [`papers_final_merged.xlsx`](papers_final_merged.xlsx)！ |

---

### 💡 系统底层三大核心工作原理

```text
[ 原始成果数据 raw_data/ ] ---> (1. 正式出版年份校验) ---> [ 在期成果流 ]
                                                              │
[ 师生档案库 config/    ] ---> (2. 拼音特征消歧认领) <--------┤ (非本组条目安全归入排除表)
                                                              │
                                                              ▼
[ 跨库去重与复合标记   ] <--- (3. DOI+题目哈希双重去重) <---- [ 本组在期成果 ]
       │
       ▼
[ 直出交付大表 papers_final_merged.xlsx ] 
  ├── Sheet 1: 【课题组入库成果】 (支持 SCI+EI 复合收录标注)
  └── Sheet 2: 【未认领排除成果】 (含原库作者全文，便于溯源排查)
```

1. **智能消歧与安全认领**：
   - 基于课题组固定人员的全格式英文拼音、缩写变体与检索特征库，对论文作者进行深度模糊匹配；
   - **严格准入**：仅保留本组人员为作者的成果直出终稿，整篇无本组师生的条目自动安全排除至 Sheet 2，并在控制台单行备注排除明细。
2. **正式出版年份精准过滤**：
   - 严格依据官方正式出版年份（`PY` / `Publication year` / `Year-年`）为过滤准绳，排除网络预发表跨年干扰，确保年度统计数据的权威性与严谨性。
3. **跨库去重与复合收录合流**：
   - 采用 **`DOI` 主键 + `纯净题目哈希` 备用键** 双重消歧算法，完美消除 WOS、EI、知网之间的重复收录；
   - 跨库重复成果自动合并并标记为 **`SCI+EI`** 复合收录标签，字段智能互补。
4. **100% 真实动态扫描统计**：
   - 控制台刷新全部基于产出结果逐行实测统计，彻底杜绝任何比例推算或数量硬编码。

---

## 📊 3. 课题组 7 大核心台账字段规范

打开生成的 [`papers_final_merged.xlsx`](papers_final_merged.xlsx)，包含完整的 7 大标准字段：

| 列号 | 字段名称 | 规范与示例说明 |
| :---: | :--- | :--- |
| **A** | **序号** | 从 1 开始自增纯数字编号（1, 2, 3...） |
| **B** | **论文题目** | 英文首字母大写规范，末尾无多余句点（如 *Aerodynamic drag study of speed skaters*） |
| **C** | **期刊名称** | 严格 **Title Case** 英文标题大小写规范（如 *Journal of Manufacturing Processes*） |
| **D** | **卷** | 纯净卷号（Volume，如 *133*） |
| **E** | **期** | 纯净期号（Issue，如 *4*） |
| **F** | **作者** | 入库表仅保留属于本课题组人员姓名（多位成员用分号 `; ` 间隔，如 *Zhang, Jianfu; Feng, Pingfa*）；排除表保留原库作者全文。 |
| **G** | **收录类型** | `SCI`、`EI`、`中文核心`，跨库双收录成果精准标注为 **`SCI+EI`** |

---

## 🛠️ 4. 原始数据最高精度导出规范

为了确保零错列与极速解析，推荐按以下格式导出数据放入 [`raw_data/`](raw_data/)：
- **WOS (SCI)**：选择 **`Tab delimited file`（制表符分隔文件）**，记录内容勾选 **`Full Record（全记录）`**，命名为 [`raw_data/WOS.txt`](raw_data/WOS.txt)；
- **EI Compendex**：选择 **`Excel`**，记录内容选择 **`Detailed record`**，勾选 **`include columns without data`**，命名为 [`raw_data/EI.xlsx`](raw_data/EI.xlsx)；
- **知网 (CNKI)**：自定义导出 Excel，命名为 [`raw_data/CNKI.xls`](raw_data/CNKI.xls)。

---

## ❓ 5. 常见问题与答疑 (FAQ)

### Q1：打开 Excel 提示“被阻止的内容 / 宏无法运行”怎么办？
- **永久彻底解决方法**：
  1. 在 Excel 点击 **【文件】 $\rightarrow$ 【选项】 $\rightarrow$ 【信任中心】 $\rightarrow$ 【信任中心设置】**；
  2. 点击左侧 **【受信任位置】 $\rightarrow$ 【添加新位置...】**；
  3. 浏览选中当前项目根目录，**勾选【同时信任此位置的子文件夹】**，点击确定即可永久关闭拦截！
  4. 在【宏设置】中勾选 **【信任对 VBA 工程对象模型的访问】**（以支持代码热更新）。

### Q2：课题组新增了老师或硕博研究生，怎么让他生效？
- 直接打开 [`config/teachers_profile.xlsx`](config/teachers_profile.xlsx)，在表格最后一行填入新成员的 **中文姓名** 和团队方向并保存；
- 点击控制台右上角 **`[ >>> 一键自动化执行全流程 <<< ]`**，系统会自动为新增成员构建拼音库并全量重新归属匹配！

### Q3：换了一台新电脑或者把文件夹拷走了，还能用吗？
- **100% 可以！** 本系统底层全部采用动态相对路径（`ThisWorkbook.Path`），不绑定任何固定盘符或绝对路径，随处移动均能正常运行。

---

## 📖 6. 详细技术规范与工程治理
- 项目专有上下文与技术约束：[`PROJECT.md`](PROJECT.md)
- Agent 协作与开发规范：[`AGENTS.md`](AGENTS.md)
- 跨数据库字段映射与清洗规则对照表：[`docs/PIPELINE_SPEC.md`](docs/PIPELINE_SPEC.md)
- 项目变更日志：[`CHANGELOG.md`](CHANGELOG.md)
- 版本发布指南：[`docs/release.md`](docs/release.md)
- 开发工作文档与决策记录：[`docs/dev/`](docs/dev/)
