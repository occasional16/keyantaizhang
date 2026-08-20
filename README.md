# 📊 Research Group Scholarly Output & Ledger Hub (keyantaizhang)

> **Intelligent Academic Output Ledger & Multi-dimensional Statistics Hub for Research Groups and PIs.**  
> *"Clarify research assets in one ledger; generate standard reports in one click."*

Chinese documentation: [README.zh-CN.md](README.zh-CN.md)

---

## 📁 1. Project Directory Structure

```text
📁 keyantaizhang/
│
├── 📄 PROJECT.md                 # [Project Context] Scope, constraints, validation, and release info
├── 📄 AGENTS.md                  # [Governance] Agent workflow and authorization guidelines
├── 📄 CHANGELOG.md               # [Changelog] Version history following Keep a Changelog
├── 📄 .project-template.json     # [Manifest] Template tracking manifest (v2)
│
├── 📂 config/                    # [Profiles & Dicts]
│   ├── teachers_profile.xlsx     # Member roster (names, teams, directions, alias database)
│   └── journal_if.xlsx           # Journal Impact Factor / Quartile dictionary (reserved)
│
├── 📂 raw_data/                  # [Raw Input Layer] (Official database exports, read-only)
│   ├── WOS.txt / WOS.xlsx        # Web of Science exports (Tab-delimited .txt preferred)
│   ├── EI.xlsx / EI.xls          # EI Compendex exports (Detailed record Excel preferred)
│   ├── CNKI.xls / CNKI.xlsx      # CNKI exports (Custom Excel format)
│   └── scopus.csv                # Scopus exports (Optional)
│
├── 🎮 console_dashboard.xlsm     # [Interactive Dashboard] Control panel & status board
├── 🏆 papers_final_merged.xlsx   # [Output Ledger] Sheet1: Claimed papers | Sheet2: Excluded papers
│
├── 📂 vba_modules/               # [VBA Core Engine] (Layered & Field-Oriented GBK source modules)
│   ├── Mod_Sync.bas              # Independent hot-sync bootloader (dynamic auto-scan)
│   ├── Mod0_ControlPanel.bas     # Control panel UI, button callbacks and workflow orchestrator
│   ├── Mod0_MetricsEngine.bas    # 100% dynamic row-by-row scanner & status board engine
│   ├── Mod1_TeacherPinyin.bas    # Pinyin variant and alias feature generator
│   ├── Mod2_PipelineMain.bas     # Pipeline coordinator, deduplication & delivery workbook generator
│   ├── Mod2_IngestSources.bas    # Multi-source parsers (WOS / EI / CNKI file readers)
│   ├── Mod3_Field_Author.bas     # Author field: alias loading, affiliation tag strip & claiming
│   ├── Mod3_Field_JournalIF.bas  # Journal/IF field: Title Case format & high-speed JIF lookup
│   ├── Mod3_Field_Date.bas       # Date field: publication year & date range filter
│   └── Mod3_Field_Deduplication.bas # Deduplication field: title cleaning & normalized hash keys
│
└── 📂 docs/                      # [Specifications & Dev Docs]
    ├── PIPELINE_SPEC.md          # Multi-source field mapping and cleaning rules
    ├── DATABASE_RETRIEVAL_GUIDE.md # Database search & export practical SOP
    ├── release.md                # Release gates and verification guide
    └── dev/                      # Development documents and decision records
        └── README.md             # Dev doc lifecycle and review rules
```

---

## 🚀 2. Quickstart (3-Step Workflow)

| Step | Action | Description |
| :---: | :--- | :--- |
| **1** | **Prepare Data** | Place raw export files (WOS, EI, CNKI) in [`raw_data/`](raw_data/) and populate member names in [`config/teachers_profile.xlsx`](config/teachers_profile.xlsx). |
| **2** | **Set Year Range** | Open [`console_dashboard.xlsm`](console_dashboard.xlsm) and configure the target publication date range (defaults to current full year). |
| **3** | **One-Click Run** | Click the orange button **`[ >>> 一键自动化执行全流程 <<< ]`** to automatically perform date filtering, disambiguation, deduplication, and output [`papers_final_merged.xlsx`](papers_final_merged.xlsx). |

---

## 💡 3. Core Engine Pipeline

```text
[ Raw Data: raw_data/ ] ---> (1. Official Publication Year Filter) ---> [ In-Period Records ]
                                                                               │
[ Member Roster: config/ ] -> (2. Pinyin Disambiguation & Claiming) <----------┤ (Non-members to Sheet 2)
                                                                               │
                                                                               ▼
[ Cross-DB Deduplication ] <- (3. DOI + Clean Title Double-Key Hash) <--------- [ Claimed Papers ]
       │
       ▼
[ Output Ledger: papers_final_merged.xlsx ]
  ├── Sheet 1: [ Claimed Group Papers ] (8 standard columns + SCI/EI composite tags + Impact Factor)
  └── Sheet 2: [ Excluded / Unclaimed Papers ] (Full raw author info for trace & audit + Impact Factor)
```

1. **Member Disambiguation**: Matches paper authors against all pinyin permutations and abbreviations of group members.
2. **Official Publication Year**: Strictly filters by `PY` / `Publication year` / `Year-年` to eliminate online-first cross-year discrepancies.
3. **Double-Key Deduplication**: Primary key `DOI` + secondary key `Clean Title Hash` to merge duplicates and label composite indexed items (e.g. `SCI+EI`).
4. **Impact Factor Matching**: Automatically queries `config/journal_if.xlsx` to populate Journal Impact Factor (JIF) values.
5. **100% Dynamic Statistics**: Real-time row-by-row scanning of actual results on dashboard; zero historical estimation or hardcoding.

---

## 📊 4. Standard 8-Column Output Specifications

The generated [`papers_final_merged.xlsx`](papers_final_merged.xlsx) contains 8 standard columns:

| Column | Field Name | Description |
| :---: | :--- | :--- |
| **A** | **序号 (No.)** | Auto-incremented sequence number (1, 2, 3...) |
| **B** | **论文题目 (Title)** | Capitalized title format without trailing period |
| **C** | **期刊名称 (Journal)** | Standard Title Case format preserving domain acronyms (e.g., IEEE, ASME) |
| **D** | **卷 (Volume)** | Clean Volume number |
| **E** | **期 (Issue)** | Clean Issue number |
| **F** | **作者 (Authors)** | Sheet 1: Group members only (separated by `; `). Sheet 2: Full raw authors. |
| **G** | **收录类型 (Index)** | `SCI`, `EI`, `中文核心`, or composite `SCI+EI` |
| **H** | **影响因子 (Impact Factor)** | Matched Journal Impact Factor from `config/journal_if.xlsx` |

---

## 📖 5. Project Governance & Specifications

- Project-owned context and technical constraints: [`PROJECT.md`](PROJECT.md)
- Agent development and authorization guidelines: [`AGENTS.md`](AGENTS.md)
- Complete pipeline field mapping matrix: [`docs/PIPELINE_SPEC.md`](docs/PIPELINE_SPEC.md)
- Project version history: [`CHANGELOG.md`](CHANGELOG.md)
- Release verification checklist: [`docs/release.md`](docs/release.md)
- Development workspace docs: [`docs/dev/`](docs/dev/)
