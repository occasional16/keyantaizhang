# Changelog

All notable changes to the **科研台账 (keyantaizhang)** project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added
- **通讯作者 (Corresponding Author) 白名单智能提取与第 9 列交付**：
  - 在成果交付大表 `papers_final_merged.xlsx`（Sheet 1 与 Sheet 2）中新增第 7 列 **`通讯作者`**（Col G），整体升级为 9 列标准结构；
  - **白名单准入**：严格只保留属于课题组师生档案库（`teachers_profile.xlsx`）的通讯作者，外部非本组成员自动过滤；
  - **多源格式化**：英文来源自动格式化为 `英文名(中文名)`（如 `Tian, Yu(田煜)`），知网导师格式化为 `中文名(导师)`（如 `邵天敏(导师)`）；
  - **跨库智能去重合流**：WOS 与 EI 复合收录条目自动按教师去重合并，优先保留全拼规范格式。
- **期刊影响因子 (JIF) 自动匹配与第 9 列输出**：
  - 新增 `config/journal_if.xlsx` 高速哈希字典载入与期刊名称字母数字归一化匹配引擎；
  - 影响因子顺延至第 9 列（Col I），双工作表直出。

### Refactored
- **VBA 核心引擎分层与字段解耦模块化重构**：
  - 彻底拆分原 900 行单体大文件，按表现层、统计层、采集层、总控层与专属字段层（作者、期刊/IF、时间、去重键）解耦为 10 个独立轻量级模块；
  - `Mod_Sync.bas` 升级为动态扫描载入机制，自动枚举 `vba_modules/*.bas` 实现秒级无损热更新；
  - 彻底消除跨模块公共函数命名冲突，保障后期字段无缝横向扩展。

## [0.1.0] - 2026-08-20

### Added
- **标准化工程治理框架**：接入 `bootstrap-code-project` 规范，建立 `PROJECT.md`、`AGENTS.md`、`docs/dev/README.md`、`docs/release.md` 与 `.project-template.json`。
- **全流程科研成果自动化台账工作台**：
  - 支持 WOS、EI、知网 (CNKI) 多源文献导出数据一键抽取与对齐；
  - 严格以官方正式出版年份 (`PY` / `Publication year` / `Year-年`) 为准进行日期过滤；
  - 师生全格式拼音与英文别名多维特征匹配消歧认领；
  - DOI 主键 + 纯净题目二级备用键双重跨库去重，自动合并 `SCI+EI` 复合收录标签。
- **双工作表成果交付大表 (`papers_final_merged.xlsx`)**：
  - **Sheet 1: `【课题组入库成果】`**：标准 7 列正式成果交付大表（仅含本组成员成果）；
  - **Sheet 2: `【未认领排除成果】`**：标准 7 列排除论文明细（包含原库作者全文），便于课题组老师溯源排查。
- **100% 真实逐行动态扫描统计**：控制台看板数据全部基于实际生成大表动态逐行实测统计，零历史比例估算，零数量硬编码。
- **双核解耦与热更新底座**：`Mod_Sync.bas` 支持秒级无损代码热重载，`Mod0_ControlPanel.bas` 负责交互调度。
- **中英文双语文档与开发规范**：提供 `README.md`、`README.zh-CN.md`、`docs/PIPELINE_SPEC.md`、`docs/DATABASE_RETRIEVAL_GUIDE.md` 与开发工作区规范 `docs/dev/`。
