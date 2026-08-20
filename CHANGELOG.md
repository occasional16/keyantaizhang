# Changelog

All notable changes to the **科研台账 (keyantaizhang)** project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

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
- **中英文双语文档与开发规范**：提供 `README.md`、`README.zh-CN.md`、`docs/PIPELINE_SPEC.md` 与开发任务文档 `docs/dev/0.1-01-architecture-and-governance.md`。
