# HMFOR 实验变量抽取 Agent

用于批量读取 HMFOR / Ni-Co 相关 PDF 文献，抽取催化剂、材料、反应条件、产物分析、性能指标、证据来源和质量问题，并生成可复核、可横向比较的数据表。

## 快速使用

默认读取 `test` 文件夹，输出到 `output-test`：

```powershell
.venv\Scripts\python.exe extract_hmfor_variables.py
```

如需联网补全 DOI / Crossref 元数据和 SI 候选链接：

```powershell
.venv\Scripts\python.exe extract_hmfor_variables.py --enable-doi-online
```

如需生成 RHE 候选换算审计值：

```powershell
.venv\Scripts\python.exe extract_hmfor_variables.py --enable-potential-conversion
```

如需尝试自动发现并下载 SI 附件：

```powershell
.venv\Scripts\python.exe extract_hmfor_variables.py --enable-doi-online --enable-si-download
```

注意：联网 DOI、SI 下载和 RHE 换算都只写入审计/补全表，不覆盖主表人工校准结果。若 SI 成功下载，脚本会自动二次抽取并将本地 SI 内容写入 `si_links.csv` 和 `si_field_evidence_log.csv`。

## Python API

目标系统推荐通过薄 API 调用核心工具：

```python
from hmfor_extractor import extract

manifest = extract(
    input_dir="test",
    output_dir="output-test",
    config_path="config.yaml",
    enable_doi_online=True,
    enable_si_download=True,
)
```

详细说明见 `API.md`。

## 正式输出目录

`output-test` 是当前正式输出目录。临时烟测目录已清理。

最推荐先打开：

- `run_manifest.json`：运行清单，记录工具版本、schema 版本、输入/输出目录、配置、开关、统计和错误码。
- `output_schema.json`：机器可读输出字段说明和表行数。
- `hmfor_experiment_variables.xlsx`：总工作簿，包含主表、证据、质量、SI、DOI、RHE、单位、去重、比较准备等 sheet。
- `hmfor_experiment_variables.csv`：HMFOR 核心/相关论文主表，一篇一行。
- `manual_review_queue.csv`：人工复核队列，适合优先处理不确定项。
- `comparison_ready_metrics.csv`：横向比较和后续画图应使用的干净指标表。

## 输出表说明

- `papers.csv`：一篇一行的文献信息、分类、文本状态、References 删除状态。
- `materials.csv`：材料/催化剂信息。包含 exclude 论文；exclude 行标记 `table_scope=not_applicable_for_exclude`。
- `metrics_long_table.csv`：一个指标一行的长表，保留 HMFOR 和严格模板下的背景指标。
- `comparison_ready_metrics.csv`：通过数据闸门的可比较指标。带 `scale needs check`、reference section、very_low、背景泛句等不会进入。
- `hmfor_evidence_log.csv`：主文证据摘录。
- `si_links.csv`：本地/下载 SI 匹配状态、SI 需求、Publisher/DOI 候选 SI URL。
- `si_download_report.csv`：SI 自动发现/下载审计表，记录 candidate、HTML 发现、下载成功、403/404、非 PDF/DOCX 等状态。
- `si_field_evidence_log.csv`：本地或下载的 SI PDF/DOCX 中实际扫到的字段证据。当前 P0011 已可从 Springer ESM DOCX 中解析出部分 SI 证据。
- `doi_enrichment_queue.csv`：需要 DOI/出版社补全的任务队列。
- `doi_enrichment_results.csv`：Crossref 联网补全结果；状态包括 `success`、`failed_*`、`skipped_metadata_complete`、`skipped_offline_default`。
- `potential_conversion_audit.csv`：RHE 审计表，包含 `converted_candidate_vs_rhe`、`final_potential_vs_rhe`、`potential_trust_level`。
- `unit_normalization_log.csv`：单位标准化日志。
- `dedup_report.csv`：DOI、标题指纹、文件哈希综合去重报告。
- `quality_issues.csv`：机器可识别质量问题，含 `source_issue_id`。
- `quality_scores.csv`：文献质量子分和总分。
- `visualization_readiness_report.csv`：不同 performance_context 的画图准备状态。

## 当前边界

- 没有本地 SI 或可下载 SI 附件时，工具只能生成 SI 候选链接和复核提醒。
- SI 自动下载目前支持 PDF/DOCX 附件；出版社 403/404 或只暴露网页时会进入 `si_download_report.csv`，不会伪装成已下载。
- P04/P07/P11 仍保留 `manual_overrides.json` 人工校准，但部分字段已反向沉淀为规则；`override_fields` 只统计仍由人工修正改变的字段。
- RHE 换算候选值不等于最终结论，低置信值不会进入 `final_potential_vs_rhe`。
- `quality_issues.csv` 是机器质检全集，`manual_review_queue.csv` 是从中筛出的人工处理优先队列。

## 回归测试

```powershell
.venv\Scripts\python.exe tests\regression_smoke.py
```

测试会检查 P04/P07/P11 核心数值、P0007 不确定电位不进入比较表、P0011 conversion 缺口提示、P0012 exclude 材料表范围、manifest/schema 是否存在。
