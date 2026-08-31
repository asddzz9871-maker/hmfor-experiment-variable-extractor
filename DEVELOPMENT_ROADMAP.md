# HMFOR 抽取 Agent 开发需求与优先级

更新日期：2026-08-30

## V1.9 Cleanup 收尾目标

V1 不继续加复杂功能，只做收尾清理，目标是减少明显误抽和提高质检可见性。

必修项：

1. 修短值异常：所有文本字段长度小于 20 且不是标准枚举值时，进入 `quality_issues.csv`。
2. 修 P0007 电位字段：主文出现 `onset potential of 1.32 V` 和 `50 mA/cm2 at 1.38 V`，V1.9 先写成 `scale needs check`，不完全留空。
3. 修 stability 同步：如果 `stability_performance` 有内容而 `stability` 缺失，则 `stability` 使用同一摘要。
4. 强化背景指标质检：OER/HER/UOR/电容指标若是泛泛背景句、作者/标题残片或缺少数值，标为 `very_low` 并进入 `quality_issues.csv`。
5. carbon_balance 语义化：HMFOR 核心论文未发现碳平衡时，标为 `not_reported_in_main_text`，避免看起来像抽取失败。

## V2 开发顺序

1. 自动去重。最先做，避免 DOI 补全、SI 关联和横向比较被重复论文污染。
2. Supporting Information 识别。很多关键条件、产物分析、循环稳定性和催化剂配比在 SI，应优先解决。
3. 单位标准化。统一 `mA cm-2 / mA/cm2 / mA cm−2`、`M/m`、`V vs RHE` 等表达。
4. RHE 电位自动换算。依赖 reference electrode、pH/KOH 浓度和标准化后的电位字段。
5. DOI 联网补全。用于补标题、期刊、年份和可能的 SI 链接，但不是抽取准确性的前置条件。
6. 文献质量评分。应建立在字段标准化、SI 识别和证据质量可追踪之后。
7. 横向对比雷达图/热图。最后做，等 `metrics_long_table.csv` 稳定后再进入展示层。

## V3 候选方向

与多智能体系统对接，把稳定的 Excel/CSV 结果交给实验设计 Agent，用于复现候选、变量矩阵设计和下一轮实验建议。


## V2 当前开发状态

已启动 V2 数据底座第一批：

- 已新增 `papers.csv` / `papers.xlsx`。
- 已新增 `materials.csv` / `materials.xlsx`。
- 已新增 `dedup_report.csv` / `dedup_report.xlsx`，当前基于 DOI 和标题指纹做本地去重。
- 已新增 `si_links.csv` / `si_links.xlsx`，当前基于文件名共享词识别主文与 SI。
- 已扩展 `metrics_long_table.csv` / `.xlsx`，加入标准化值和标准化单位列。

下一步建议继续做 Supporting Information 识别增强和单位标准化规则表。


## V2 可信数据底座扩展状态

已继续加入：

- `comparison_eligible` / `exclusion_reason`，用于控制哪些指标可进入横向比较。
- `field_confidence.csv`，记录字段级置信度、原因和证据强度。
- `potential_conversion_audit.csv`，记录电位换算审计，不在证据不足时强行换算。
- `unit_normalization_log.csv`，记录单位标准化痕迹。
- `manual_review_queue.csv`，提炼真正需要人工看的条目。
- `quality_scores.csv`，拆分 completeness、traceability、relevance、mechanism、reproducibility、SI support 和 overall score。

下一步建议围绕 SI 识别增强和 RHE 换算公式表继续开发。


## V2 RHE/SI 审计增强状态

已继续加入：

- `rhe_conversion_rules.json`：集中保存 Hg/HgO、Ag/AgCl、SCE 到 RHE 的候选审计公式和所需假设。
- `unit_normalization_rules.json`：集中保存单位标准化规则说明。
- `papers.csv` 增加 authors、paper_type、relevance_level、si_mentioned_in_text、si_mention_evidence。
- `si_links.csv` 增加 si_needed 和 si_needed_reason，用于区分“没找到 SI”和“疑似需要 SI”。
- `potential_conversion_audit.csv` 升级为公式审计：已是 vs RHE 则不换算；标尺不确定或参比电极未核实时只给公式和假设，不覆盖主表。

后续建议：

1. 将 RHE 换算做成严格的 opt-in 参数，只有用户确认参比电极、pH/KOH 和温度假设后才写入 converted value。
2. 增强 SI 文件识别：支持 DOI 附件名、`-si`、`suppinfo`、`mmc`、`esm`、期刊下载名等更多模式。
3. 增加真实 SI 测试样本，验证 fields_filled_from_si 是否可靠。


## V2 opt-in RHE 换算与 SI 匹配增强

已继续加入：

- 新增命令行参数 `--enable-potential-conversion`。默认关闭；开启后只在 `potential_conversion_audit.csv` 中输出估算值，不回写主表字段。
- `potential_conversion_audit.csv` 新增 `conversion_enabled`，便于区分默认审计和 opt-in 估算。
- SI 文件名识别增强，支持 `suppinfo`、`supplementary`、`supporting`、`si`、`esm`、`mmc` 等常见附件命名片段。
- SI 关联逻辑增强：除共享词外，加入 DOI/article-id 风格片段匹配。

注意：P0007 等 `scale needs check` 电位即使开启换算也只给低置信候选结果，不能直接作为主表结论。

## V2 SI 内容审计增强

已继续加入：

- `si_links.csv` 增加 `si_text_status`、`si_used_for_extraction`、`fields_filled_from_si`、`si_field_evidence_summary`，不只判断有没有 SI 文件，也记录 SI 文本中实际扫到哪些关键字段。
- SI 内容扫描覆盖 HMF 浓度、HMF conversion、FDCA yield、FDCA selectivity、FE、HPLC/NMR/GC、稳定性、碳平衡、Ni/Co 比例和金属负载。
- `manual_review_queue.csv` 增加 “SI found but no relevant fields detected” 类型，用于发现已匹配 SI 但未能贡献字段证据的情况。
- 已用本地 synthetic SI 烟测验证：可以识别 HMF 浓度、HMF conversion、FDCA yield、FE、HPLC、stability 和 carbon balance。

当前边界：

- SI 扫描目前作为审计和复核提示，不直接覆盖主表字段。
- 真实 SI 文件较少时，`si_links.csv` 更多体现“是否缺 SI / 是否疑似需要 SI”。
- 文本过短的 SI 仍会被标为 `ocr_needed`，但若能扫到明确字段，仍记录 `fields_filled_from_si` 作为辅助证据。

后续建议：

1. 加入真实 SI 样本，验证不同出版社附件文本结构。
2. 细化 `fields_filled_from_si` 到字段级来源表，和 `evidence_log.csv` 合并。
3. 扩展单位标准化规则覆盖浓度、时间、质量负载、面积归一化电流密度。
4. 细化去重策略，加入 DOI、标题指纹、文件哈希和 SI 主文关联四层报告。

## V2 数据底座收尾状态

已按既定顺序补齐 V2 剩余功能：

1. 自动去重增强
   - `dedup_report.csv` 增加文件哈希、标题指纹、匹配依据、主记录和重复原因。
   - 去重顺序：DOI -> 标题指纹 -> 文件哈希 -> 文件名兜底。

2. Supporting Information 识别与证据审计
   - `si_links.csv` 继续记录 SI 是否找到、文本状态、是否贡献字段、字段摘要。
   - 新增 `si_field_evidence_log.csv`，将 SI 字段证据拆成长表，source_scope 固定为 `supporting_information`。

3. 单位标准化
   - `unit_normalization_log.csv` 覆盖 unicode minus、mA/cm2、mA cm-2、A cm-2、cm²/cm^2、mol L-1、wt %、vs. RHE 等规范化痕迹。

4. RHE 电位换算审计
   - `potential_conversion_audit.csv` 保持默认只审计不换算。
   - `--enable-potential-conversion` 开启后只在审计表中给候选换算，不回写主表。

5. DOI 联网补全
   - `doi_enrichment_queue.csv` 默认离线生成待补全队列。
   - 新增 `--enable-doi-online`，开启后查询 Crossref，并输出 `doi_enrichment_results.csv`。
   - 联网结果只进入补全结果表，不覆盖主表抽取字段。

6. 文献质量评分
   - `quality_scores.csv` 已拆成 completeness、traceability、relevance、mechanism、reproducibility、SI support、overall。

7. 横向比较准备
   - `comparison_ready_metrics.csv` 只保留 comparison_eligible=yes、非重复主记录且数值可标准化的指标。
   - `visualization_readiness_report.csv` 汇总每类 performance_context 是否适合画热图/柱状图。

验证记录：

- 默认运行：`python extract_hmfor_variables.py`，14 篇 PDF 全量通过。
- opt-in RHE：`python extract_hmfor_variables.py --out-dir output-test-conversion --enable-potential-conversion`，已通过。
- opt-in DOI：`python extract_hmfor_variables.py --out-dir output-test-doi-online-smoke --limit 4 --enable-doi-online`，P0003/P0004 Crossref 匹配成功。

当前 V2 边界：

- SI 字段证据目前用于审计和复核，不自动覆盖主表字段。
- DOI 联网补全不自动改写 title/year/journal/doi，避免网络元数据污染人工校准结果。
- RHE 换算是候选值，必须保留 assumption_used 和 conversion_confidence。
- 热图/雷达图仍建议等 comparison_ready_metrics 稳定后再做展示层。

## V2 评审纠偏与实跑状态

根据最新评审已修正：

- DOI 补全不再描述为默认已完成。默认模式只生成队列和候选 URL；真正联网需使用 `--enable-doi-online`。
- 已对 `output-test` 实际运行 `--enable-doi-online`，当前 `doi_enrichment_results.csv` 结果为：12 条 `success`，2 条 `skipped_metadata_complete`，0 条失败。
- DOI 查询前新增尾部标点清理，修复 `10.1002/anie.202100371.` 这类 DOI 查询失败问题。
- DOI/SI 候选发现新增出版社规则：ACS、Elsevier、Wiley、Springer/Nature、RSC；输出 DOI landing、Google 检索候选和 publisher SI candidate URL。
- RHE 审计新增 `converted_candidate_vs_rhe`、`final_potential_vs_rhe`、`potential_trust_level`。低置信候选值不进入 final 字段。
- P0011 的 best sentence 问题由 `best_sentence_missing_target_metrics` 细分为 `best_sentence_missing_conversion_only`，并新增 `reported_in_figure_not_quantified_in_text` 用于表示 HMF conversion 可能在图/SI 中但正文未量化。
- 背景指标抽取改为严格模板，泛背景句和作者/标题残片不再进入 `metrics_long_table.csv`。
- 图表生成前闸门继续使用 `comparison_ready_metrics.csv`：仅允许 `comparison_eligible=yes`、`plot_ready=yes`、非重复主记录、非 reference section、非 very_low confidence 的指标进入展示层。

当前真实边界：

- SI 自动发现目前提供候选 URL 和本地 SI 匹配；还没有自动下载 SI 文件。
- `si_field_evidence_log.csv` 在当前 14 篇本地文件中仍为 0 行，因为文件夹内没有匹配到真实 SI PDF。
- RHE candidate 是审计候选，不是最终结论；`final_potential_vs_rhe` 只在高置信或原文明确 vs RHE 时填入。
- P04/P07/P11 仍高度依赖 manual overrides，V2 后续若继续提升准确率，应把 override 反向沉淀为抽取规则。

## V2 输出收口修正

根据最新输出审阅，已完成以下收口：

- 带 `scale needs check` 的指标不再进入横向比较：`metrics_long_table.csv` 中保留但 `comparison_eligible=no`，`comparison_ready_metrics.csv` 中不再出现该指标。
- 单字/短值残留清洗：`n/M/A/G/T` 等异常非枚举短值在正式输出中转为 `not_reported`，并在 `quality_issues.csv` 记录 `bad_short_value_cleaned`。
- 异常引号继续保留质检：如 P0008 `indirect oxidation”` 同时记录 `truncated_or_odd_quote` 和清洗记录。
- `si_links.csv` 合并 DOI/Crossref 与出版社规则生成的 `publisher_si_candidate_urls`。
- `quality_issues.csv` 新增 `source_issue_id`，`manual_review_queue.csv` 新增同名字段，用于追踪人工复核条目来源。
- `materials.csv` 改为 14 行，exclude 论文保留一行并标记 `table_scope=not_applicable_for_exclude`。
- `doi_enrichment_results.csv` 中 `skipped_metadata_complete` 行回填本地 DOI、标题、期刊和年份，避免统一读取时误判为空。
- `README.md` 已更新为 V2 输出说明，明确正式输出、审计表、质量表和当前边界。

验证结果：

- P0007 的 `current_density = 50 mA cm-2 at 1.38 V, scale needs check` 已从 `comparison_ready_metrics.csv` 移除。
- P0004/P0007 的 `coupled_electrolyzer_performance = n` 已清洗为 `not_reported`。
- `quality_issues.csv` 当前包含 source_issue_id，并保留 P0011 conversion 缺口、P0008 异常引号、P0004/P0007 短值清洗记录。
- `materials.csv` 当前 14 行，P0012 标记为 `not_applicable_for_exclude`。

## V2 闭环开发：SI 下载与 Override 规则沉淀

已继续完成：

- 新增 `--enable-si-download`：在 `--enable-doi-online` 基础上尝试从 DOI/出版社候选入口发现 SI 附件，下载到 `output-test/downloaded-si`。
- 新增 `si_download_report.csv/xlsx` 和总工作簿 `si_download` sheet，记录 candidate_only、html_no_si_links、html_with_candidate_links、discovery_failed、not_pdf、downloaded 等状态。
- SI 下载规则已收紧：普通 article PDF 不再被误当作 SI；只有 URL/页面链接明确带 supp/suppl/support/esm/mmc/MOESM/mediaobjects 等信号才会作为 SI 尝试。
- 新增 DOCX SI 解析：Springer ESM `.docx` 可下载并解析文本。当前 P0011 已成功下载 `P0011_10.1007_s40820-024-01493-3_si.docx`，并在 `si_field_evidence_log.csv` 中输出 HMF concentration、HPLC/product analysis、stability、metal loading 等 SI 证据。
- 下载成功后脚本会自动二次抽取，使 `si_links.csv` 从候选状态推进到本地/下载 SI 关联状态。
- 新增 `apply_hmfor_pattern_rules()`，将 P04/P07/P11 的部分人工校准模式反向沉淀到规则层。
- `apply_manual_overrides()` 只统计实际改变字段，`override_fields` 现在更能反映“仍靠人工补丁”的范围。

当前边界：

- P04/P07 的 SI 入口受限或未暴露直接附件，当前只能输出候选/失败状态。
- P0011 的 SI 已闭环下载和解析，但 HMF conversion 仍未从可解析文本中获得明确量化值，继续保留人工复核状态。
- 自动下载不会绕过出版社权限；403/404 会诚实记录。

## 工程化封装与迁移前加固

已完成迁移前 8 项加固中的核心部分：

- 新增 `TOOL_VERSION=2.1.0` 和 `SCHEMA_VERSION=hmfor-extractor-output-v2.1`。
- 所有 CSV/XLSX 输出表新增 `schema_version` 列。
- 新增 `run_manifest.json`：记录输入/输出目录、运行时间、工具版本、schema 版本、配置路径、manual override 版本、联网/SI/RHE 开关、论文分类统计、表行数、质量问题数、复核数、DOI/SI 状态统计和错误码。
- 新增 `output_schema.json`：机器可读表字段、行数和错误码。
- 新增 `config.yaml`：集中描述路径、运行开关、分类集合、plot gate、SI 下载规则、必填字段。当前代码已读取该配置，后续可逐步把更多硬编码规则迁出。
- 新增 `OUTPUT_SCHEMA.md`：解释每个输出表用途、字段消费边界和 plot gate。
- 新增 `API.md` 和 `hmfor_extractor.extract()` 薄 API：目标系统应调用 CLI/API，不复制业务逻辑。
- 新增 `tests/regression_smoke.py`：固定 P04/P07/P11 关键数值断言、背景/排除反例、P0007 plot gate、P0011 缺口提示、manifest/schema 存在性。

验证：

- `python -m py_compile extract_hmfor_variables.py` 通过。
- `python extract_hmfor_variables.py --enable-doi-online --enable-si-download` 全量通过。
- `python tests/regression_smoke.py` 通过。
- `hmfor_extractor.extract()` API smoke test 通过并返回 manifest。

当前建议迁移方式：

- 核心：保留 `extract_hmfor_variables.py` 与 `hmfor_extractor` API。
- CLI：继续使用 `extract_hmfor_variables.py`。
- 插件/目标系统：只调用 `hmfor_extractor.extract()` 或 CLI，然后读取 `run_manifest.json` 与 `output_schema.json` 判断结果可用性。

