# HMFOR PDF 抽取工具项目记忆

更新日期：2026-08-30

## Agent 定位

建议名称：HMFOR 实验变量抽取 Agent。更具体的项目名称可以是：Ni-O-Co HMFOR 文献实验变量抽取 Agent。

一句话定义：这个 Agent 用于批量读取 HMFOR/Ni-Co 相关 PDF 文献，抽取催化剂设计、合成参数、电化学条件、产物分析、性能指标、对照组和机制证据，并生成可横向比较的实验变量数据库。

## 任务边界

这个 Agent 负责从 PDF 中抽取实验变量、建立统一对比表、标注原文证据位置、判断信息是否缺失、记录催化剂/条件/性能/表征/对照，并输出 Excel、CSV 或 Markdown 表格。

这个 Agent 不替用户最终判断课题创新性，不无证据编造参数，不把“未报告”脑补成默认值，不直接改原始 PDF，也不自动生成最终实验方案，除非后续明确扩展为实验设计工作流。

## 核心使用场景

- 批量整理 HMFOR/Ni-Co 相关 PDF 文献。
- 回答哪些 Ni/Co 比例、基底、合成方法、电解液浓度和 HMF 浓度被文献使用过。
- 判断哪些论文真正报告了 HPLC、碳平衡、转化率、FDCA 收率、FE 和稳定性。
- 区分 HMFOR 性能证据、OER/HER/UOR/电容背景证据和机制/表征参考。
- 支持后续复现实验对象筛选、横向比较和证据核对。

## 输入类型

- PDF 文件夹：批量抽取并生成总表。
- 单篇 PDF：抽取该论文实验变量，可追加到已有表格。
- 已有表格 + 新 PDF：增量抽取，并在后续版本中避免重复记录。

## 工作目录

`D:\cursor\solo program\experiment input Agent`

## 主要文件

- `extract_hmfor_variables.py`：主脚本。
- `manual_overrides.json`：人工校准规则，存放论文级字段修正，不在主脚本中硬编码具体论文补丁。
- `test`：默认 PDF 输入文件夹。
- `output-test`：默认输出文件夹。

## 默认运行方式

```powershell
.\.venv\Scripts\python.exe extract_hmfor_variables.py
```

如果使用系统 Python，需要先确保依赖已安装：`pymupdf`、`pandas`、`openpyxl`。

## 稳定数据结构

V2 数据底座应优先拆成多张表，而不是无限扩展一张宽表：

- `papers.csv`：一篇一行，标题、DOI、分类、年份、期刊、文本状态。
- `materials.csv`：催化剂、组成、载体、合成方法、金属负载。
- `metrics_long_table.csv`：一个指标一行，保存 performance context、metric name、metric value、condition、evidence sentence 和 source scope。
- `evidence_log.csv`：字段、页码、原文证据、来源范围。
- `quality_issues.csv`：所有可疑点。
- `dedup_report.csv`：重复判定原因。
- `si_links.csv`：主文与 Supporting Information 的关联。

## 主表字段分组

长期字段设计按 8 组维护：文献信息、反应体系、催化剂设计、合成条件、电化学测试、产物与性能、机制与表征、质量判断。

关键字段包括 title、year、journal、doi、reaction_type、target_product、electrolyte、KOH_concentration、HMF_concentration、cell_type、catalyst_name、catalyst_family、substrate、Ni_Co_ratio、synthesis_method、reference_electrode、potential_vs_RHE、current_density、HMF_conversion、FDCA_yield、FDCA_selectivity、Faradaic_efficiency、carbon_balance、product_analysis_method、controls、site_decoupling_evidence_level、reproducibility_value、recommended_use。

## 证据等级

HMFOR 数据质量评分建议保留 1-5 分：

- 1：只报告电流，几乎没有产物分析。
- 2：有 FDCA 数据，但产物分析不完整。
- 3：有转化率、选择性、FE。
- 4：有中间体检测和稳定性。
- 5：有完整产物分布、碳平衡、对照和稳定性。

位点解耦证据等级建议保留 L0-L4：

- L0：只是口头声称协同。
- L1：有 Ni、Co、NiCo 性能对照。
- L2：有结构/价态表征支持电子相互作用。
- L3：有中间体吸附、反应路径或 DFT 支持位点分工。
- L4：有原位/操作态证据 + 对照实验 + DFT 闭环证明。

## 稳定维护规则

- 先判断论文类别，再进入对应抽取路径。
- HMFOR 专属字段只应用于 HMFOR 核心或相关论文。
- 背景论文用于材料、合成、机制或 OER/HER/UOR/电容背景参考，不作为 HMFOR 产物性能证据。
- 性能数据必须绑定测试条件，例如电位、HMF 浓度、KOH 浓度、时间和 cell type。
- 不能把 OER 数据误当作 HMFOR 数据。
- 不能把无 HMF 条件下的电流用于说明 HMF 氧化性能。
- 电位可以额外换算为 vs RHE，但必须保留原参比电极和换算依据。
- 宁可标注“未报告”或“需要人工确认”，不要猜。
- `not_reported` 应尽量语义化，例如 HMFOR 核心论文中 carbon balance 可标注为 `not_reported_in_main_text`。
- 人工校准写入 `manual_overrides.json`，并在输出中保留 override 标记，便于区分算法抽取和人工修正。
- 修改脚本后应至少运行一次默认测试集，并检查输出 CSV/XLSX 是否正常生成。
