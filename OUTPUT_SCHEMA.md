# HMFOR Extractor Output Schema

Current schema: `hmfor-extractor-output-v2.1`.

Every CSV/XLSX output table includes a `schema_version` column. Machine consumers should check `run_manifest.json` and `output_schema.json` before reading downstream tables.

## Core Files

- `run_manifest.json`: run time, input/output paths, tool/config/manual override versions, runtime flags, paper counts, issue counts, DOI/SI status summaries, and standard error codes.
- `output_schema.json`: machine-readable list of all output tables, columns, row counts, schema version, and error code dictionary.
- `hmfor_experiment_variables.xlsx`: human-facing workbook containing all tables as sheets.

## Main Tables

- `papers.csv`: one paper per row; bibliographic metadata, category, text status, reference-section removal audit.
- `hmfor_experiment_variables.csv`: HMFOR core/related papers, one row per paper.
- `background_or_excluded_papers.csv`: non-HMFOR papers after category routing.
- `materials.csv`: one row per paper, including exclude rows as `not_applicable_for_exclude`.
- `metrics_long_table.csv`: one metric per row; includes comparison eligibility and exclusion reason.
- `comparison_ready_metrics.csv`: only metrics passing the plot/comparison gate.

## Evidence And Quality

- `hmfor_evidence_log.csv`: main-text evidence snippets.
- `si_links.csv`: local/downloaded SI association and publisher SI candidate URLs.
- `si_download_report.csv`: SI discovery/download audit, including failed attempts and candidate-only URLs.
- `si_field_evidence_log.csv`: field-level evidence extracted from SI PDF/DOCX attachments.
- `quality_issues.csv`: all machine-detectable issues, each with `source_issue_id`.
- `manual_review_queue.csv`: prioritized human review queue; rows sourced from quality issues carry the same `source_issue_id`.
- `quality_scores.csv`: sub-scores and overall quality score.

## Audit Tables

- `dedup_report.csv`: DOI/title/hash duplicate audit.
- `doi_enrichment_queue.csv`: offline queue of records needing DOI/publisher metadata enrichment.
- `doi_enrichment_results.csv`: Crossref lookup results when `--enable-doi-online` is used; otherwise skipped/offline statuses.
- `potential_conversion_audit.csv`: original potentials, candidate conversion, final trusted potential, trust level, assumptions.
- `unit_normalization_log.csv`: raw-to-normalized unit changes.
- `visualization_readiness_report.csv`: context-level readiness summary for plotting.

## Plot Gate

A metric is allowed into `comparison_ready_metrics.csv` only when it is comparison-eligible, numeric after normalization, not from references, not `very_low` confidence, not a non-primary duplicate, and not marked with `scale needs check`.

## Stable Error Codes

See `run_manifest.json` or `output_schema.json` for the current error code dictionary, including `PDF_TEXT_EMPTY`, `DOI_LOOKUP_FAILED`, `SI_NOT_FOUND`, `SI_DOWNLOAD_FAILED`, `RHE_REFERENCE_UNCLEAR`, `FIELD_EVIDENCE_WEAK`, `PLOT_GATE_BLOCKED`, `MANUAL_OVERRIDE_USED`, and `BAD_SHORT_VALUE_CLEANED`.
