# HMFOR Extractor API

This project is organized as a tool kernel plus thin adapters. The current stable boundary is the CLI, and the Python API delegates to that CLI so extraction logic stays in one place.

## Python API

```python
from hmfor_extractor import extract

manifest = extract(
    input_dir="test",
    output_dir="output-test",
    config_path="config.yaml",
    enable_doi_online=True,
    enable_si_download=True,
    enable_potential_conversion=False,
)
print(manifest["schema_version"])
```

Returns the parsed `run_manifest.json` when available.

## CLI

```powershell
.venv\Scripts\python.exe extract_hmfor_variables.py --pdf-dir test --out-dir output-test --config config.yaml
```

Optional flags:

- `--enable-doi-online`: query Crossref into `doi_enrichment_results.csv`.
- `--enable-si-download`: discover/download SI PDF/DOCX candidates and rerun extraction with downloaded SI.
- `--enable-potential-conversion`: write candidate RHE conversions into `potential_conversion_audit.csv`.
- `--limit N`: process only the first N main PDFs.

## Adapter Rule

External systems should call either the CLI or `hmfor_extractor.extract()`. Do not copy parsing, DOI, SI, RHE, quality scoring, or output-generation logic into the adapter layer.
