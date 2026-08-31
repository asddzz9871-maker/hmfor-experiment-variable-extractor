
from pathlib import Path
import subprocess
import sys
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
PY = ROOT / ".venv" / "Scripts" / "python.exe"
OUT = ROOT / "output-regression"


def run_extractor():
    cmd = [str(PY if PY.exists() else sys.executable), str(ROOT / "extract_hmfor_variables.py"), "--out-dir", str(OUT), "--enable-doi-online"]
    subprocess.run(cmd, cwd=ROOT, check=True)


def assert_equal(actual, expected, label):
    if str(actual) != str(expected):
        raise AssertionError(f"{label}: expected {expected!r}, got {actual!r}")


def main():
    run_extractor()
    hmfor = pd.read_csv(OUT / "hmfor_experiment_variables.csv")
    metrics = pd.read_csv(OUT / "comparison_ready_metrics.csv")
    quality = pd.read_csv(OUT / "quality_issues.csv")
    materials = pd.read_csv(OUT / "materials.csv")
    manifest = OUT / "run_manifest.json"
    schema = OUT / "output_schema.json"
    if not manifest.exists() or not schema.exists():
        raise AssertionError("manifest/schema files missing")
    for df_name, df in [("hmfor", hmfor), ("metrics", metrics), ("quality", quality), ("materials", materials)]:
        if "schema_version" not in df.columns:
            raise AssertionError(f"{df_name} missing schema_version")
    p04 = hmfor[hmfor.paper_id == "P0004"].iloc[0]
    assert_equal(p04["hmf_conversion_value"], "100.0", "P0004 conversion")
    assert_equal(p04["fdca_yield_value"], "95.6", "P0004 FDCA yield")
    assert_equal(p04["faradaic_efficiency_value"], "90.5", "P0004 FE")
    p07 = hmfor[hmfor.paper_id == "P0007"].iloc[0]
    assert_equal(p07["hmf_conversion_value"], "99.6", "P0007 conversion")
    assert_equal(p07["fdca_yield_value"], "96.5", "P0007 FDCA yield")
    assert_equal(p07["faradaic_efficiency_value"], "95.9", "P0007 FE")
    if ((metrics.paper_id == "P0007") & (metrics.metric_name == "current_density")).any():
        raise AssertionError("P0007 uncertain current_density entered comparison_ready_metrics")
    p11 = hmfor[hmfor.paper_id == "P0011"].iloc[0]
    assert_equal(p11["fdca_yield_value"], "99.5", "P0011 FDCA yield")
    assert_equal(p11["faradaic_efficiency_value"], "99.6", "P0011 FE")
    if "reported_in_figure_not_quantified_in_text" not in set(quality.issue_type.astype(str)):
        raise AssertionError("P0011 figure-not-quantified issue missing")
    p12 = materials[materials.paper_id == "P0012"].iloc[0]
    assert_equal(p12["table_scope"], "not_applicable_for_exclude", "P0012 materials scope")
    print("regression smoke passed")


if __name__ == "__main__":
    main()
