from __future__ import annotations

import subprocess
import sys
from pathlib import Path
from typing import Any


def extract(input_dir: str | Path, output_dir: str | Path, config_path: str | Path | None = None, *, enable_doi_online: bool = False, enable_si_download: bool = False, enable_potential_conversion: bool = False, limit: int | None = None) -> dict[str, Any]:
    """Run the HMFOR extractor through the stable CLI boundary.

    The API intentionally delegates to the CLI script so the target system can call
    one stable function while extraction rules remain centralized in
    extract_hmfor_variables.py. It returns the parsed run_manifest.json when present.
    """
    root = Path(__file__).resolve().parents[1]
    py = root / ".venv" / "Scripts" / "python.exe"
    interpreter = str(py) if py.exists() else sys.executable
    output_dir = Path(output_dir)
    cmd = [interpreter, str(root / "extract_hmfor_variables.py"), "--pdf-dir", str(input_dir), "--out-dir", str(output_dir)]
    if config_path is not None:
        cmd += ["--config", str(config_path)]
    if limit is not None:
        cmd += ["--limit", str(limit)]
    if enable_doi_online:
        cmd.append("--enable-doi-online")
    if enable_si_download:
        cmd.append("--enable-si-download")
    if enable_potential_conversion:
        cmd.append("--enable-potential-conversion")
    subprocess.run(cmd, cwd=root, check=True)
    manifest_path = output_dir / "run_manifest.json"
    if manifest_path.exists():
        import json
        return json.loads(manifest_path.read_text(encoding="utf-8"))
    return {"output_dir": str(output_dir), "manifest": "missing"}
