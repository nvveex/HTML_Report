#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import subprocess
import sys
from datetime import date
from pathlib import Path


SCRIPT_DIR = Path(__file__).resolve().parent


def run(cmd: list[str]) -> None:
    proc = subprocess.run(cmd)
    if proc.returncode != 0:
        raise SystemExit(proc.returncode)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--bundle-root", default=".")
    parser.add_argument("--workspace-root")
    parser.add_argument("--input-dir")
    parser.add_argument("--output-dir")
    parser.add_argument("--artifacts-dir")
    parser.add_argument("--school-name")
    args = parser.parse_args()

    bundle_root = Path(args.bundle_root).resolve()
    workspace_root = Path(args.workspace_root).resolve() if args.workspace_root else bundle_root / "workspace"
    input_dir = Path(args.input_dir).resolve() if args.input_dir else workspace_root / "input"
    output_dir = Path(args.output_dir).resolve() if args.output_dir else workspace_root / "output"
    artifacts_dir = Path(args.artifacts_dir).resolve() if args.artifacts_dir else workspace_root / "artifacts"

    output_dir.mkdir(parents=True, exist_ok=True)
    artifacts_dir.mkdir(parents=True, exist_ok=True)

    metrics_json = artifacts_dir / "report_metrics.json"
    detect_cmd = [
        sys.executable,
        str(SCRIPT_DIR / "detect_seiue_project.py"),
        "--workspace-root",
        str(workspace_root),
        "--input-dir",
        str(input_dir),
        "--output-json",
        str(artifacts_dir / "workspace_detect.json"),
    ]
    run(detect_cmd)

    extract_cmd = [
        sys.executable,
        str(SCRIPT_DIR / "extract_seiue_metrics.py"),
        "--workspace-root",
        str(workspace_root),
        "--input-dir",
        str(input_dir),
        "--artifacts-dir",
        str(artifacts_dir),
        "--output-json",
        str(metrics_json),
    ]
    if args.school_name:
        extract_cmd.extend(["--school-name", args.school_name])
    run(extract_cmd)

    metrics = json.loads(metrics_json.read_text(encoding="utf-8"))
    school = metrics.get("overview", {}).get("school_name") or "XX学校"
    html_name = f"{school}-希悦系统使用情况说明-{date.today().isoformat()}.html"
    output_html = output_dir / html_name

    render_cmd = [
        sys.executable,
        str(SCRIPT_DIR / "build_seiue_html_report.py"),
        "--metrics-json",
        str(metrics_json),
        "--output-html",
        str(output_html),
    ]
    run(render_cmd)
    print(json.dumps({"output_html": str(output_html.resolve()), "metrics_json": str(metrics_json.resolve())}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
