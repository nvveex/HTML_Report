#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
from pathlib import Path


def summarize_workspace(workspace_root: Path, input_dir: Path) -> dict[str, object]:
    xlsx_files = sorted(input_dir.glob("*.xlsx")) if input_dir.exists() else []
    output_dir = workspace_root / "output"
    artifacts_dir = workspace_root / "artifacts"
    detected = input_dir.exists() and len(xlsx_files) > 0
    return {
        "workspace_root": str(workspace_root),
        "input_dir": str(input_dir),
        "output_dir": str(output_dir),
        "artifacts_dir": str(artifacts_dir),
        "detected": detected,
        "xlsx_count": len(xlsx_files),
        "sample_xlsx": [p.name for p in xlsx_files[:8]],
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--workspace-root", required=True)
    parser.add_argument("--input-dir", required=True)
    parser.add_argument("--output-json")
    args = parser.parse_args()

    workspace_root = Path(args.workspace_root).resolve()
    input_dir = Path(args.input_dir).resolve()
    summary = summarize_workspace(workspace_root, input_dir)
    text = json.dumps(summary, ensure_ascii=False, indent=2)
    print(text)
    if args.output_json:
        Path(args.output_json).write_text(text, encoding="utf-8")
    return 0 if summary["detected"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
