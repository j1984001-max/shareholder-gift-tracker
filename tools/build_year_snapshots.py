#!/usr/bin/env python3
from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import server  # noqa: E402


def parse_years(raw: str) -> list[int]:
    if not raw:
        return server.DEFAULT_COMPARE_ROC_YEARS
    return server.parse_roc_years(raw)


def main() -> int:
    parser = argparse.ArgumentParser(description="Build lookup snapshots for multiple shareholder meeting years.")
    parser.add_argument(
        "--years",
        default=",".join(str(year) for year in server.DEFAULT_COMPARE_ROC_YEARS),
        help="Comma-separated ROC or western years, e.g. 115,114,113 or 2026,2025,2024.",
    )
    parser.add_argument(
        "--watchlist-file",
        default=str(ROOT / "data" / "mops_seed_watchlist.txt"),
        help="Text file containing stock codes to include.",
    )
    parser.add_argument(
        "--allow-live-notice-fetch",
        action="store_true",
        help="Allow each snapshot build to fetch missing notice PDFs during generation.",
    )
    args = parser.parse_args()

    years = parse_years(args.years)
    for year in years:
        command = [
            sys.executable,
            "tools/build_lookup_snapshot.py",
            "--roc-year",
            str(year),
            "--watchlist-file",
            args.watchlist_file,
        ]
        if args.allow_live_notice_fetch:
            command.append("--allow-live-notice-fetch")
        print("")
        print(f"$ {' '.join(command)}", flush=True)
        completed = subprocess.run(command, cwd=ROOT)
        if completed.returncode:
            return completed.returncode
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
