#!/usr/bin/env python3
from __future__ import annotations

import argparse
import os
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
WATCHLIST_PATH = ROOT / "data" / "mops_seed_watchlist.txt"
OFFICIAL_SOURCES_PATH = ROOT / "data" / "official_notice_sources.json"
OFFICIAL_SCAN_CACHE_PATH = ROOT / "data" / "official_site_scan_cache.json"
VENDOR_PATH = ROOT / ".vendor"


def run_step(command: list[str], env: dict[str, str]) -> int:
    print("")
    print(f"$ {' '.join(command)}", flush=True)
    completed = subprocess.run(command, cwd=ROOT, env=env)
    return completed.returncode


def main() -> int:
    parser = argparse.ArgumentParser(description="Run the local-first notice refresh pipeline and rebuild deployable lookup data.")
    parser.add_argument("--watchlist-file", default=str(WATCHLIST_PATH), help="Text file containing stock codes.")
    parser.add_argument("--official-limit", type=int, default=120, help="Maximum company sites to scan per run.")
    parser.add_argument("--official-sleep", type=float, default=0.6, help="Seconds to sleep between official site scans.")
    parser.add_argument("--mops-limit", type=int, default=5, help="Maximum MOPS/official fallback updates per run.")
    parser.add_argument("--mops-sleep", type=float, default=10.0, help="Seconds to sleep between MOPS requests.")
    parser.add_argument("--delay-min-ms", type=int, default=3000, help="Minimum HTTP delay in milliseconds.")
    parser.add_argument("--delay-max-ms", type=int, default=7000, help="Maximum HTTP delay in milliseconds.")
    parser.add_argument("--skip-official-scan", action="store_true", help="Skip official site discovery.")
    parser.add_argument("--skip-mops", action="store_true", help="Skip the MOPS/official fallback step.")
    parser.add_argument("--allow-live-notice-fetch", action="store_true", help="Let snapshot build fetch missing notice PDFs.")
    args = parser.parse_args()

    env = os.environ.copy()
    env["PYTHONPATH"] = str(VENDOR_PATH)
    env["HTTP_REQUEST_DELAY_MIN_MS"] = str(args.delay_min_ms)
    env["HTTP_REQUEST_DELAY_MAX_MS"] = str(args.delay_max_ms)

    if not args.skip_official_scan:
        exit_code = run_step(
            [
                sys.executable,
                "tools/discover_official_notices.py",
                "--watchlist-file",
                args.watchlist_file,
                "--sources-file",
                str(OFFICIAL_SOURCES_PATH),
                "--scan-cache-file",
                str(OFFICIAL_SCAN_CACHE_PATH),
                "--limit",
                str(args.official_limit),
                "--max-pages-per-company",
                "10",
                "--skip-cached-pickup",
                "--skip-recent-attempt-hours",
                "168",
                "--update-seed",
                "--sleep",
                str(args.official_sleep),
            ],
            env,
        )
        if exit_code:
            return exit_code

    if not args.skip_mops:
        exit_code = run_step(
            [
                sys.executable,
                "tools/update_mops_seed.py",
                "--watchlist-file",
                args.watchlist_file,
                "--official-pdf-sources",
                str(OFFICIAL_SOURCES_PATH),
                "--limit",
                str(args.mops_limit),
                "--retry-empty",
                "--skip-existing",
                "--prefer-mops",
                "--sleep",
                str(args.mops_sleep),
            ],
            env,
        )
        if exit_code:
            return exit_code

    build_command = [
        sys.executable,
        "tools/build_lookup_snapshot.py",
        "--watchlist-file",
        args.watchlist_file,
    ]
    if args.allow_live_notice_fetch:
        build_command.append("--allow-live-notice-fetch")
    return run_step(build_command, env)


if __name__ == "__main__":
    raise SystemExit(main())
