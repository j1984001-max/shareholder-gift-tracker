#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import server  # noqa: E402


def unique_codes(values: list[str]) -> list[str]:
    seen: set[str] = set()
    out: list[str] = []
    for value in values:
        if value and value not in seen:
            seen.add(value)
            out.append(value)
    return out


def read_codes_file(path: Path) -> list[str]:
    if not path.exists():
        return []
    return server.clean_codes(path.read_text(encoding="utf-8"))


def main() -> int:
    parser = argparse.ArgumentParser(description="Build a deployable lookup snapshot from current source data and notice cache.")
    parser.add_argument("codes", nargs="*", help="Stock codes or mixed text.")
    parser.add_argument(
        "--watchlist-file",
        default=str(ROOT / "data" / "mops_seed_watchlist.txt"),
        help="Text file containing stock codes to include.",
    )
    parser.add_argument(
        "--output",
        default=str(server.LOOKUP_SNAPSHOT_PATH),
        help="Output JSON path for the built lookup snapshot.",
    )
    parser.add_argument(
        "--allow-live-notice-fetch",
        action="store_true",
        help="Allow build_record() to fetch missing notice PDFs during snapshot generation.",
    )
    args = parser.parse_args()

    codes = server.clean_codes(" ".join(args.codes))
    if args.watchlist_file:
        codes.extend(read_codes_file(Path(args.watchlist_file)))
    codes = unique_codes(codes)
    if not codes:
        print("No stock codes found.")
        return 0

    sources = server.source_bundle()
    records = {}
    for index, code in enumerate(codes, 1):
        print(f"[{index}/{len(codes)}] build {code}", flush=True)
        records[code] = server.build_record(
            code,
            sources,
            allow_live_notice_fetch=args.allow_live_notice_fetch,
        )

    payload = {
        "generatedAt": time.strftime("%Y-%m-%dT%H:%M:%S%z"),
        "sourceStats": server.build_source_stats(sources),
        "records": records,
    }
    output_path = Path(args.output)
    output_path.parent.mkdir(exist_ok=True)
    output_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(f"Wrote {output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
