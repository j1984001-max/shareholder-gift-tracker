#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import server  # noqa: E402


def load_seed() -> dict[str, dict]:
    if server.NOTICE_SEED_CACHE_PATH.exists():
        return json.loads(server.NOTICE_SEED_CACHE_PATH.read_text(encoding="utf-8"))
    return {}


def save_seed(seed: dict[str, dict]) -> None:
    server.NOTICE_SEED_CACHE_PATH.parent.mkdir(exist_ok=True)
    server.NOTICE_SEED_CACHE_PATH.write_text(json.dumps(seed, ensure_ascii=False, indent=2), encoding="utf-8")


def main() -> int:
    parser = argparse.ArgumentParser(description="Update deployable MOPS notice seed cache for selected stocks.")
    parser.add_argument("codes", nargs="*", help="Stock codes or mixed text, e.g. 1101 2317 or '台泥（1101）'.")
    args = parser.parse_args()

    raw_codes = " ".join(args.codes)
    codes = server.clean_codes(raw_codes)
    if not codes:
        print("No stock codes found.", file=sys.stderr)
        return 2

    sources = server.source_bundle()
    seed = load_seed()

    for index, code in enumerate(codes, 1):
        wespai = sources["wespai"].get(code)
        ideal = sources["ideal"].get(code)
        meeting_date = (ideal or {}).get("meeting_date") or (wespai or {}).get("meeting_date")
        print(f"[{index}/{len(codes)}] {code} meeting_date={meeting_date or '-'}")

        info, error = server.safe_get_mops_notice_info(code, meeting_date)
        if not info:
            print(f"  skipped: {error or 'no notice info'}")
            continue

        seed[code] = info
        start_date = info.get("evotePickupStartDate") or "-"
        end_date = info.get("evotePickupEndDate") or "-"
        print(f"  cached: {info.get('filename')} pickup={start_date}~{end_date}")

    save_seed(seed)
    print(f"Updated {server.NOTICE_SEED_CACHE_PATH}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
