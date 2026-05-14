#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import sys
import time
import urllib.request
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


def read_codes_file(path: Path) -> list[str]:
    if not path.exists():
        return []
    return server.clean_codes(path.read_text(encoding="utf-8"))


def fetch_remote_requested_codes(url: str) -> list[str]:
    if not url:
        return []
    with urllib.request.urlopen(url, timeout=30) as response:
        payload = json.loads(response.read().decode("utf-8"))
    return [code for code in payload.get("codes", []) if isinstance(code, str)]


def unique_codes(codes: list[str]) -> list[str]:
    seen: set[str] = set()
    out: list[str] = []
    for code in codes:
        if code not in seen:
            seen.add(code)
            out.append(code)
    return out


def main() -> int:
    parser = argparse.ArgumentParser(description="Update deployable MOPS notice seed cache for selected stocks.")
    parser.add_argument("codes", nargs="*", help="Stock codes or mixed text, e.g. 1101 2317 or '台泥（1101）'.")
    parser.add_argument("--all", action="store_true", help="Update all candidate stocks from the current source bundle.")
    parser.add_argument("--limit", type=int, default=0, help="Maximum number of codes to process. 0 means no limit.")
    parser.add_argument("--skip-existing", action="store_true", help="Skip codes that already exist in the seed cache.")
    parser.add_argument("--force", action="store_true", help="Refresh selected codes even if they already exist.")
    parser.add_argument("--watchlist-file", default="", help="Text file containing stock codes to include.")
    parser.add_argument("--remote-requested-url", default="", help="URL of /api/requested-codes to include searched codes.")
    parser.add_argument("--sleep", type=float, default=0.3, help="Seconds to sleep between MOPS requests.")
    args = parser.parse_args()

    sources = server.source_bundle()
    seed = load_seed()
    if args.all:
        source_codes = set(sources["wespai"]) | set(sources["ideal"]) | set(sources["honsec"])
        codes = sorted(
            code
            for code in source_codes
            if server.should_fetch_mops_notice(sources["ideal"].get(code), sources["honsec"].get(code))
        )
    else:
        raw_codes = " ".join(args.codes)
        codes = server.clean_codes(raw_codes)

    if args.watchlist_file:
        codes.extend(read_codes_file(Path(args.watchlist_file)))

    if args.remote_requested_url:
        try:
            codes.extend(fetch_remote_requested_codes(args.remote_requested_url))
        except Exception as error:
            print(f"Could not fetch remote requested codes: {error}", file=sys.stderr)

    codes = unique_codes(codes)

    if args.skip_existing:
        codes = [code for code in codes if code not in seed]

    if args.limit > 0:
        codes = codes[: args.limit]

    if not codes:
        print("No stock codes found.")
        return 0

    for index, code in enumerate(codes, 1):
        wespai = sources["wespai"].get(code)
        ideal = sources["ideal"].get(code)
        meeting_date = (ideal or {}).get("meeting_date") or (wespai or {}).get("meeting_date")
        print(f"[{index}/{len(codes)}] {code} meeting_date={meeting_date or '-'}")

        if args.force:
            seed.pop(code, None)
            server.NOTICE_CACHE_MEMORY = dict(seed)

        info, error = server.safe_get_mops_notice_info(code, meeting_date)
        if not info:
            print(f"  skipped: {error or 'no notice info'}")
            continue

        seed[code] = info
        start_date = info.get("evotePickupStartDate") or "-"
        end_date = info.get("evotePickupEndDate") or "-"
        print(f"  cached: {info.get('filename')} pickup={start_date}~{end_date}")
        if args.sleep > 0 and index < len(codes):
            time.sleep(args.sleep)

    save_seed(seed)
    print(f"Updated {server.NOTICE_SEED_CACHE_PATH}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
