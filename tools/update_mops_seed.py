#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import sys
import time
import urllib.parse
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


def load_official_pdf_sources(path: Path) -> dict[str, list[dict[str, str]]]:
    if not path.exists():
        return {}
    payload = json.loads(path.read_text(encoding="utf-8"))
    sources: dict[str, list[dict[str, str]]] = {}
    if not isinstance(payload, dict):
        return sources
    for raw_code, raw_items in payload.items():
        codes = server.clean_codes(str(raw_code))
        if not codes:
            continue
        code = codes[0]
        items = raw_items if isinstance(raw_items, list) else [raw_items]
        cleaned_items: list[dict[str, str]] = []
        for item in items:
            if isinstance(item, str):
                url = item.strip()
                label = "官方通知書"
                source_type = "official_pdf"
            elif isinstance(item, dict):
                url = str(item.get("url", "")).strip()
                label = str(item.get("label") or item.get("source") or "官方通知書")
                source_type = str(item.get("sourceType") or "official_pdf")
            else:
                continue
            if url:
                cleaned_item = {"url": url, "label": label, "sourceType": source_type}
                for field in (
                    "giftSummary",
                    "evotePickupRule",
                    "evotePickupStartDate",
                    "evotePickupEndDate",
                    "evotePickupPeriodText",
                    "evotePickupLocation",
                    "evotePickupDocuments",
                ):
                    if isinstance(item, dict) and item.get(field):
                        cleaned_item[field] = str(item[field])
                cleaned_items.append(cleaned_item)
        if cleaned_items:
            sources.setdefault(code, []).extend(cleaned_items)
    return sources


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


def has_pickup_details(item: dict | None) -> bool:
    if not item:
        return False
    return any(
        item.get(field)
        for field in (
            "evotePickupStartDate",
            "evotePickupEndDate",
            "evotePickupLocation",
            "evotePickupDocuments",
        )
    )


def filename_from_url(url: str, code: str) -> str:
    path = urllib.parse.urlparse(url).path
    name = urllib.parse.unquote(Path(path).name)
    if name.lower().endswith(".pdf"):
        return name
    return f"{code}_official_notice.pdf"


def fetch_official_pdf_notice_info(code: str, source: dict[str, str]) -> tuple[dict | None, str]:
    url = source["url"]
    try:
        pdf_bytes = server.fetch_bytes(url)
        text = server.extract_notice_text(pdf_bytes)
        summary = server.extract_notice_summary(text)
    except Exception as error:
        return None, str(error)

    for field in (
        "giftSummary",
        "evotePickupRule",
        "evotePickupStartDate",
        "evotePickupEndDate",
        "evotePickupPeriodText",
    ):
        if source.get(field):
            summary[field] = source[field]

    if not any(
        summary.get(field)
        for field in (
            "evotePickupStartDate",
            "evotePickupEndDate",
            "evotePickupLocation",
            "evotePickupRule",
            "evotePickupDocuments",
        )
    ):
        return None, "parsed PDF but no evote pickup details were found"

    entry = {
        "code": code,
        "filename": filename_from_url(url, code),
        "parserVersion": server.NOTICE_CACHE_VERSION,
        "uploadedAt": "",
        "queryUrl": url,
        "pdfUrl": url,
        "fetchedAt": time.strftime("%Y-%m-%dT%H:%M:%S%z"),
        "cacheStatus": "miss",
        "sourceType": source.get("sourceType") or "official_pdf",
        "sourceLabel": source.get("label") or "官方通知書",
        **summary,
    }
    for field in ("evotePickupLocation", "evotePickupDocuments"):
        if source.get(field):
            entry[field] = source[field]
    return entry, ""


def main() -> int:
    parser = argparse.ArgumentParser(description="Update deployable MOPS notice seed cache for selected stocks.")
    parser.add_argument("codes", nargs="*", help="Stock codes or mixed text, e.g. 1101 2317 or '台泥（1101）'.")
    parser.add_argument("--all", action="store_true", help="Update all candidate stocks from the current source bundle.")
    parser.add_argument("--limit", type=int, default=0, help="Maximum number of codes to process. 0 means no limit.")
    parser.add_argument("--skip-existing", action="store_true", help="Skip codes that already exist in the seed cache.")
    parser.add_argument(
        "--retry-empty",
        action="store_true",
        help="When skipping existing records, retry records that exist but still have no pickup details.",
    )
    parser.add_argument("--force", action="store_true", help="Refresh selected codes even if they already exist.")
    parser.add_argument("--watchlist-file", default="", help="Text file containing stock codes to include.")
    parser.add_argument("--official-pdf-sources", default="", help="JSON file mapping stock codes to official PDF URLs.")
    parser.add_argument("--remote-requested-url", default="", help="URL of /api/requested-codes to include searched codes.")
    parser.add_argument("--sleep", type=float, default=0.3, help="Seconds to sleep between MOPS requests.")
    args = parser.parse_args()

    sources = server.source_bundle()
    seed = load_seed()
    official_sources = load_official_pdf_sources(Path(args.official_pdf_sources)) if args.official_pdf_sources else {}
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

    codes.extend(official_sources.keys())

    if args.remote_requested_url:
        try:
            codes.extend(fetch_remote_requested_codes(args.remote_requested_url))
        except Exception as error:
            print(f"Could not fetch remote requested codes: {error}", file=sys.stderr)

    codes = unique_codes(codes)

    if args.skip_existing:
        codes = [
            code
            for code in codes
            if code not in seed
            or code in official_sources
            or (args.retry_empty and not has_pickup_details(seed.get(code)))
        ]

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

        info = None
        error = ""
        for source in official_sources.get(code, []):
            info, error = fetch_official_pdf_notice_info(code, source)
            if info:
                print(f"  official pdf: {source.get('label') or source['url']}")
                break
            print(f"  official pdf skipped: {source.get('label') or source['url']} ({error})")

        if not info:
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
