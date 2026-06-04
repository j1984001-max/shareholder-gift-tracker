#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import random
import sys
import time
import urllib.parse
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


def load_attempt_log() -> dict[str, dict]:
    if server.MOPS_ATTEMPT_LOG_PATH.exists():
        payload = json.loads(server.MOPS_ATTEMPT_LOG_PATH.read_text(encoding="utf-8"))
        return payload if isinstance(payload, dict) else {}
    return {}


def save_attempt_log(attempts: dict[str, dict]) -> None:
    server.MOPS_ATTEMPT_LOG_PATH.parent.mkdir(exist_ok=True)
    server.MOPS_ATTEMPT_LOG_PATH.write_text(
        json.dumps(dict(sorted(attempts.items())), ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )


def record_mops_attempt(
    attempts: dict[str, dict],
    code: str,
    roc_year: int,
    status: str,
    error: str = "",
    info: dict | None = None,
    meeting_date: str | None = None,
) -> None:
    attempt_key = server.notice_cache_storage_key(code, roc_year)
    previous = attempts.get(attempt_key) if isinstance(attempts.get(attempt_key), dict) else {}
    attempts[attempt_key] = {
        "code": code,
        "rocYear": roc_year,
        "year": roc_year + 1911,
        "attemptedAt": time.strftime("%Y-%m-%dT%H:%M:%S%z"),
        "attemptCount": int(previous.get("attemptCount") or 0) + 1,
        "status": status,
        "error": (error or "")[:500],
        "meetingDate": meeting_date or "",
        "filename": (info or {}).get("filename", ""),
        "sourceType": (info or {}).get("sourceType", "mops" if info else ""),
        "hasNotice": bool(info and ((info or {}).get("filename") or (info or {}).get("pdfUrl"))),
        "hasPickupDate": bool(
            info
            and (
                (info or {}).get("evotePickupStartDate")
                or (info or {}).get("evotePickupEndDate")
                or (info or {}).get("evotePickupPeriodText")
            )
        ),
    }


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
    payload = json.loads(server.fetch_bytes(url).decode("utf-8"))
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


def rate_limited(error: str) -> bool:
    return any(
        phrase in (error or "")
        for phrase in (
            "查詢過量",
            "THE PAGE CANNOT BE ACCESSED",
            "FOR SECURITY REASONS",
            "Too Many Requests",
            "HTTP Error 429",
        )
    )


def prioritize_codes(
    codes: list[str],
    seed: dict[str, dict],
    official_sources: dict[str, list[dict[str, str]]],
    retry_empty: bool,
    roc_year: int,
) -> list[str]:
    missing = [
        code
        for code in codes
        if not server.notice_cache_lookup(seed, code, roc_year)
    ]
    official = [
        code
        for code in missing
        if code in official_sources
    ]
    missing_mops = [code for code in missing if code not in official_sources]
    retryable_empty = [
        code
        for code in codes
        if retry_empty
        and server.notice_cache_lookup(seed, code, roc_year)
        and code not in official_sources
        and not has_pickup_details(server.notice_cache_lookup(seed, code, roc_year))
    ]
    return unique_codes([*official, *missing_mops, *retryable_empty])


def rotate_codes(codes: list[str], offset: int) -> list[str]:
    if not codes:
        return []
    real_offset = offset % len(codes)
    return [*codes[real_offset:], *codes[:real_offset]]


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
        text, text_metadata = server.extract_notice_text_with_metadata(pdf_bytes)
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
        "textEngine": text_metadata.get("engine", ""),
        "textScore": text_metadata.get("score"),
        "ocrUsed": text_metadata.get("ocrUsed", False),
        **summary,
    }
    for field in ("evotePickupLocation", "evotePickupDocuments"):
        if source.get(field):
            entry[field] = source[field]
    return entry, ""


def fetch_from_official_sources(code: str, sources: list[dict[str, str]]) -> tuple[dict | None, str]:
    last_error = ""
    for source in sources:
        info, error = fetch_official_pdf_notice_info(code, source)
        if info:
            print(f"  official pdf: {source.get('label') or source['url']}")
            return info, ""
        last_error = error
        print(f"  official pdf skipped: {source.get('label') or source['url']} ({error})")
    return None, last_error


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
    parser.add_argument("--rotate-offset", type=int, default=0, help="Rotate selected codes before limiting to avoid retrying the same front slice.")
    parser.add_argument("--shuffle", action="store_true", help="Shuffle candidate codes before limiting.")
    parser.add_argument("--prefer-mops", action="store_true", help="Try MOPS before known official PDF fallbacks.")
    parser.add_argument(
        "--roc-year",
        default=str(server.CURRENT_ROC_YEAR),
        help="ROC meeting year to update, e.g. 115, 114, 113. Western years like 2026 also work.",
    )
    args = parser.parse_args()

    roc_year = server.normalize_roc_year(args.roc_year)
    sources = server.source_bundle(roc_year)
    seed = load_seed()
    attempts = load_attempt_log()
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
        codes = prioritize_codes(codes, seed, official_sources, args.retry_empty, roc_year)

    if args.shuffle:
        random.shuffle(codes)
    elif args.rotate_offset:
        codes = rotate_codes(codes, args.rotate_offset)

    if args.limit > 0:
        codes = codes[: args.limit]

    if not codes:
        print("No stock codes found.")
        return 0

    hit_rate_limit = False
    for index, code in enumerate(codes, 1):
        wespai = sources["wespai"].get(code)
        ideal = sources["ideal"].get(code)
        meeting_date = (ideal or {}).get("meeting_date") or (wespai or {}).get("meeting_date")
        print(f"[{index}/{len(codes)}] {code} meeting_date={meeting_date or '-'}")

        if args.force:
            seed.pop(server.notice_cache_storage_key(code, roc_year), None)
            server.NOTICE_CACHE_MEMORY = dict(seed)

        attempted_mops = False
        mops_error = ""
        if args.prefer_mops:
            attempted_mops = True
            info, error = server.safe_get_mops_notice_info(code, meeting_date, roc_year)
            mops_error = error
            if not info and not rate_limited(error):
                fallback_info, fallback_error = fetch_from_official_sources(code, official_sources.get(code, []))
                if fallback_info:
                    info, error = fallback_info, ""
                elif fallback_error:
                    error = fallback_error
        else:
            info, error = fetch_from_official_sources(code, official_sources.get(code, []))
            if not info:
                attempted_mops = True
                info, error = server.safe_get_mops_notice_info(code, meeting_date, roc_year)
                mops_error = error
        if not info:
            print(f"  skipped: {error or 'no notice info'}")
            if attempted_mops:
                status = "rate_limited" if rate_limited(error) else "not_found"
                record_mops_attempt(attempts, code, roc_year, status, error, meeting_date=meeting_date)
                save_attempt_log(attempts)
            if rate_limited(error):
                print("  MOPS appears rate-limited; stopping this run to avoid wasting requests.")
                hit_rate_limit = True
                break
            continue

        seed[server.notice_cache_storage_key(code, roc_year)] = info
        if attempted_mops:
            status = "success" if not mops_error else "fallback_success"
            record_mops_attempt(attempts, code, roc_year, status, mops_error, info, meeting_date=meeting_date)
            save_attempt_log(attempts)
        start_date = info.get("evotePickupStartDate") or "-"
        end_date = info.get("evotePickupEndDate") or "-"
        print(f"  cached: {info.get('filename')} pickup={start_date}~{end_date}")
        if args.sleep > 0 and index < len(codes):
            time.sleep(args.sleep)

    save_seed(seed)
    print(f"Updated {server.NOTICE_SEED_CACHE_PATH}")
    return 2 if hit_rate_limit else 0


if __name__ == "__main__":
    raise SystemExit(main())
