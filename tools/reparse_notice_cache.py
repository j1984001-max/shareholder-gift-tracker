#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import re
import time
import urllib.error
import urllib.parse
from pathlib import Path
import sys
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
CACHE_PATH = ROOT / "data" / "mops_notice_seed_cache.json"
sys.path.insert(0, str(ROOT))

import server  # noqa: E402


def has_pickup_date(item: dict[str, Any]) -> bool:
    return bool(item.get("evotePickupStartDate") and item.get("evotePickupEndDate"))


def has_any_pickup_detail(item: dict[str, Any]) -> bool:
    return any(
        item.get(field)
        for field in (
            "evotePickupStartDate",
            "evotePickupEndDate",
            "evotePickupPeriodText",
            "evotePickupRule",
            "evotePickupLocation",
            "evotePickupDocuments",
        )
    )


def cache_code(cache_key: str, item: dict[str, Any]) -> str:
    code = str(item.get("code") or "").strip()
    if code:
        return code
    return cache_key.split(":", 1)[-1]


def cache_roc_year(cache_key: str, item: dict[str, Any]) -> int:
    if item.get("rocYear"):
        return server.normalize_roc_year(item["rocYear"])
    if ":" in cache_key:
        return server.normalize_roc_year(cache_key.split(":", 1)[0])
    return server.CURRENT_ROC_YEAR


def filename_from_item(code: str, item: dict[str, Any]) -> str:
    filename = str(item.get("filename") or "").strip()
    if filename:
        return filename
    pdf_url = str(item.get("pdfUrl") or "").strip()
    name = urllib.parse.unquote(Path(urllib.parse.urlparse(pdf_url).path).name)
    if name.lower().endswith(".pdf"):
        return name
    return f"{code}_notice.pdf"


def date_year(value: Any) -> int | None:
    raw = str(value or "")
    if len(raw) >= 4 and raw[:4].isdigit():
        return int(raw[:4])
    return None


def clear_out_of_year_pickup(item: dict[str, Any], roc_year: int) -> bool:
    expected_year = roc_year + 1911
    years = [
        year
        for year in (
            date_year(item.get("evotePickupStartDate")),
            date_year(item.get("evotePickupEndDate")),
        )
        if year is not None
    ]
    if not years or all(year == expected_year for year in years):
        return False

    item["evotePickupStartDate"] = None
    item["evotePickupEndDate"] = None
    item["evotePickupPeriodText"] = ""
    return True


def clear_vote_period_only_pickup(item: dict[str, Any]) -> bool:
    if not has_any_pickup_detail(item):
        return False
    rule = str(item.get("evotePickupRule") or "")
    has_dates = bool(item.get("evotePickupStartDate") or item.get("evotePickupEndDate"))
    if not (
        server.looks_like_vote_period_only(rule)
        or has_voting_period_matching_pickup_dates(item)
        or not server.has_meaningful_pickup_rule(rule)
        or (has_dates and not server.has_core_pickup_evidence(rule))
    ):
        return False

    item["evotePickupRule"] = ""
    item["evotePickupStartDate"] = None
    item["evotePickupEndDate"] = None
    item["evotePickupPeriodText"] = ""
    item["evotePickupLocation"] = ""
    item["evotePickupDocuments"] = ""
    return True


def roc_date_pattern(value: Any) -> str:
    raw = str(value or "")
    match = re.fullmatch(r"(\d{4})-(\d{2})-(\d{2})", raw)
    if not match:
        return ""
    year, month, day = (int(part) for part in match.groups())
    roc_year = year - 1911
    return rf"(?:民國)?{roc_year}年0?{month}月0?{day}日|{roc_year}/0?{month}/0?{day}"


def has_voting_period_matching_pickup_dates(item: dict[str, Any]) -> bool:
    rule = server.compact_text(str(item.get("evotePickupRule") or ""))
    start_pattern = roc_date_pattern(item.get("evotePickupStartDate"))
    end_pattern = roc_date_pattern(item.get("evotePickupEndDate"))
    if not rule or not start_pattern or not end_pattern:
        return False
    return bool(
        re.search(
            rf"(行使期間|電子投票期間).{{0,90}}"
            rf"(?:{start_pattern}).{{0,35}}(?:{end_pattern})",
            rule,
        )
    )


def clean_non_pickup_summary(summary: dict[str, Any]) -> tuple[dict[str, Any], bool]:
    rule = str(summary.get("evotePickupRule") or "")
    has_dates = bool(summary.get("evotePickupStartDate") or summary.get("evotePickupEndDate"))
    if not rule or not (
        server.looks_like_vote_period_only(rule)
        or not server.has_meaningful_pickup_rule(rule)
        or (has_dates and not server.has_core_pickup_evidence(rule))
    ):
        return summary, False

    cleaned = dict(summary)
    cleaned["evotePickupRule"] = ""
    cleaned["evotePickupStartDate"] = None
    cleaned["evotePickupEndDate"] = None
    cleaned["evotePickupPeriodText"] = ""
    return cleaned, True


def summary_with_valid_year(summary: dict[str, Any], roc_year: int) -> tuple[dict[str, Any], bool]:
    expected_year = roc_year + 1911
    years = [
        year
        for year in (
            date_year(summary.get("evotePickupStartDate")),
            date_year(summary.get("evotePickupEndDate")),
        )
        if year is not None
    ]
    if not years or all(year == expected_year for year in years):
        return summary, False

    cleaned = dict(summary)
    cleaned["evotePickupStartDate"] = None
    cleaned["evotePickupEndDate"] = None
    cleaned["evotePickupPeriodText"] = ""
    return cleaned, True


def should_try_mops_refresh(item: dict[str, Any], filename: str) -> bool:
    source_type = str(item.get("sourceType") or "").strip()
    if source_type and source_type != "mops":
        return False
    return bool(filename and filename.lower().endswith(".pdf"))


def is_rate_limited_error(error: Any) -> bool:
    text = str(error or "")
    return any(
        marker in text
        for marker in (
            "查詢過量",
            "�d�߹L�q",
            "THE PAGE CANNOT BE ACCESSED",
            "FOR SECURITY REASONS",
            "Too Many Requests",
            "HTTP Error 429",
        )
    )


def download_pdf_url(url: str) -> bytes:
    return server.fetch_bytes(url, timeout=60)


def load_pdf_bytes(
    code: str,
    item: dict[str, Any],
    download_missing: bool,
    refresh_stale_pdf_url: bool,
) -> tuple[bytes | None, str, str]:
    filename = filename_from_item(code, item)
    pdf_path = server.NOTICE_PDF_DIR / filename
    if pdf_path.exists():
        return pdf_path.read_bytes(), "local", ""

    if not download_missing:
        return None, "missing", "local PDF not found"

    pdf_url = str(item.get("pdfUrl") or "").strip()
    if not pdf_url:
        return None, "missing", "no pdfUrl"

    try:
        pdf_bytes = download_pdf_url(pdf_url)
    except Exception as error:
        if is_rate_limited_error(error):
            return None, "rate_limited", str(error)
        if (
            refresh_stale_pdf_url
            and isinstance(error, urllib.error.HTTPError)
            and error.code == 404
            and should_try_mops_refresh(item, filename)
        ):
            try:
                refreshed_url = server.resolve_notice_pdf_url(code, filename)
                pdf_bytes = download_pdf_url(refreshed_url)
                item["pdfUrl"] = refreshed_url
                source = "refreshed"
            except Exception as refresh_error:
                if is_rate_limited_error(refresh_error):
                    return None, "rate_limited", str(refresh_error)
                return None, "refresh_error", str(refresh_error)
        else:
            return None, "download_error", str(error)
    else:
        source = "downloaded"

    if not pdf_bytes.strip():
        return None, "download_error", "empty PDF response"

    pdf_path.parent.mkdir(exist_ok=True)
    pdf_path.write_bytes(pdf_bytes)
    return pdf_bytes, source, ""


def merge_summary(
    item: dict[str, Any],
    summary: dict[str, Any],
    text_metadata: dict[str, Any],
    overwrite_empty: bool,
    roc_year: int,
) -> tuple[bool, bool]:
    clear_out_of_year_pickup(item, roc_year)
    clear_vote_period_only_pickup(item)
    summary, _ = summary_with_valid_year(summary, roc_year)
    summary, _ = clean_non_pickup_summary(summary)
    before_has_date = has_pickup_date(item)
    before_has_detail = has_any_pickup_detail(item)

    if summary.get("giftSummary") or overwrite_empty:
        item["giftSummary"] = summary.get("giftSummary", "") or ""

    if summary.get("evotePickupRule") or overwrite_empty:
        item["evotePickupRule"] = summary.get("evotePickupRule", "") or ""

    for field in ("evotePickupStartDate", "evotePickupEndDate", "evotePickupPeriodText"):
        if summary.get(field) or overwrite_empty:
            item[field] = summary.get(field)

    item["parserVersion"] = server.NOTICE_CACHE_VERSION
    item["textEngine"] = text_metadata.get("engine", "")
    item["textScore"] = text_metadata.get("score")
    item["ocrUsed"] = text_metadata.get("ocrUsed", False)
    item["reparsedAt"] = time.strftime("%Y-%m-%dT%H:%M:%S%z")

    after_has_date = has_pickup_date(item)
    after_has_detail = has_any_pickup_detail(item)
    return after_has_date and not before_has_date, after_has_detail and not before_has_detail


def should_reparse_item(
    item: dict[str, Any],
    only_old: bool,
    include_current: bool,
    only_missing_pickup: bool,
) -> bool:
    parser_version = int(item.get("parserVersion") or 0)
    if only_missing_pickup and has_pickup_date(item):
        return False
    if include_current:
        return True
    if only_old:
        return parser_version < server.NOTICE_CACHE_VERSION
    return True


def save_cache(data: dict[str, Any]) -> None:
    CACHE_PATH.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Reparse cached shareholder meeting notices with the current PDF/OCR parser."
    )
    parser.add_argument("codes", nargs="*", help="Optional stock codes or mixed text to reparse.")
    parser.add_argument("--limit", type=int, default=0, help="Maximum number of cache entries to process.")
    parser.add_argument(
        "--only-missing-pickup",
        action="store_true",
        help="Only reparse entries that still do not have a complete electronic-voting pickup date.",
    )
    parser.add_argument(
        "--include-current",
        action="store_true",
        help="Also reparse entries already using the current parser version.",
    )
    parser.add_argument(
        "--no-download",
        action="store_true",
        help="Do not download PDFs that are missing from .cache/mops-notices.",
    )
    parser.add_argument(
        "--no-refresh-stale-pdf-url",
        action="store_true",
        help="Do not resolve a fresh MOPS PDF URL when a cached MOPS PDF URL returns 404.",
    )
    parser.add_argument("--sleep", type=float, default=0.2, help="Seconds to sleep after a downloaded PDF.")
    parser.add_argument(
        "--rate-limit-sleep",
        type=float,
        default=0,
        help="Seconds to wait and retry once when MOPS returns a rate-limit page. 0 means stop safely.",
    )
    parser.add_argument("--save-every", type=int, default=25, help="Persist progress every N reparsed entries.")
    parser.add_argument(
        "--overwrite-empty",
        action="store_true",
        help="Allow empty second-pass fields to overwrite existing fields. Off by default to avoid data loss.",
    )
    parser.add_argument("--dry-run", action="store_true", help="Parse and report without writing cache updates.")
    args = parser.parse_args()

    data: dict[str, Any] = json.loads(CACHE_PATH.read_text(encoding="utf-8"))
    requested_codes = set(server.clean_codes(" ".join(args.codes))) if args.codes else set()

    candidates: list[tuple[str, dict[str, Any]]] = []
    for key, raw_item in data.items():
        if not isinstance(raw_item, dict):
            continue
        code = cache_code(key, raw_item)
        if requested_codes and code not in requested_codes:
            continue
        if should_reparse_item(
            raw_item,
            only_old=True,
            include_current=args.include_current,
            only_missing_pickup=args.only_missing_pickup,
        ):
            candidates.append((key, raw_item))

    if args.limit > 0:
        candidates = candidates[: args.limit]

    stats = {
        "selected": len(candidates),
        "reparsed": 0,
        "downloaded": 0,
        "refreshed": 0,
        "local": 0,
        "missing": 0,
        "failed": 0,
        "new_pickup_dates": 0,
        "new_pickup_details": 0,
        "out_of_year_dates_ignored": 0,
        "vote_period_only_cleared": 0,
        "ocr_used": 0,
        "rate_limited": 0,
    }

    for index, (key, item) in enumerate(candidates, 1):
        code = cache_code(key, item)
        roc_year = cache_roc_year(key, item)
        removed_existing_out_of_year = clear_out_of_year_pickup(item, roc_year)
        removed_vote_period_only = clear_vote_period_only_pickup(item)
        pdf_bytes, source, error = load_pdf_bytes(
            code,
            item,
            download_missing=not args.no_download,
            refresh_stale_pdf_url=not args.no_refresh_stale_pdf_url,
        )
        if source == "rate_limited" and args.rate_limit_sleep > 0:
            print(f"[{index}/{len(candidates)}] {code} rate_limited: sleeping {args.rate_limit_sleep:.0f}s")
            time.sleep(args.rate_limit_sleep)
            pdf_bytes, source, error = load_pdf_bytes(
                code,
                item,
                download_missing=not args.no_download,
                refresh_stale_pdf_url=not args.no_refresh_stale_pdf_url,
            )

        if source == "rate_limited":
            stats["rate_limited"] += 1
            stats["out_of_year_dates_ignored"] += int(removed_existing_out_of_year)
            stats["vote_period_only_cleared"] += int(removed_vote_period_only)
            print(f"[{index}/{len(candidates)}] {code} rate_limited: {error}")
            if not args.dry_run:
                save_cache(data)
                print("  saved progress before stopping on rate limit")
            break

        if not pdf_bytes:
            if source == "missing":
                stats["missing"] += 1
            else:
                stats["failed"] += 1
            stats["out_of_year_dates_ignored"] += int(removed_existing_out_of_year)
            stats["vote_period_only_cleared"] += int(removed_vote_period_only)
            print(f"[{index}/{len(candidates)}] {code} skip {source}: {error}")
            continue

        try:
            text, text_metadata = server.extract_notice_text_with_metadata(pdf_bytes)
            summary = server.extract_notice_summary(text)
            summary, ignored_out_of_year = summary_with_valid_year(summary, roc_year)
        except Exception as error:
            stats["failed"] += 1
            stats["out_of_year_dates_ignored"] += int(removed_existing_out_of_year)
            stats["vote_period_only_cleared"] += int(removed_vote_period_only)
            print(f"[{index}/{len(candidates)}] {code} parse_error: {error}")
            continue

        if source == "downloaded":
            stats["downloaded"] += 1
        elif source == "refreshed":
            stats["refreshed"] += 1
        elif source == "local":
            stats["local"] += 1

        new_date, new_detail = merge_summary(
            item,
            summary,
            text_metadata,
            overwrite_empty=args.overwrite_empty,
            roc_year=roc_year,
        )
        item["code"] = code
        item["rocYear"] = roc_year
        item["year"] = roc_year + 1911
        if source == "refreshed":
            item["cacheStatus"] = "reparsed-refreshed"
        elif source == "downloaded":
            item["cacheStatus"] = "reparsed-download"
        else:
            item["cacheStatus"] = "reparsed-local"

        stats["reparsed"] += 1
        stats["new_pickup_dates"] += int(new_date)
        stats["new_pickup_details"] += int(new_detail)
        stats["out_of_year_dates_ignored"] += int(removed_existing_out_of_year or ignored_out_of_year)
        stats["vote_period_only_cleared"] += int(removed_vote_period_only)
        stats["ocr_used"] += int(bool(text_metadata.get("ocrUsed")))

        start_date = item.get("evotePickupStartDate") or "-"
        end_date = item.get("evotePickupEndDate") or "-"
        engine = text_metadata.get("engine", "")
        score = text_metadata.get("score")
        print(
            f"[{index}/{len(candidates)}] {code} {source} engine={engine} score={score} "
            f"pickup={start_date}~{end_date}"
        )

        if not args.dry_run and args.save_every > 0 and stats["reparsed"] % args.save_every == 0:
            save_cache(data)
            print(f"  saved progress: reparsed={stats['reparsed']}")

        if source in {"downloaded", "refreshed"} and args.sleep > 0 and index < len(candidates):
            time.sleep(args.sleep)

    if not args.dry_run:
        save_cache(data)

    print(
        "done "
        + " ".join(
            f"{name}={value}"
            for name, value in stats.items()
        )
    )
    return 0 if stats["failed"] == 0 and stats["rate_limited"] == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
