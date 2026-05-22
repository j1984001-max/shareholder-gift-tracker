from __future__ import annotations

import json
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
CACHE_PATH = ROOT / "data" / "mops_notice_seed_cache.json"
sys.path.insert(0, str(ROOT))

import server


def main() -> int:
    data = json.loads(CACHE_PATH.read_text())
    reparsed = 0
    cleaned_only = 0

    for code, item in data.items():
        filename = item.get("filename") or ""
        pdf_path = server.NOTICE_PDF_DIR / filename if filename else None
        summary = None

        if pdf_path and pdf_path.exists():
            text = server.extract_notice_text(pdf_path.read_bytes())
            summary = server.extract_notice_summary(text)
            reparsed += 1
        else:
            cleaned_summary = server.trim_notice_summary(item.get("giftSummary", ""))
            cleaned_rule = server.strip_notice_noise(item.get("evotePickupRule", ""))
            start_date, end_date, period_text = server.parse_pickup_roc_range_from_text(cleaned_rule)
            summary = {
                "giftSummary": cleaned_summary,
                "evotePickupRule": cleaned_rule,
                "evotePickupStartDate": start_date,
                "evotePickupEndDate": end_date,
                "evotePickupPeriodText": period_text,
            }
            cleaned_only += 1

        item["giftSummary"] = summary.get("giftSummary", "") or ""
        item["evotePickupRule"] = summary.get("evotePickupRule", "") or ""
        item["evotePickupStartDate"] = summary.get("evotePickupStartDate")
        item["evotePickupEndDate"] = summary.get("evotePickupEndDate")
        item["evotePickupPeriodText"] = summary.get("evotePickupPeriodText", "") or ""

    CACHE_PATH.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n")
    print(f"reparsed={reparsed} cleaned_only={cleaned_only} total={len(data)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
