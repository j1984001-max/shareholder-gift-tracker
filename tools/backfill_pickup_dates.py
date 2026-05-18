#!/usr/bin/env python3
from __future__ import annotations

import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import server  # noqa: E402


def main() -> int:
    path = server.NOTICE_SEED_CACHE_PATH
    if not path.exists():
        print(f"Missing seed cache: {path}")
        return 1

    payload = json.loads(path.read_text(encoding="utf-8"))
    updated = 0
    for code, item in payload.items():
        if not isinstance(item, dict):
            continue
        has_pickup = bool(item.get("evotePickupStartDate") or item.get("evotePickupEndDate") or item.get("evotePickupPeriodText"))
        if has_pickup:
            continue

        text_candidates = [
            str(item.get("evotePickupRule", "")),
            str(item.get("giftSummary", "")),
        ]
        merged = " ".join(part for part in text_candidates if part).strip()
        if not merged:
            continue

        start_date, end_date, period_text = server.parse_pickup_roc_range_from_text(merged)
        if not (start_date or end_date or period_text):
            continue

        item["evotePickupStartDate"] = start_date
        item["evotePickupEndDate"] = end_date
        item["evotePickupPeriodText"] = period_text
        if not item.get("evotePickupLocation"):
            item["evotePickupLocation"] = server.extract_pickup_location(
                "開會通知書",
                "",
                str(item.get("evotePickupRule", "")),
                "",
            )
        if not item.get("evotePickupDocuments"):
            item["evotePickupDocuments"] = server.extract_pickup_documents(str(item.get("evotePickupRule", "")))
        updated += 1

    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(f"Updated {updated} records in {path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
