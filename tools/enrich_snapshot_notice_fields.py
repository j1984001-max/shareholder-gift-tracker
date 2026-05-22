from __future__ import annotations

import json
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import server


SNAPSHOT_PATH = ROOT / "data" / "lookup_snapshot.json"
CACHE_PATH = ROOT / "data" / "mops_notice_seed_cache.json"


def main() -> int:
    snapshot = json.loads(SNAPSHOT_PATH.read_text())
    cache = json.loads(CACHE_PATH.read_text())
    records = snapshot.get("records", {})
    updated = 0

    for code, record in records.items():
        cache_item = cache.get(code)
        if isinstance(cache_item, dict):
            record["noticeGiftSummary"] = cache_item.get("giftSummary", "") or record.get("noticeGiftSummary", "")
            record["evotePickupRule"] = cache_item.get("evotePickupRule", "") or record.get("evotePickupRule", "")
            record["evotePickupStartDate"] = cache_item.get("evotePickupStartDate") or record.get("evotePickupStartDate")
            record["evotePickupEndDate"] = cache_item.get("evotePickupEndDate") or record.get("evotePickupEndDate")
            record["noticeEvotePickupPeriodText"] = cache_item.get("evotePickupPeriodText", "") or record.get(
                "noticeEvotePickupPeriodText", ""
            )
            record["evotePickupLocation"] = cache_item.get("evotePickupLocation", "") or record.get(
                "evotePickupLocation", ""
            )
            record["evotePickupDocuments"] = cache_item.get("evotePickupDocuments", "") or record.get(
                "evotePickupDocuments", ""
            )
        server.enrich_record_notice_fields(record)
        updated += 1

    snapshot["records"] = records
    SNAPSHOT_PATH.write_text(json.dumps(snapshot, ensure_ascii=False, indent=2) + "\n")
    print(f"updated={updated}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
