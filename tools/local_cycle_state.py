#!/usr/bin/env python3
from __future__ import annotations

import json
import time
from pathlib import Path
from typing import Any

DEFAULT_STATE = {
    "cycleCount": 0,
    "mopsCursor": 0,
    "officialCursor": 0,
    "mopsCooldownUntil": 0,
    "lastMopsAttemptAt": 0,
    "lastMopsSuccessAt": 0,
    "lastOfficialScanAt": 0,
    "lastRateLimitAt": 0,
    "lastRateLimitError": "",
    "consecutiveRateLimitCount": 0,
}


def load_state(path: Path) -> dict[str, Any]:
    if not path.exists():
        return dict(DEFAULT_STATE)
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return dict(DEFAULT_STATE)
    state = dict(DEFAULT_STATE)
    if isinstance(payload, dict):
        state.update(payload)
    return state


def save_state(path: Path, state: dict[str, Any]) -> None:
    path.parent.mkdir(exist_ok=True)
    path.write_text(json.dumps(state, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def now_ts() -> int:
    return int(time.time())
