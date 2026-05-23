#!/usr/bin/env python3
from __future__ import annotations

import os
import argparse
import random
import subprocess
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from tools.local_cycle_state import load_state, now_ts, save_state
import server

WATCHLIST_PATH = ROOT / "data" / "mops_seed_watchlist.txt"
OFFICIAL_SOURCES_PATH = ROOT / "data" / "official_notice_sources.json"
LOCK_PATH = ROOT / ".cache" / "local-mops-cycle.lock"
STATE_PATH = ROOT / ".cache" / "local-mops-cycle-state.json"
VENDOR_PATH = ROOT / ".vendor"
LOG_PREFIX = "[local-mops-cycle]"
OFFICIAL_SCAN_BATCH = 24
OFFICIAL_SCAN_INTERVAL_SECONDS = 12 * 3600
MOPS_LIMIT = 12
MOPS_SLEEP_SECONDS = 10
FULL_MOPS_SLEEP_SECONDS = 0
MOPS_MIN_INTERVAL_SECONDS = 30 * 60
MOPS_RATE_LIMIT_COOLDOWN_SECONDS = 3 * 3600
MOPS_RATE_LIMIT_MAX_MULTIPLIER = 8
MOPS_RATE_LIMIT_JITTER_SECONDS = 15 * 60


def rate_limit_cooldown_seconds(state: dict[str, object]) -> int:
    count = max(1, int(state.get("consecutiveRateLimitCount", 0)) + 1)
    multiplier = min(MOPS_RATE_LIMIT_MAX_MULTIPLIER, 2 ** (count - 1))
    jitter = random.randint(0, MOPS_RATE_LIMIT_JITTER_SECONDS)
    return MOPS_RATE_LIMIT_COOLDOWN_SECONDS * multiplier + jitter


def run_step(command: list[str], env: dict[str, str]) -> int:
    print("")
    print(f"{LOG_PREFIX} $ {' '.join(command)}", flush=True)
    completed = subprocess.run(command, cwd=ROOT, env=env)
    return completed.returncode


def watchlist_count() -> int:
    if not WATCHLIST_PATH.exists():
        return 0
    return len(server.clean_codes(WATCHLIST_PATH.read_text(encoding="utf-8")))


def acquire_lock() -> int | None:
    LOCK_PATH.parent.mkdir(exist_ok=True)
    try:
        return os.open(str(LOCK_PATH), os.O_CREAT | os.O_EXCL | os.O_WRONLY)
    except FileExistsError:
        return None


def release_lock(fd: int | None) -> None:
    if fd is None:
        return
    try:
        os.close(fd)
    finally:
        try:
            LOCK_PATH.unlink()
        except FileNotFoundError:
            pass


def main() -> int:
    parser = argparse.ArgumentParser(description="Run one local shareholder notice refresh cycle.")
    parser.add_argument("--force-mops-now", action="store_true", help="Ignore the normal MOPS interval guard for this manual run.")
    parser.add_argument("--full-mops-now", action="store_true", help="Run a one-off full MOPS sweep across the whole watchlist for this manual run.")
    args = parser.parse_args()

    lock_fd = acquire_lock()
    if lock_fd is None:
        print(f"{LOG_PREFIX} skip: previous cycle is still running", flush=True)
        return 0

    try:
        os.write(lock_fd, f"{os.getpid()} {time.time()}".encode("utf-8"))
        env = os.environ.copy()
        env["PYTHONPATH"] = str(VENDOR_PATH)
        env["PYTHONUNBUFFERED"] = "1"
        env.setdefault("HTTP_REQUEST_DELAY_MIN_MS", "3000")
        env.setdefault("HTTP_REQUEST_DELAY_MAX_MS", "7000")
        state = load_state(STATE_PATH)
        state["cycleCount"] = int(state.get("cycleCount", 0)) + 1
        current_ts = now_ts()
        should_run_official_by_schedule = current_ts - int(state.get("lastOfficialScanAt", 0)) >= OFFICIAL_SCAN_INTERVAL_SECONDS
        should_run_mops = current_ts >= int(state.get("mopsCooldownUntil", 0))
        should_run_mops = should_run_mops and current_ts - int(state.get("lastMopsAttemptAt", 0)) >= MOPS_MIN_INTERVAL_SECONDS
        if args.force_mops_now or args.full_mops_now:
            should_run_mops = True
        should_run_official = False
        mops_rate_limited = False
        current_mops_limit = watchlist_count() if args.full_mops_now else MOPS_LIMIT
        current_mops_sleep = FULL_MOPS_SLEEP_SECONDS if args.full_mops_now else MOPS_SLEEP_SECONDS

        if should_run_mops:
            update_code = run_step(
                [
                    sys.executable,
                    "tools/update_mops_seed.py",
                    "--watchlist-file",
                    str(WATCHLIST_PATH),
                    "--official-pdf-sources",
                    str(OFFICIAL_SOURCES_PATH),
                    "--limit",
                    str(current_mops_limit),
                    "--retry-empty",
                    "--skip-existing",
                    "--prefer-mops",
                    "--sleep",
                    str(current_mops_sleep),
                    "--rotate-offset",
                    str(int(state.get("mopsCursor", 0))),
                ],
                env,
            )
            state["lastMopsAttemptAt"] = current_ts
            state["mopsCursor"] = int(state.get("mopsCursor", 0)) + MOPS_LIMIT
            if update_code == 2:
                state["lastRateLimitAt"] = current_ts
                state["lastRateLimitError"] = "MOPS rate limited"
                state["consecutiveRateLimitCount"] = int(state.get("consecutiveRateLimitCount", 0)) + 1
                cooldown_seconds = rate_limit_cooldown_seconds(state)
                state["mopsCooldownUntil"] = current_ts + cooldown_seconds
                mops_rate_limited = True
                should_run_official = True
                print(
                    f"{LOG_PREFIX} entered MOPS cooldown for {cooldown_seconds}s "
                    f"(rate-limit streak={state['consecutiveRateLimitCount']})",
                    flush=True,
                )
            elif update_code:
                save_state(STATE_PATH, state)
                return update_code
            else:
                state["lastMopsSuccessAt"] = current_ts
                state["mopsCooldownUntil"] = 0
                state["lastRateLimitError"] = ""
                state["consecutiveRateLimitCount"] = 0
                should_run_official = False
        else:
            remaining = int(state.get("mopsCooldownUntil", 0)) - current_ts
            if remaining > 0:
                print(f"{LOG_PREFIX} MOPS cooldown active for {remaining}s", flush=True)
                should_run_official = True
            else:
                should_run_official = should_run_official_by_schedule

        if should_run_official:
            official_code = run_step(
                [
                    sys.executable,
                    "tools/discover_official_notices.py",
                    "--watchlist-file",
                    str(WATCHLIST_PATH),
                    "--sources-file",
                    str(OFFICIAL_SOURCES_PATH),
                    "--scan-cache-file",
                    str(ROOT / "data" / "official_site_scan_cache.json"),
                    "--limit",
                    str(OFFICIAL_SCAN_BATCH),
                    "--max-pages-per-company",
                    "10",
                    "--company-timeout-seconds",
                    "25",
                    "--fetch-engine",
                    "auto",
                    "--skip-cached-pickup",
                    "--skip-recent-attempt-hours",
                    "168",
                    "--rotate-offset",
                    str(int(state.get("officialCursor", 0))),
                    "--update-seed",
                    "--sleep",
                    "0.6",
                ],
                env,
            )
            if official_code:
                return official_code
            state["lastOfficialScanAt"] = current_ts
            state["officialCursor"] = int(state.get("officialCursor", 0)) + OFFICIAL_SCAN_BATCH
        elif mops_rate_limited:
            save_state(STATE_PATH, state)
            return 0

        snapshot_code = run_step(
            [
                sys.executable,
                "tools/build_lookup_snapshot.py",
                "--watchlist-file",
                str(WATCHLIST_PATH),
            ],
            env,
        )
        save_state(STATE_PATH, state)
        return snapshot_code
    finally:
        release_lock(lock_fd)


if __name__ == "__main__":
    raise SystemExit(main())
