#!/usr/bin/env python3
from __future__ import annotations

import os
import subprocess
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
WATCHLIST_PATH = ROOT / "data" / "mops_seed_watchlist.txt"
OFFICIAL_SOURCES_PATH = ROOT / "data" / "official_notice_sources.json"
LOCK_PATH = ROOT / ".cache" / "local-mops-cycle.lock"
VENDOR_PATH = ROOT / ".vendor"
LOG_PREFIX = "[local-mops-cycle]"


def run_step(command: list[str], env: dict[str, str]) -> int:
    print("")
    print(f"{LOG_PREFIX} $ {' '.join(command)}", flush=True)
    completed = subprocess.run(command, cwd=ROOT, env=env)
    return completed.returncode


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
    lock_fd = acquire_lock()
    if lock_fd is None:
        print(f"{LOG_PREFIX} skip: previous cycle is still running", flush=True)
        return 0

    try:
        os.write(lock_fd, f"{os.getpid()} {time.time()}".encode("utf-8"))
        env = os.environ.copy()
        env["PYTHONPATH"] = str(VENDOR_PATH)
        env.setdefault("HTTP_REQUEST_DELAY_MIN_MS", "3000")
        env.setdefault("HTTP_REQUEST_DELAY_MAX_MS", "7000")

        update_code = run_step(
            [
                sys.executable,
                "tools/update_mops_seed.py",
                "--watchlist-file",
                str(WATCHLIST_PATH),
                "--official-pdf-sources",
                str(OFFICIAL_SOURCES_PATH),
                "--limit",
                "5",
                "--retry-empty",
                "--skip-existing",
                "--sleep",
                "10",
            ],
            env,
        )
        if update_code:
            return update_code

        return run_step(
            [
                sys.executable,
                "tools/build_lookup_snapshot.py",
                "--watchlist-file",
                str(WATCHLIST_PATH),
            ],
            env,
        )
    finally:
        release_lock(lock_fd)


if __name__ == "__main__":
    raise SystemExit(main())
