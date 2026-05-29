#!/usr/bin/env python3
"""Keep the job agent alive and run it after login, wake, and periodic intervals."""

from __future__ import annotations

import argparse
import datetime as dt
import sys
import time
from pathlib import Path

import job_agent


def log(message: str) -> None:
    print(f"{dt.datetime.now().isoformat(timespec='seconds')} {message}", flush=True)


def run_once(config_path: Path) -> None:
    try:
        exit_code = job_agent.run(config_path)
        log(f"job run completed with exit code {exit_code}")
    except Exception as exc:
        log(f"job run failed: {exc}")


def main() -> int:
    parser = argparse.ArgumentParser(description="Run the job agent at login, after wake, and on an interval.")
    parser.add_argument("--config", default=str(job_agent.DEFAULT_CONFIG), help="Path to config.json")
    parser.add_argument("--interval-seconds", type=int, default=21600, help="Periodic run interval while logged in")
    parser.add_argument("--poll-seconds", type=int, default=60, help="Wake detection polling interval")
    args = parser.parse_args()

    config_path = Path(args.config).expanduser().resolve()
    interval_seconds = max(args.interval_seconds, 300)
    poll_seconds = max(args.poll_seconds, 15)
    wake_gap_seconds = poll_seconds * 3

    log("daemon started")
    run_once(config_path)
    last_run = time.time()
    last_tick = time.time()

    while True:
        time.sleep(poll_seconds)
        now = time.time()
        gap = now - last_tick
        woke_from_sleep = gap > wake_gap_seconds
        interval_elapsed = now - last_run >= interval_seconds

        if woke_from_sleep:
            log(f"wake detected after {int(gap)} seconds")
        if woke_from_sleep or interval_elapsed:
            run_once(config_path)
            last_run = time.time()
        last_tick = now


if __name__ == "__main__":
    raise SystemExit(main())
