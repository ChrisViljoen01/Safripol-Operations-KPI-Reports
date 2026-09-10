"""Scheduled refresh: rebuild the snapshot and publish it to GitHub Pages.

Delegated auth means the refresh has to run somewhere that holds a signed-in
token cache, so this runs on an operator machine rather than on a CI runner.
It is deliberately quiet and defensive: a failed refresh must never take the
published report down, so a bad build is discarded and the previous snapshot
stays live.

Register it with tools/install_task.ps1.
"""

from __future__ import annotations

import os
import shutil
import subprocess
import sys
from datetime import datetime, timezone
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
LOG = REPO / "refresh.log"
DATA = REPO / "docs" / "data"
MAX_LOG_BYTES = 512_000


def log(msg: str) -> None:
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"{stamp}  {msg}"
    print(line, flush=True)
    try:
        if LOG.exists() and LOG.stat().st_size > MAX_LOG_BYTES:
            LOG.replace(LOG.with_suffix(".log.1"))
        with LOG.open("a", encoding="utf-8") as fh:
            fh.write(line + "\n")
    except OSError:
        pass


def run(args: list[str], **kw) -> subprocess.CompletedProcess:
    return subprocess.run(
        args, cwd=REPO, capture_output=True, text=True, timeout=600, **kw
    )


def main() -> int:
    env = os.environ.copy()
    env.setdefault("SAFRIPOL_AUTH_MODE", "delegated")
    env["PYTHONPATH"] = str(REPO)
    env["PYTHONUTF8"] = "1"

    # Keep the current snapshot so a failed build can be rolled back.
    backup = None
    snapshot = DATA / "snapshot.json"
    if snapshot.exists():
        backup = snapshot.read_bytes()

    log("refresh starting")
    proc = subprocess.run(
        [sys.executable, "-u", "-m", "ingest", "--quiet"],
        cwd=REPO, capture_output=True, text=True, env=env, timeout=900,
    )
    tail = (proc.stdout or "").strip().splitlines()[-6:]
    for line in tail:
        log(f"  {line}")

    if proc.returncode not in (0, 2):
        log(f"ingest FAILED rc={proc.returncode}: {(proc.stderr or '')[-500:]}")
        if backup:
            snapshot.write_bytes(backup)
            log("restored previous snapshot")
        return 1

    if proc.returncode == 2:
        # A required source was unreachable, so the build is missing real
        # figures. Publishing it would blank a working report.
        log("degraded build - not publishing; keeping last good snapshot")
        if backup:
            snapshot.write_bytes(backup)
        return 2

    # Only commit the data files. Code changes are pushed deliberately, by hand.
    run(["git", "add", "docs/data/snapshot.json", "docs/data/version.json"])
    staged = run(["git", "diff", "--cached", "--quiet"])
    if staged.returncode == 0:
        log("no data change")
        return 0

    stamp = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")
    run(["git", "-c", "user.name=safripol-report-bot",
         "-c", "user.email=safripol-report-bot@connectlogistics.co.za",
         "commit", "-m", f"data: refresh snapshot {stamp}"])
    run(["git", "pull", "--rebase", "--autostash", "origin", "main"])
    push = run(["git", "push", "origin", "main"])
    if push.returncode != 0:
        log(f"push failed: {(push.stderr or '')[-300:]}")
        return 1

    log("published")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as exc:  # noqa: BLE001
        log(f"unhandled error: {exc}")
        sys.exit(1)
