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
import json
import subprocess
import sys
from datetime import datetime, timezone
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
LOG = REPO / "refresh.log"
DATA = REPO / "docs" / "data"
MAX_LOG_BYTES = 512_000


def _hidden_process_options() -> dict:
    """Prevent git.exe / python.exe child windows flashing during scheduled runs."""
    if os.name != "nt":
        return {}
    startupinfo = subprocess.STARTUPINFO()
    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    startupinfo.wShowWindow = 0
    return {
        "startupinfo": startupinfo,
        "creationflags": getattr(subprocess, "CREATE_NO_WINDOW", 0),
    }


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
    kw = {**_hidden_process_options(), **kw}
    return subprocess.run(
        args, cwd=REPO, capture_output=True, text=True, timeout=600, **kw
    )


def sync_clean_main() -> None:
    """Ensure the local repo is cleanly synced with origin/main and not wedged in a rebase."""
    run(["git", "rebase", "--abort"])
    run(["git", "merge", "--abort"])
    run(["git", "fetch", "origin", "main"])
    run(["git", "reset", "--hard", "origin/main"])


def regressed_sources(previous: bytes | None, current_path: Path) -> list[str]:
    """Return sources whose new snapshot timestamp is older than the published one."""
    if not previous or not current_path.exists():
        return []
    try:
        old = json.loads(previous)
        new = json.loads(current_path.read_text(encoding="utf-8"))
        old_sources = {
            row["key"]: row for row in old.get("meta", {}).get("sources", [])
            if row.get("key") and row.get("last_modified")
        }
        regressed = []
        for row in new.get("meta", {}).get("sources", []):
            key = row.get("key")
            old_row = old_sources.get(key)
            if not old_row or not row.get("ok") or not row.get("last_modified"):
                continue
            old_time = datetime.fromisoformat(old_row["last_modified"])
            new_time = datetime.fromisoformat(row["last_modified"])
            if new_time < old_time:
                regressed.append(
                    f"{key} ({row['last_modified']} < {old_row['last_modified']})"
                )
        return regressed
    except (OSError, ValueError, TypeError, KeyError):
        # Do not block a refresh merely because an old snapshot predates source metadata.
        return []


def main() -> int:
    env = os.environ.copy()
    # Prefer app-only Graph auth for unattended runs. It works without a signed-in
    # desktop session and avoids falling back to a stale OneDrive copy.
    env.setdefault("SAFRIPOL_AUTH_MODE", "app")
    env["PYTHONPATH"] = str(REPO)
    env["PYTHONUTF8"] = "1"

    # Ensure we are working from the latest remote state before generating data.
    sync_clean_main()

    # Keep the current snapshot so a failed build can be rolled back.
    backup = None
    snapshot = DATA / "snapshot.json"
    if snapshot.exists():
        backup = snapshot.read_bytes()

    log("refresh starting")
    proc = subprocess.run(
        [sys.executable, "-u", "-m", "ingest", "--quiet"],
        cwd=REPO, capture_output=True, text=True, env=env, timeout=900,
        **_hidden_process_options(),
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

    regressions = regressed_sources(backup, snapshot)
    if regressions:
        log("stale source regression - not publishing: " + "; ".join(regressions))
        if backup:
            snapshot.write_bytes(backup)
        version = DATA / "version.json"
        run(["git", "restore", "--source=HEAD", "--", str(version.relative_to(REPO))])
        return 3

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

    push = run(["git", "push", "origin", "main"])
    if push.returncode != 0:
        # Remote moved during ingest. Re-align on top of remote HEAD without rebasing old JSON diffs.
        log("push rejected; re-aligning with remote HEAD and retrying")
        run(["git", "fetch", "origin", "main"])
        run(["git", "reset", "--soft", "origin/main"])
        run(["git", "-c", "user.name=safripol-report-bot",
             "-c", "user.email=safripol-report-bot@connectlogistics.co.za",
             "commit", "-m", f"data: refresh snapshot {stamp}"])
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
