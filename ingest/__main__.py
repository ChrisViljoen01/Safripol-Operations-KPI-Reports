"""CLI entry point.

    python -m ingest                      refresh from Microsoft Graph
    python -m ingest --local "<folder>"   refresh from a local folder of workbooks
    python -m ingest --check              verify Graph access to every source
"""

from __future__ import annotations

import argparse
import json
import logging
import os
import sys

from .config import DATA_DIR, SOURCES, settings
from .snapshot import build_snapshot, write_snapshot


def _check() -> int:
    from .graph import GraphClient, GraphError

    try:
        client = GraphClient()
    except GraphError as exc:
        print(f"FAIL  {exc}")
        return 2
    print(f"Auth: {client.mode} - signed in as {client.signed_in_as()}\n")
    rc = 0
    for source in SOURCES:
        try:
            meta = client.item_metadata(source.url)
            print(f"OK    {source.label}  ({meta.get('lastModifiedDateTime')}, "
                  f"{round(meta.get('size', 0) / 1024)} KB)")
        except Exception as exc:  # noqa: BLE001
            flag = "FAIL " if source.required else "WARN "
            print(f"{flag} {source.label}: {exc}")
            if source.required:
                rc = 1
    return rc


def _login() -> int:
    """One-off interactive sign-in that seeds the refresh-token cache."""
    from .graph import TOKEN_CACHE_PATH, GraphClient, GraphError

    settings.allow_interactive = True
    settings.auth_mode = "delegated"
    try:
        client = GraphClient(prefer="delegated")
        client.item_metadata(SOURCES[0].url)
    except GraphError as exc:
        print(f"\nSign-in failed: {exc}")
        return 1
    except Exception as exc:  # noqa: BLE001
        print(f"\nSigned in, but the first read failed: {exc}")
        return 1
    print(f"\nSigned in as {client.signed_in_as()}.")
    print(f"Token cached at {TOKEN_CACHE_PATH}")
    print("Scheduled refreshes on this machine will now run without prompting.")
    return 0


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(prog="ingest")
    ap.add_argument("--local", default="", help="folder holding local copies of the workbooks")
    ap.add_argument("--offline", action="store_true", help="skip Graph entirely")
    ap.add_argument("--check", action="store_true", help="verify Graph access and exit")
    ap.add_argument("--login", action="store_true",
                    help="sign in once (device code) and cache the refresh token")
    ap.add_argument("--out", default="", help="output folder (default docs/data)")
    ap.add_argument("--quiet", action="store_true")
    args = ap.parse_args(argv)

    logging.basicConfig(
        level=logging.WARNING if args.quiet else logging.INFO,
        format="%(levelname)-7s %(message)s",
    )

    if args.local:
        settings.local_dir = args.local
        os.environ["SAFRIPOL_LOCAL_DIR"] = args.local
    if args.offline:
        settings.offline = True

    if args.check:
        return _check()

    if args.login:
        return _login()

    snapshot = build_snapshot()
    from pathlib import Path

    out_dir = Path(args.out) if args.out else DATA_DIR
    path = write_snapshot(snapshot, out_dir)

    meta = snapshot["meta"]
    h = snapshot["headline"]
    print(f"\nSnapshot written: {path}")
    print(f"  Generated      : {meta['generated_display']}")
    print(f"  Delivered      : {h['delivered_mt']} MT of {h['target_total_mt']} MT "
          f"({(h['completion_pct'] or 0) * 100:.2f}%)")
    print(f"  Received       : {h['received_admin_mt']} MT")
    print(f"  Connect SOH    : {h['pta_balance_mt']} MT")
    print(f"  Loads          : {h['loads_total']}")
    print(f"  Avg decant     : {h['avg_decant_hours']} h")
    print(f"  Exceptions     : {h['open_exceptions']} high severity")
    for s in meta["sources"]:
        state = "ok" if s["ok"] else "UNAVAILABLE"
        print(f"  source {s['key']:<9} {state:<12} via {s['origin']}")
    return 2 if meta["degraded"] else 0


if __name__ == "__main__":
    sys.exit(main())
