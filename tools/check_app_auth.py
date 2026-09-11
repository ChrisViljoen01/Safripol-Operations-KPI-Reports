"""Preflight for cloud (app-only) refresh.

Answers one question: can this app read every source on its own, with no human
signed in? That is the only thing standing between the current PC-bound refresh
and an unattended GitHub Actions schedule.

Run it after the Entra admin grants consent:

    python tools\\check_app_auth.py

Exit code 0 means the refresh can move to GitHub Actions.
"""

from __future__ import annotations

import base64
import json
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))

from dotenv import load_dotenv  # noqa: E402

load_dotenv(REPO / ".env")

from ingest.config import SOURCES, settings  # noqa: E402
from ingest.graph import GraphClient, GraphError  # noqa: E402

OK = "  [OK]  "
BAD = "  [--]  "


def decode_roles(token: str) -> list[str]:
    payload = token.split(".")[1]
    payload += "=" * (-len(payload) % 4)
    return json.loads(base64.urlsafe_b64decode(payload)).get("roles") or []


def main() -> int:
    print("\nApp-only (unattended) Graph check")
    print("=" * 62)

    if not settings.has_graph_credentials:
        print(f"{BAD}No client secret found. Check .env.")
        return 1

    try:
        client = GraphClient(prefer="app")
        token = client._acquire()  # noqa: SLF001 - deliberate preflight introspection
    except GraphError as exc:
        print(f"{BAD}Could not get an app token: {exc}")
        return 1

    roles = decode_roles(token)
    if roles:
        print(f"{OK}Application permissions granted: {', '.join(roles)}")
    else:
        print(f"{BAD}No application permissions on the token (roles is empty).")
        print("         Admin consent has not been granted yet, so the app still")
        print("         cannot read SharePoint by itself.")

    print("-" * 62)
    failures = 0
    for source in SOURCES:
        try:
            meta = client.item_metadata(source.url)
            size_kb = round((meta.get("size") or 0) / 1024)
            print(f"{OK}{source.label}")
            print(f"         modified {meta.get('lastModifiedDateTime')}  ({size_kb} KB)")
        except Exception as exc:  # noqa: BLE001 - report, never raise
            failures += 1
            required = "required" if source.required else "optional"
            print(f"{BAD}{source.label}  [{required}]")
            print(f"         {str(exc)[:220]}")

    print("=" * 62)
    if failures == 0 and roles:
        print("READY. The refresh can run in GitHub Actions without your PC.\n")
        return 0
    print("NOT READY. Grant admin consent, then grant this app access to the")
    print("two site collections listed in ENTRA_SETUP.md, and run this again.\n")
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
