"""Cache-bust the front-end assets referenced by docs/index.html.

GitHub Pages serves assets/app.js and assets/styles.css with a long cache
lifetime, so a returning visitor can keep running an old build for hours after
a deploy - the report looks unchanged even though the data is current. Stamping
each reference with a short content hash makes the URL change whenever the file
changes, which forces a fetch of the new build and nothing more.

Run after editing anything under docs/assets.
"""

from __future__ import annotations

import hashlib
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
INDEX = ROOT / "docs" / "index.html"
ASSETS = ("assets/app.js", "assets/styles.css")


def digest(rel: str) -> str:
    path = ROOT / "docs" / rel
    return hashlib.sha1(path.read_bytes()).hexdigest()[:8]


def main() -> int:
    if not INDEX.exists():
        print("index.html not found", file=sys.stderr)
        return 1

    html = original = INDEX.read_text(encoding="utf-8")
    for rel in ASSETS:
        h = digest(rel)
        # Match the reference with or without an existing ?v= stamp.
        pattern = re.compile(rf'({re.escape(rel)})(\?v=[0-9a-f]+)?')
        html, n = pattern.subn(rf'\1?v={h}', html)
        if not n:
            print(f"warning: no reference to {rel} in index.html", file=sys.stderr)
        else:
            print(f"  {rel} -> ?v={h}")

    if html != original:
        INDEX.write_text(html, encoding="utf-8")
        print("index.html stamped")
    else:
        print("index.html already current")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
