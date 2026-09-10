#!/usr/bin/env bash
# Refresh loop + static server in one container.
set -uo pipefail

refresh() {
  echo "[$(date -u +'%H:%M:%S')] refreshing snapshot..."
  python -m ingest --quiet
  echo "[$(date -u +'%H:%M:%S')] refresh exit=$?"
}

refresh   # build once immediately so the page is never empty

(
  while true; do
    sleep "${REFRESH_SECONDS}"
    refresh
  done
) &

cd /app/docs
echo "serving on :${PORT}"
exec python -m http.server "${PORT}" --bind 0.0.0.0
