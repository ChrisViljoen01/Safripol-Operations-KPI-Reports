# Self-hosted mode: runs the ingest on a timer and serves docs/ over HTTP.
# Only needed if the report is hosted internally instead of on GitHub Pages.
FROM python:3.12-slim

ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    REFRESH_SECONDS=900 \
    PORT=8080

WORKDIR /app

RUN apt-get update \
 && apt-get install -y --no-install-recommends curl \
 && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY ingest/ ./ingest/
COPY docs/ ./docs/
COPY docker/entrypoint.sh /usr/local/bin/entrypoint.sh
RUN chmod +x /usr/local/bin/entrypoint.sh

EXPOSE 8080

HEALTHCHECK --interval=60s --timeout=5s --start-period=30s --retries=3 \
  CMD curl -fsS "http://localhost:${PORT}/data/version.json" > /dev/null || exit 1

ENTRYPOINT ["/usr/local/bin/entrypoint.sh"]
