import json

from tools.refresh_and_publish import regressed_sources


def snapshot(source_time):
    return json.dumps({
        "meta": {
            "sources": [{
                "key": "tracking",
                "ok": True,
                "last_modified": source_time,
            }]
        }
    }).encode()


def test_detects_source_timestamp_regression(tmp_path):
    current = tmp_path / "snapshot.json"
    current.write_bytes(snapshot("2026-09-15T15:31:00+02:00"))

    assert regressed_sources(
        snapshot("2026-09-16T02:39:00+02:00"), current
    ) == ["tracking (2026-09-15T15:31:00+02:00 < 2026-09-16T02:39:00+02:00)"]


def test_allows_same_or_newer_source_timestamp(tmp_path):
    current = tmp_path / "snapshot.json"
    current.write_bytes(snapshot("2026-09-16T02:39:00+02:00"))

    assert regressed_sources(
        snapshot("2026-09-16T02:39:00+02:00"), current
    ) == []
