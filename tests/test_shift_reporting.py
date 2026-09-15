from datetime import date, datetime

from ingest.sources import _operational_shift


def test_operational_shift_boundaries():
    assert _operational_shift(datetime(2026, 9, 14, 6, 0)) == (
        date(2026, 9, 14), "Dayshift"
    )
    assert _operational_shift(datetime(2026, 9, 14, 17, 59, 59)) == (
        date(2026, 9, 14), "Dayshift"
    )
    assert _operational_shift(datetime(2026, 9, 14, 18, 0)) == (
        date(2026, 9, 14), "Nightshift"
    )
    assert _operational_shift(datetime(2026, 9, 15, 5, 59, 59)) == (
        date(2026, 9, 14), "Nightshift"
    )
