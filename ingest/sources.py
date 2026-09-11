"""Workbook parsers.

Each parser turns a downloaded xlsx into tidy pandas DataFrames with normalised
column names, types and isotainer references.

Design note: for the ops tracker we deliberately recompute every derived value in
Python from the raw captured columns rather than trusting the workbook's formula
cache. A file edited in Excel Online and downloaded through Graph may carry stale or
missing cached results, and the report must never show a stale duration.
"""

from __future__ import annotations

import logging
import re
import warnings
from datetime import date, datetime, time, timedelta
from pathlib import Path

import pandas as pd

from .config import (
    ALLOWED_DECANT_HRS,
    BAG_WEIGHT_MT,
    ISO_ALIASES,
    SITE_CONNECT,
    SITE_SAFRIPOL,
    TARGET_LOAD_KG,
)

log = logging.getLogger(__name__)

# GPS data-quality bounds.
#
# A tracker that goes quiet and later resumes in the same zone looks, to a naive
# reader, like one continuous stay. Allowing a little slack absorbs ordinary
# reporting jitter; beyond it the gap is treated as a separate visit.
MAX_PING_GAP_HOURS = 6.0
# An isotainer parked between jobs is not an operational turnaround. Cycles longer
# than this are excluded from turnaround averages and the trend, and counted
# separately so the exclusion is visible rather than silent.
MAX_TURNAROUND_HOURS = 24.0


# --------------------------------------------------------------------------- #
# helpers
# --------------------------------------------------------------------------- #
def _norm_iso(value) -> str | None:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    # Legacy trackers stored the isotainer as a bare number (1..5).
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        n = int(value)
        return f"ISO-{n:02d}" if 1 <= n <= 99 else None
    text = re.sub(r"\s+", " ", str(value)).strip().upper()
    if not text:
        return None
    if re.fullmatch(r"\d{1,2}(\.0)?", text):
        return f"ISO-{int(float(text)):02d}"
    if text in ISO_ALIASES:
        return ISO_ALIASES[text]
    text2 = text.replace("ISOT", "ISO").replace("  ", " ")
    if text2 in ISO_ALIASES:
        return ISO_ALIASES[text2]
    m = re.match(r"^ISO[\s\-_]*0*(\d{1,2})$", text2)
    if m:
        return f"ISO-{int(m.group(1)):02d}"
    return text


def _clean_cols(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [re.sub(r"\s+", " ", str(c)).strip() for c in df.columns]
    return df


def _to_date(value):
    """Parse the mixed date representations found in the operational workbooks."""
    if value is None or value == "" or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)):
        try:  # Excel serial
            return (datetime(1899, 12, 30) + timedelta(days=float(value))).date()
        except Exception:
            return None
    text = str(value).strip()
    if not text:
        return None
    # The legacy tracker holds US-style m/d/yyyy text dates; the new tracker holds real
    # Excel dates. Try slash dates month-first, everything else day-first.
    orders = (False, True) if "/" in text else (True, False)
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        for dayfirst in orders:
            parsed = pd.to_datetime(text, errors="coerce", dayfirst=dayfirst)
            if not pd.isna(parsed):
                return parsed.date()
    return None


def _num(value, default=None):
    if value is None or value == "" or (isinstance(value, float) and pd.isna(value)):
        return default
    if isinstance(value, (int, float)):
        return float(value)
    text = re.sub(r"[^\d\.\-]", "", str(value))
    if text in ("", "-", "."):
        return default
    try:
        return float(text)
    except ValueError:
        return default


def _read(path: Path, sheet: str | int = 0, header: int = 0) -> pd.DataFrame:
    return _clean_cols(pd.read_excel(path, sheet_name=sheet, header=header, engine="openpyxl"))


def _sheet_names(path: Path) -> list[str]:
    import openpyxl

    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    names = list(wb.sheetnames)
    wb.close()
    return names


def _find_sheet(path: Path, *candidates: str) -> str | None:
    names = _sheet_names(path)
    lowered = {n.strip().lower(): n for n in names}
    for c in candidates:
        if c.strip().lower() in lowered:
            return lowered[c.strip().lower()]
    for c in candidates:
        for low, orig in lowered.items():
            if c.strip().lower() in low:
                return orig
    return None


# --------------------------------------------------------------------------- #
# 1. Safripol Stock Report - dispatch, receipts, storage
# --------------------------------------------------------------------------- #
def parse_dispatch(path: Path) -> pd.DataFrame:
    sheet = _find_sheet(path, "3PL Dispatch Detail", "Dispatch Detail", "Dispatch")
    if sheet is None:
        return pd.DataFrame()
    df = _read(path, sheet)
    df = df.loc[:, ~df.columns.str.startswith("Unnamed")]
    if df.empty:
        return df

    ren = {
        "Loaded Weight (KG's)": "loaded_weight_kg",
        "Loaded Weight (KGs)": "loaded_weight_kg",
        "Actual Weight Offloaded": "actual_weight_offloaded_kg",
        "Actual Weight Loaded": "actual_weight_loaded_kg",
        "Remained in Tank (After Offloading)": "remained_in_tank_kg",
        "Isotainer Empty Weight": "iso_empty_weight_kg",
        "Actual Delivery Date": "actual_delivery_date",
        "Bank Release Date": "bank_release_date",
        "Isotainer Ref": "iso",
        "Vessel Name": "vessel",
        "BL Number": "bl_number",
        "Pick Number": "pick_number",
        "Product": "product",
        "Truck Registration/ Drivers Name": "truck_driver",
        "Loaded Weight Var": "loaded_weight_var_kg",
    }
    df = df.rename(columns={k: v for k, v in ren.items() if k in df.columns})
    for col in ("iso",):
        if col in df:
            df[col] = df[col].map(_norm_iso)
    for col in ("actual_delivery_date", "bank_release_date"):
        if col in df:
            df[col] = df[col].map(_to_date)
    for col in ("loaded_weight_kg", "actual_weight_offloaded_kg", "actual_weight_loaded_kg",
                "remained_in_tank_kg", "iso_empty_weight_kg", "loaded_weight_var_kg"):
        if col in df:
            df[col] = df[col].map(lambda v: _num(v))

    # Business rule: recognised delivered tonnage uses Loaded Weight, by design.
    df["offloaded_mt"] = (df.get("loaded_weight_kg", pd.Series(dtype=float)) / 1000).fillna(0)
    df["is_delivered"] = df.get("actual_delivery_date").notna() if "actual_delivery_date" in df else False
    df = df[df.get("iso").notna() | df["is_delivered"]] if "iso" in df else df
    return df.reset_index(drop=True)


def parse_receipts(path: Path) -> pd.DataFrame:
    ren = {
        "Arrival Date": "arrival_date",
        "Total Bags": "total_bags",
        "Damaged Bags": "damaged_bags",
        "Total Received (KG's)": "total_received_kg",
        "Total Received (KGs)": "total_received_kg",
        "Vessel Name": "vessel",
        "BL Reference": "bl_reference",
        "Type": "receipt_type",
        "Load Number": "load_number",
        "Product": "product",
        "Delivery Note": "delivery_note",
        "Warehouse": "warehouse",
    }

    frames = []
    for sheet in (
        _find_sheet(path, "3PL Receipt Detail"),
        _find_sheet(path, "Leasehold Receipt Detail"),
    ):
        if sheet is None:
            continue
        raw = _read(path, sheet)
        raw = raw.loc[:, ~raw.columns.str.startswith("Unnamed")]
        if raw.empty:
            continue
        raw = raw.rename(columns={k: v for k, v in ren.items() if k in raw.columns})
        if "arrival_date" not in raw:
            continue
        if "receipt_type" not in raw:
            raw["receipt_type"] = (
                "Leasehold" if "leasehold" in sheet.lower() else "Direct from Vessel"
            )
        raw["source_sheet"] = sheet
        frames.append(raw)

    if not frames:
        return pd.DataFrame()
    df = pd.concat(frames, ignore_index=True, sort=False)
    df["arrival_date"] = df["arrival_date"].map(_to_date)
    df = df[df["arrival_date"].notna()].copy()
    for col in ("total_bags", "damaged_bags", "total_received_kg"):
        if col in df:
            df[col] = df[col].map(lambda v: _num(v, 0.0))
        else:
            df[col] = 0.0
    df["receipt_type"] = df.get("receipt_type", pd.Series(dtype=str)).astype(str).str.strip()
    df["received_mt_physical"] = df["total_received_kg"] / 1000.0
    df["received_mt_admin"] = df["total_bags"] * BAG_WEIGHT_MT
    return df.reset_index(drop=True)


# --------------------------------------------------------------------------- #
# 2. GPS dwell feed -> operational visits
# --------------------------------------------------------------------------- #
def _duration_hours(value) -> float | None:
    """The dwell feed stores duration as time, timedelta, seconds or '1899-.. hh:mm:ss'."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, timedelta):
        return value.total_seconds() / 3600.0
    if isinstance(value, time):
        return value.hour + value.minute / 60 + value.second / 3600
    if isinstance(value, datetime):
        return value.hour + value.minute / 60 + value.second / 3600
    if isinstance(value, (int, float)):
        v = float(value)
        return v * 24.0 if v <= 3 else v / 3600.0      # excel fraction vs raw seconds
    text = str(value).strip()
    if not text:
        return None
    if " " in text and "1899" in text:
        text = text.split(" ", 1)[1]
    m = re.match(r"^(\d+):(\d{1,2})(?::(\d{1,2}))?$", text)
    if m:
        h, mi, s = int(m.group(1)), int(m.group(2)), int(m.group(3) or 0)
        return h + mi / 60 + s / 3600
    try:
        return float(text) / 3600.0
    except ValueError:
        return None


def parse_dwell_visits(path: Path) -> pd.DataFrame:
    """Collapse raw GPS pings into one row per site visit.

    Mirrors the FactDwells_AllEvents Power Query logic: a new visit starts when the
    zone changes or when the running dwell counter resets downwards, and the visit is
    represented by its terminal (largest dwell) ping.
    """
    df = _read(path, 0)
    if df.empty:
        return df
    ren = {
        "name": "iso", "imei_number": "imei", "zone_name": "site",
        "gps_datetime": "gps_datetime", "zone_max_secs_allowed_in_zone": "zone_max_secs",
        "total_secs_spent_in_zone": "dwell_raw", "is_overstay": "is_overstay",
        "zone_status": "zone_status",
    }
    df = df.rename(columns={k: v for k, v in ren.items() if k in df.columns})
    required = {"iso", "site", "gps_datetime", "dwell_raw"}
    if not required.issubset(df.columns):
        log.warning("Dwell feed missing columns: %s", required - set(df.columns))
        return pd.DataFrame()

    df["iso"] = df["iso"].map(_norm_iso)
    df["site"] = df["site"].astype(str).str.strip()
    df = df[df["site"].isin([SITE_CONNECT, SITE_SAFRIPOL])].copy()
    df["gps_datetime"] = pd.to_datetime(df["gps_datetime"], errors="coerce")
    df = df[df["gps_datetime"].notna()]
    df["dwell_hours"] = df["dwell_raw"].map(_duration_hours).fillna(0.0)
    df["is_overstay_flag"] = (
        df.get("is_overstay", pd.Series(dtype=str)).astype(str).str.strip().str.upper().eq("YES")
        .astype(int)
    )
    df["zone_max_secs"] = df.get("zone_max_secs", pd.Series(dtype=float)).map(lambda v: _num(v))

    df = df.sort_values(["iso", "gps_datetime"]).reset_index(drop=True)
    visits: list[dict] = []
    for iso, grp in df.groupby("iso", sort=False):
        grp = grp.sort_values("gps_datetime")
        visit_id = 0
        prev_site = None
        prev_dwell = None
        prev_time = None
        current: dict | None = None
        for _, row in grp.iterrows():
            # The dwell counter is the authority on continuous presence. If far more
            # wall-clock time has passed than the counter advanced, the unit was not
            # sitting in the zone the whole time - the tracker went quiet (parked,
            # off, out of coverage) and came back. Treating that as one visit merges
            # months into a single "visit" and produces absurd turnaround times.
            gap_hours = (
                (row["gps_datetime"] - prev_time).total_seconds() / 3600.0
                if prev_time is not None else 0.0
            )
            dwell_delta = (
                row["dwell_hours"] - prev_dwell if prev_dwell is not None else 0.0
            )
            stale_gap = (
                prev_time is not None
                and gap_hours > dwell_delta + MAX_PING_GAP_HOURS
            )
            new_visit = (
                prev_site is None
                or row["site"] != prev_site
                or (prev_dwell is not None and row["dwell_hours"] < prev_dwell)
                or stale_gap
            )
            if new_visit:
                if current is not None:
                    visits.append(current)
                visit_id += 1
                current = {
                    "iso": iso,
                    "imei": row.get("imei"),
                    "site": row["site"],
                    "visit_id": visit_id,
                    "visit_start": row["gps_datetime"],
                    "visit_end": row["gps_datetime"],
                    "dwell_hours": row["dwell_hours"],
                    "is_overstay_flag": int(row["is_overstay_flag"]),
                    "zone_max_secs": row["zone_max_secs"],
                }
            else:
                assert current is not None
                if row["dwell_hours"] >= current["dwell_hours"]:
                    current["dwell_hours"] = row["dwell_hours"]
                    current["visit_end"] = row["gps_datetime"]
                current["is_overstay_flag"] = max(
                    current["is_overstay_flag"], int(row["is_overstay_flag"])
                )
            prev_site = row["site"]
            prev_dwell = row["dwell_hours"]
            prev_time = row["gps_datetime"]
        if current is not None:
            visits.append(current)

    out = pd.DataFrame(visits)
    if out.empty:
        return out
    out = out[out["dwell_hours"] > 0].copy()
    out["date"] = out["visit_start"].dt.date
    out["exceeds_site_target"] = out.apply(
        lambda r: None if pd.isna(r["zone_max_secs"])
        else int(r["dwell_hours"] > (r["zone_max_secs"] / 3600.0)),
        axis=1,
    )
    return out.sort_values(["iso", "visit_start"]).reset_index(drop=True)


def build_turnaround(visits: pd.DataFrame) -> pd.DataFrame:
    """Derive a Connect -> transit -> Safripol -> transit -> Connect cycle per isotainer.

    Turnaround is computed from the GPS visit sequence rather than being captured by
    hand, so the yard team only records decanting.
    """
    if visits.empty:
        return pd.DataFrame()
    rows: list[dict] = []
    for iso, grp in visits.groupby("iso", sort=False):
        seq = grp.sort_values("visit_start").to_dict("records")
        for i, v in enumerate(seq):
            if v["site"] != SITE_CONNECT:
                continue
            saf = next((s for s in seq[i + 1:] if s["site"] == SITE_SAFRIPOL), None)
            if saf is None:
                continue
            back = next(
                (s for s in seq
                 if s["site"] == SITE_CONNECT and s["visit_start"] > saf["visit_end"]),
                None,
            )
            depart_connect = v["visit_end"]
            transit_out = (saf["visit_start"] - depart_connect).total_seconds() / 3600.0
            if transit_out < 0 or transit_out > 24:
                continue
            row = {
                "iso": iso,
                "cycle_start": v["visit_start"],
                "connect_dwell_hours": v["dwell_hours"],
                "depart_connect": depart_connect,
                "arrive_safripol": saf["visit_start"],
                "transit_out_hours": transit_out,
                "safripol_dwell_hours": saf["dwell_hours"],
                "depart_safripol": saf["visit_end"],
                "date": v["visit_start"].date(),
            }
            if back is not None:
                ret = (back["visit_start"] - saf["visit_end"]).total_seconds() / 3600.0
                if 0 <= ret <= 24:
                    turnaround = (
                        (back["visit_start"] - v["visit_start"]).total_seconds() / 3600.0
                    )
                    row["return_connect"] = back["visit_start"]
                    row["transit_return_hours"] = ret
                    row["turnaround_hours"] = turnaround
                    # Both legs can be sane while the cycle is not: the unit stood at
                    # Connect for days before departing. Flag rather than drop, so the
                    # exception stays auditable.
                    row["turnaround_excluded"] = int(turnaround > MAX_TURNAROUND_HOURS)
            rows.append(row)
    return pd.DataFrame(rows)


# --------------------------------------------------------------------------- #
# 3. Ops tracker
# --------------------------------------------------------------------------- #
def _shift_datetime(value, fallback_date):
    """Accept a full datetime, or a time-only value anchored to the shift date."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, datetime):
        return value
    if isinstance(value, time) and fallback_date is not None:
        return datetime.combine(fallback_date, value)
    if isinstance(value, str):
        parsed = pd.to_datetime(value, errors="coerce", dayfirst=True)
        if not pd.isna(parsed):
            return parsed.to_pydatetime()
    if isinstance(value, (int, float)) and fallback_date is not None:
        frac = float(value)
        if 0 <= frac < 1:
            return datetime.combine(fallback_date, time(0)) + timedelta(days=frac)
        try:
            return datetime(1899, 12, 30) + timedelta(days=frac)
        except Exception:
            return None
    return None


def parse_decant_log(path: Path) -> pd.DataFrame:
    """Read the Decant Log and recompute all derived fields in Python.

    Also tolerates the legacy tracker layout (the previous shipment's sheet) so the
    report still works if the old file is pointed at by mistake.
    """
    sheet = _find_sheet(path, "Decant Log", "New Version", "Old Version")
    if sheet is None:
        return pd.DataFrame()
    raw = _read(path, sheet)
    if raw.empty:
        return raw

    cols = {c.lower().strip(): c for c in raw.columns}

    def col(*names):
        for n in names:
            if n.lower() in cols:
                return cols[n.lower()]
        return None

    c_date = col("Shift Date", "Date")
    c_shift = col("Shift")
    c_team = col("Team")
    c_heads = col("Heads on Shift", "Heads")
    c_iso = col("Isotainer")
    c_start = col("Start Date & Time", "Start Time", "Start")
    c_end = col("End Date & Time", "End Time", "End")
    c_tare = col("Tare Weight (kg)", "Tare Weight")
    c_empty = col("Empty Weight (kg)", "Empty Weight")
    c_loaded = col("Loaded Weight (kg)", "Loaded Weight")
    c_delay = col("Delay Minutes")
    c_cat = col("Delay Category")
    c_owner = col("Delay Owner")
    c_detail = col("Delay Detail")
    c_esc = col("Escalate?")
    c_comment = col("Comments")

    if c_date is None or c_start is None:
        return pd.DataFrame()

    rows: list[dict] = []
    for _, r in raw.iterrows():
        shift_date = _to_date(r.get(c_date))
        if shift_date is None:
            continue
        start = _shift_datetime(r.get(c_start), shift_date)
        end = _shift_datetime(r.get(c_end), shift_date)
        if start is None:
            continue
        # nightshift roll-over: an end before the start belongs to the next calendar day
        if end is not None and end < start and (start - end) < timedelta(hours=20):
            end = end + timedelta(days=1)

        decant_hours = (end - start).total_seconds() / 3600.0 if end else None
        if decant_hours is not None and (decant_hours <= 0 or decant_hours > 24):
            decant_hours = None

        delay_mins = _num(r.get(c_delay) if c_delay else None, 0.0) or 0.0
        loaded = _num(r.get(c_loaded) if c_loaded else None)
        empty = _num(r.get(c_empty) if c_empty else None)
        tare = _num(r.get(c_tare) if c_tare else None)
        shift = str(r.get(c_shift) or "").strip() if c_shift else ""

        rows.append({
            "shift_date": shift_date,
            "shift": shift or "Unspecified",
            "team": str(r.get(c_team) or "").strip() if c_team else "",
            "heads": _num(r.get(c_heads) if c_heads else None),
            "iso": _norm_iso(r.get(c_iso)) if c_iso else None,
            "start": start,
            "end": end,
            "decant_hours": decant_hours,
            "allowed_hours": ALLOWED_DECANT_HRS,
            "time_variance_hours": (decant_hours - ALLOWED_DECANT_HRS)
            if decant_hours is not None else None,
            "delay_minutes": delay_mins,
            "delay_category": (str(r.get(c_cat)).strip() if c_cat and r.get(c_cat) else "None"),
            "delay_owner": (str(r.get(c_owner)).strip() if c_owner and r.get(c_owner) else ""),
            "delay_detail": (str(r.get(c_detail)).strip() if c_detail and r.get(c_detail) else ""),
            "escalate": (str(r.get(c_esc)).strip().lower() == "yes") if c_esc else False,
            "comments": (str(r.get(c_comment)).strip() if c_comment and r.get(c_comment) else ""),
            "tare_kg": tare,
            "empty_kg": empty,
            "loaded_kg": loaded,
            "carry_over_kg": (empty - tare) if (empty is not None and tare is not None) else None,
            "excess_kg": (loaded - TARGET_LOAD_KG) if loaded is not None else None,
            "net_product_kg": (loaded - empty)
            if (loaded is not None and empty is not None) else None,
        })

    df = pd.DataFrame(rows)
    if df.empty:
        return df
    df["net_product_mt"] = df["net_product_kg"] / 1000.0
    df["productive_hours"] = (df["decant_hours"] - df["delay_minutes"] / 60.0).clip(lower=0)
    df["performance"] = df["time_variance_hours"].map(
        lambda v: None if pd.isna(v) else ("Clean" if v <= 0 else "Time Exceeded")
    )
    df["shift_key"] = df["shift_date"].astype(str) + "|" + df["shift"]
    df = df.sort_values("start").reset_index(drop=True)

    # Decant interval = gap between the end of the previous decant and this start,
    # within the same shift. Computed here so it never depends on a formula cache.
    df["decant_interval_hours"] = pd.NA
    for _, grp in df.groupby("shift_key", sort=False):
        idx = grp.index.tolist()
        for i in range(1, len(idx)):
            prev_end = df.at[idx[i - 1], "end"]
            this_start = df.at[idx[i], "start"]
            if prev_end is not None and this_start is not None and this_start > prev_end:
                gap = (this_start - prev_end).total_seconds() / 3600.0
                if 0 <= gap <= 24:
                    df.at[idx[i], "decant_interval_hours"] = gap
    df["decant_interval_hours"] = pd.to_numeric(df["decant_interval_hours"], errors="coerce")
    return df


def parse_shift_plan(path: Path) -> pd.DataFrame:
    sheet = _find_sheet(path, "Shift Plan")
    if sheet is None:
        return pd.DataFrame()
    df = _read(path, sheet)
    if df.empty:
        return df
    ren = {"Shift Date": "shift_date", "Shift": "shift", "Team": "team",
           "Planned Heads": "planned_heads", "Planned Isotainers": "planned_isotainers",
           "Shift Notes": "notes"}
    df = df.rename(columns={k: v for k, v in ren.items() if k in df.columns})
    if "shift_date" not in df:
        return pd.DataFrame()
    df["shift_date"] = df["shift_date"].map(_to_date)
    df = df[df["shift_date"].notna()].copy()
    for c in ("planned_heads", "planned_isotainers"):
        df[c] = df.get(c, pd.Series(dtype=float)).map(lambda v: _num(v))
    df["shift"] = df.get("shift", pd.Series(dtype=str)).astype(str).str.strip()
    df["shift_key"] = df["shift_date"].astype(str) + "|" + df["shift"]
    keep = ["shift_date", "shift", "team", "planned_heads", "planned_isotainers",
            "shift_key", "notes"]
    return df[[c for c in keep if c in df.columns]].reset_index(drop=True)


def parse_delay_log(path: Path) -> pd.DataFrame:
    sheet = _find_sheet(path, "Delay Log")
    if sheet is None:
        return pd.DataFrame()
    df = _read(path, sheet)
    if df.empty:
        return df
    ren = {"Shift Date": "shift_date", "Shift": "shift", "Team": "team", "Area": "area",
           "Start Date & Time": "start", "End Date & Time": "end",
           "Delay Category": "delay_category", "Delay Owner": "delay_owner",
           "Isotainer Affected": "iso", "Description": "description",
           "Action Taken": "action", "Escalate?": "escalate", "Resolved?": "resolved",
           "Escalated To": "escalated_to"}
    df = df.rename(columns={k: v for k, v in ren.items() if k in df.columns})
    if "shift_date" not in df:
        return pd.DataFrame()
    df["shift_date"] = df["shift_date"].map(_to_date)
    df = df[df["shift_date"].notna()].copy()
    if df.empty:
        return df
    df["start"] = [
        _shift_datetime(s, d) for s, d in zip(df.get("start"), df["shift_date"])
    ]
    df["end"] = [
        _shift_datetime(e, d) for e, d in zip(df.get("end"), df["shift_date"])
    ]
    def mins(row):
        if row["start"] and row["end"]:
            e = row["end"]
            if e < row["start"]:
                e = e + timedelta(days=1)
            return (e - row["start"]).total_seconds() / 60.0
        return 0.0
    df["delay_minutes"] = df.apply(mins, axis=1)
    for c in ("escalate", "resolved"):
        df[c] = df.get(c, pd.Series(dtype=str)).astype(str).str.strip().str.lower().eq("yes")
    df["iso"] = df.get("iso", pd.Series(dtype=str)).map(_norm_iso)
    df["shift"] = df.get("shift", pd.Series(dtype=str)).astype(str).str.strip()
    df["shift_key"] = df["shift_date"].astype(str) + "|" + df["shift"]
    return df.reset_index(drop=True)


def _find_header_row(path: Path, sheet: str, must_have: str, limit: int = 12) -> int:
    """Locate the real header row in a sheet that opens with title/banner rows."""
    probe = pd.read_excel(path, sheet_name=sheet, header=None, nrows=limit,
                          engine="openpyxl")
    want = must_have.strip().lower()
    for i in range(len(probe)):
        cells = [str(c).strip().lower() for c in probe.iloc[i].tolist()]
        if want in cells:
            return i
    return 0


def parse_drawdown_plan(path: Path) -> pd.DataFrame:
    """Daily drawdown plan.

    Handles the standalone PTA_Drawdown_Plan workbook as well as the older
    'Drawdown Plan' sheet embedded in the ops tracker, so either can supply it.
    """
    sheet = _find_sheet(path, "Daily Drawdown Plan", "Drawdown Plan")
    if sheet is None:
        return pd.DataFrame()
    df = _read(path, sheet, header=_find_header_row(path, sheet, "Date"))
    df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]
    if df.empty:
        return df
    ren = {
        "Date": "date",
        "Daily Target (t)": "planned_mt",
        "Planned MT": "planned_mt",
        "Isotainers to Schedule": "planned_isotainers",
        "Planned Isotainers": "planned_isotainers",
        "Isotainers Req'd (exact)": "planned_isotainers_exact",
        "Cumulative Tons": "cum_planned_mt",
        "Cumulative Isotainers": "cum_planned_isotainers",
        "Balance Remaining (t)": "planned_balance_mt",
        "Day": "weekday",
        "Notes": "notes",
    }
    df = df.rename(columns={k: v for k, v in ren.items() if k in df.columns})
    if "date" not in df:
        return pd.DataFrame()
    df["date"] = df["date"].map(_to_date)
    df = df[df["date"].notna()].copy()
    if df.empty:
        return df

    df["planned_isotainers"] = df.get(
        "planned_isotainers", pd.Series([None] * len(df))).map(lambda v: _num(v, 0.0))
    df["planned_mt"] = [
        p if p is not None else (i or 0) * TARGET_LOAD_KG / 1000.0
        for p, i in zip(df.get("planned_mt", pd.Series([None] * len(df))).map(lambda v: _num(v)),
                        df["planned_isotainers"])
    ]
    df = df.sort_values("date").reset_index(drop=True)
    # Trust the workbook's own cumulative column when present; it is what the
    # ops team reads off the plan.
    if "cum_planned_mt" in df:
        df["cum_planned_mt"] = df["cum_planned_mt"].map(lambda v: _num(v))
    if "cum_planned_mt" not in df or df["cum_planned_mt"].isna().any():
        df["cum_planned_mt"] = df["planned_mt"].cumsum()
    if "cum_planned_isotainers" in df:
        df["cum_planned_isotainers"] = df["cum_planned_isotainers"].map(
            lambda v: _num(v)).round(0)
    else:
        df["cum_planned_isotainers"] = df["planned_isotainers"].cumsum().round(0)
    return df


def parse_plan_assumptions(path: Path) -> dict:
    """Key assumptions and phase table from the drawdown plan Dashboard sheet."""
    out: dict = {"phases": []}
    sheet = _find_sheet(path, "Dashboard")
    if sheet is None:
        return out
    raw = pd.read_excel(path, sheet_name=sheet, header=None, engine="openpyxl")

    labels = {
        "total pta volume": "total_mt",
        "isotainer capacity": "iso_capacity_mt",
        "plan start date": "plan_start",
        "plan end date": "plan_end",
        "total plan duration": "duration_days",
        "total isotainers require": "total_isotainers",
        "average isotainers / day": "avg_isotainers_per_day",
        "average tons / day": "avg_mt_per_day",
    }
    for row in raw.itertuples(index=False):
        cells = list(row)
        head = str(cells[0]).strip().lower() if cells and cells[0] is not None else ""
        if not head:
            continue
        for prefix, key in labels.items():
            if head.startswith(prefix):
                val = next((c for c in cells[1:] if c is not None and str(c).strip() != ""),
                           None)
                if key in ("plan_start", "plan_end"):
                    d = _to_date(val)
                    out[key] = d.isoformat() if d else None
                else:
                    out[key] = _num(val)
                break

    # Phase table sits to the right of the assumptions, under a 'Phase' header.
    hdr = None
    for i in range(len(raw)):
        cells = [str(c).strip().lower() for c in raw.iloc[i].tolist()]
        if "phase" in cells and "start date" in cells:
            hdr = i
            break
    if hdr is not None:
        cols = {str(c).strip().lower(): j for j, c in enumerate(raw.iloc[hdr].tolist())}
        def cell(r, name):
            j = cols.get(name)
            return r[j] if j is not None and j < len(r) else None
        for i in range(hdr + 1, len(raw)):
            r = raw.iloc[i].tolist()
            name = cell(r, "phase")
            if name is None or str(name).strip() == "":
                continue
            label = str(name).strip()
            if label.lower() == "total":
                break
            start, end = _to_date(cell(r, "start date")), _to_date(cell(r, "end date"))
            out["phases"].append({
                "phase": label,
                "start": start.isoformat() if start else None,
                "end": end.isoformat() if end else None,
                "days": _num(cell(r, "days")),
                "daily_target_mt": _num(cell(r, "daily target (t)")),
                "isotainers_per_day": _num(cell(r, "isotainers / day")),
                "phase_total_mt": _num(cell(r, "phase total (t)")),
            })
    return out


def parse_staff(path: Path) -> pd.DataFrame:
    sheet = _find_sheet(path, "Staff", "SAF Team")
    if sheet is None:
        return pd.DataFrame()
    df = _read(path, sheet)
    if df.empty:
        return df
    ren = {"Name": "name", "Surname": "surname", "Company": "company",
           "Position": "position", "Department": "department", "Team": "team",
           "Hours Per Day": "hours_per_day", "Allocation": "allocation",
           "Comments": "allocation", "Trained Y/N": "trained",
           "Worked Previous Shipment": "worked_previous", "Active Y/N": "active"}
    df = df.rename(columns={k: v for k, v in ren.items() if k in df.columns})
    if "name" not in df:
        return pd.DataFrame()
    df = df[df["name"].notna()].copy()
    df = df[df["name"].astype(str).str.strip().str.lower() != "name"]
    for c in ("name", "surname", "company", "position", "department", "team",
              "allocation", "trained", "worked_previous"):
        if c in df:
            df[c] = (df[c].astype(str)
                     .str.replace("\xa0", " ", regex=False)
                     .str.replace(r"\s+", " ", regex=True)
                     .str.strip())
    df["hours_per_day"] = df.get("hours_per_day", pd.Series(dtype=float)).map(
        lambda v: _num(v, 0.0))
    if "active" in df:
        df = df[~df["active"].astype(str).str.strip().str.lower().eq("no")]
    return df.reset_index(drop=True)
