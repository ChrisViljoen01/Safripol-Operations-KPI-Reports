"""Builds the snapshot JSON that the live HTML report consumes."""

from __future__ import annotations

import json
import logging
from datetime import datetime, timezone
from pathlib import Path

import pandas as pd

from . import metrics, sources
from .config import (
    DATA_DIR,
    SAST,
    SHIPMENT,
    SOURCES,
    TARGET_TOTAL_MT,
    VESSEL_LABEL,
)
from .graph import Fetched, fetch_all

log = logging.getLogger(__name__)

SNAPSHOT_VERSION = 1


def _safe(fn, *args, label="", default=None):
    try:
        return fn(*args)
    except Exception as exc:  # noqa: BLE001 - a bad sheet must not kill the refresh
        log.warning("Parser failed for %s: %s", label or fn.__name__, exc)
        return default if default is not None else pd.DataFrame()


def build_snapshot() -> dict:
    fetched: dict[str, Fetched] = fetch_all(SOURCES)

    stock = fetched.get("stock")
    dwells = fetched.get("dwells")
    tracking = fetched.get("tracking")
    staff_src = fetched.get("staff")
    plan_src = fetched.get("plan")

    dispatch = receipts = pd.DataFrame()
    if stock and stock.ok:
        dispatch = _safe(sources.parse_dispatch, stock.path, label="dispatch")
        receipts = _safe(sources.parse_receipts, stock.path, label="receipts")

    visits = turns = pd.DataFrame()
    if dwells and dwells.ok:
        visits = _safe(sources.parse_dwell_visits, dwells.path, label="dwells")
        if not visits.empty:
            turns = _safe(sources.build_turnaround, visits, label="turnaround")

    decant = shift_plan = delay_log = drawdown = pd.DataFrame()
    assumptions: dict = {}
    if tracking and tracking.ok:
        decant = _safe(sources.parse_decant_log, tracking.path, label="decant log")
        shift_plan = _safe(sources.parse_shift_plan, tracking.path, label="shift plan")
        delay_log = _safe(sources.parse_delay_log, tracking.path, label="delay log")
        drawdown = _safe(sources.parse_drawdown_plan, tracking.path, label="drawdown")
    # The standalone plan workbook is the authority; it wins over any older
    # copy of the plan embedded in the tracker.
    if plan_src and plan_src.ok:
        standalone = _safe(sources.parse_drawdown_plan, plan_src.path, label="drawdown plan")
        if not standalone.empty:
            drawdown = standalone
        assumptions = _safe(sources.parse_plan_assumptions, plan_src.path,
                            label="plan assumptions", default={}) or {}

    staff = pd.DataFrame()
    if tracking and tracking.ok:
        staff = _safe(sources.parse_staff, tracking.path, label="staff (tracker)")
    if staff.empty and staff_src and staff_src.ok:
        staff = _safe(sources.parse_staff, staff_src.path, label="staff")

    vessel = metrics.vessel_block()
    rec = metrics.receipts_block(receipts)
    disp = metrics.dispatch_block(dispatch, rec["total_physical_mt"] or 0.0, drawdown)
    dwl = metrics.dwell_block(visits, turns)
    dec = metrics.decant_block(decant, delay_log, shift_plan, staff)
    pln = metrics.plan_block(drawdown, disp["by_date"], disp["delivered_mt"] or 0.0,
                             TARGET_TOTAL_MT, assumptions)
    exc = metrics.exceptions_block(decant, delay_log, dwl, disp, dec)

    now_utc = datetime.now(timezone.utc)
    now_sast = now_utc.astimezone(SAST)

    source_status = []
    for s in SOURCES:
        f = fetched.get(s.key)
        source_status.append({
            "key": s.key,
            "label": s.label,
            "ok": bool(f and f.ok),
            "origin": f.origin if f else "none",
            "last_modified": f.last_modified.astimezone(SAST).isoformat(timespec="minutes")
            if (f and f.last_modified) else None,
            "size_kb": round(f.size / 1024) if f else 0,
            "required": s.required,
            "error": f.error if f else "unavailable",
        })

    headline = {
        "target_total_mt": TARGET_TOTAL_MT,
        "delivered_mt": disp["delivered_mt"],
        "outstanding_mt": disp["outstanding_mt"],
        "completion_pct": disp["completion_pct"],
        "pta_balance_mt": disp["pta_balance_mt"],
        "received_admin_mt": rec["total_admin_mt"],
        "received_physical_mt": rec["total_physical_mt"],
        "loads_total": disp["loads_total"],
        "avg_daily_offtake_mt": disp["avg_daily_offtake_mt"],
        "avg_loads_per_day": disp["avg_loads_per_day"],
        "avg_mt_offloaded_per_tank": disp["avg_mt_offloaded_per_tank"],
        "avg_remained_in_tank_kg": disp["avg_remained_in_tank_kg"],
        "required_rate_mt_per_day": disp["required_rate_mt_per_day"],
        "projected_completion": disp["projected_completion"],
        "projected_remaining_days": disp["projected_remaining_days"],
        "avg_decant_hours": dec["avg_decant_hours"],
        "avg_interval_hours": dec["avg_interval_hours"],
        "time_lost_hours": dec.get("total_delay_hours") or dec["time_lost_hours"],
        "decants_per_day": dec["decants_per_day"],
        "avg_attainment_pct": (dec.get("shift_summary") or {}).get("avg_attainment_pct"),
        "avg_turnaround_hours": (dwl.get("turnaround") or {}).get("avg_turnaround_hours"),
        "open_exceptions": len([e for e in exc if e["severity"] == "high"]),
    }

    # The plan block forecasts off recent pace and knows the plan end date, so it
    # is the better projection. Keep one number across the whole report.
    _proj = pln.get("projection") or {}
    if _proj.get("projected_end"):
        headline["projected_completion"] = _proj["projected_end"]
        headline["projected_remaining_days"] = _proj["days_to_go"]
    if _proj.get("required_rate_mt_per_day") is not None:
        headline["required_rate_mt_per_day"] = _proj["required_rate_mt_per_day"]
    headline["planned_completion"] = pln.get("plan_end")
    headline["days_vs_plan"] = _proj.get("days_vs_plan")
    headline["plan_status"] = _proj.get("status")

    return {
        "schema_version": SNAPSHOT_VERSION,
        "meta": {
            "shipment": SHIPMENT,
            "vessel": VESSEL_LABEL,
            "client": "Safripol",
            "operator": "Connect Logistics",
            "product": "PTA",
            "generated_at_utc": now_utc.isoformat(timespec="seconds"),
            "generated_at_sast": now_sast.isoformat(timespec="seconds"),
            "generated_display": now_sast.strftime("%d %b %Y %H:%M") + " SAST",
            "sources": source_status,
            "degraded": any(not s["ok"] and s["required"] for s in source_status),
        },
        "headline": headline,
        "vessel_discharge": vessel,
        "receipts": rec,
        "dispatch": disp,
        "dwell": dwl,
        "decant": dec,
        "plan": pln,
        "exceptions": exc,
    }


def write_snapshot(snapshot: dict, out_dir: Path | None = None) -> Path:
    out_dir = out_dir or DATA_DIR
    out_dir.mkdir(parents=True, exist_ok=True)
    path = out_dir / "snapshot.json"
    path.write_text(json.dumps(snapshot, indent=2, default=str), encoding="utf-8")

    # tiny file the page polls to detect a change without downloading the payload
    (out_dir / "version.json").write_text(
        json.dumps({
            "generated_at_utc": snapshot["meta"]["generated_at_utc"],
            "generated_display": snapshot["meta"]["generated_display"],
            "degraded": snapshot["meta"]["degraded"],
        }, indent=2),
        encoding="utf-8",
    )
    return path
