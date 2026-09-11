"""Metrics engine.

Ports the Power BI DAX model to Python and adds the operational KPIs requested for
the PTA outbound drawdown (decant rate, shift attainment, time lost by reason,
turnaround against target, exceptions).

Business rules preserved from the Power BI master scope:
  * Recognised delivered tonnage uses Loaded Weight (KG's) / 1000 - deliberate.
  * Vessel discharge outturn comes from the hatch table (~20,446.47 MT).
  * Safripol delivery target is 20,390.40 MT (16,992 bags x 1.2 MT).
  * Rows with a blank Actual Delivery Date are NOT completed deliveries.
  * PTA Balance (physical Connect stock) and Shipment Outstanding (customer
    obligation) are different measures and must not be conflated.
"""

from __future__ import annotations

import math
from collections import Counter, defaultdict
from datetime import date, datetime, timedelta

import pandas as pd

from .config import (
    ALLOWED_DECANT_HRS,
    ALLOWED_INTERVAL_HRS,
    BAG_WEIGHT_MT,
    HATCHES,
    SITE_CONNECT,
    SITE_SAFRIPOL,
    SITE_TARGET_HRS,
    TARGET_LOAD_KG,
    TARGET_TOTAL_MT,
    TRANSIT_TARGET_HRS,
    WEATHER_EVENTS,
)


# --------------------------------------------------------------------------- #
# small helpers
# --------------------------------------------------------------------------- #
def r2(v, nd=2):
    if v is None:
        return None
    try:
        f = float(v)
    except (TypeError, ValueError):
        return None
    if math.isnan(f) or math.isinf(f):
        return None
    return round(f, nd)


def iso_d(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    if isinstance(v, datetime):
        return v.date().isoformat()
    if isinstance(v, date):
        return v.isoformat()
    return str(v)


def iso_dt(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    if isinstance(v, datetime):
        return v.isoformat(timespec="minutes")
    if isinstance(v, pd.Timestamp):
        return v.to_pydatetime().isoformat(timespec="minutes")
    return str(v)


def _mean(series) -> float | None:
    s = pd.to_numeric(pd.Series(series), errors="coerce").dropna()
    return float(s.mean()) if len(s) else None


def _pct(part, whole):
    if not whole:
        return 0.0
    return float(part) / float(whole)


# --------------------------------------------------------------------------- #
# vessel
# --------------------------------------------------------------------------- #
def vessel_block() -> dict:
    hatches = []
    for h, s, e, vol, gross, down, weather, net in HATCHES:
        hatches.append({
            "hatch": h, "start": s, "end": e, "volume_mt": vol,
            "gross_hours": gross, "downtime_hours": down,
            "weather_hours": weather, "net_hours": net,
            "rate_mt_per_net_hour": r2(vol / net if net else None),
        })
    starts = [datetime.fromisoformat(h["start"]) for h in hatches]
    ends = [datetime.fromisoformat(h["end"]) for h in hatches]
    start, end = min(starts), max(ends)
    total_mt = sum(h["volume_mt"] for h in hatches)
    minutes = (end - start).total_seconds() / 60.0
    elapsed_days = minutes / 1440.0
    days = int(minutes // 1440)
    rem = int(minutes % 1440)

    weather = [{
        "type": t, "event": n, "start": s, "end": e, "duration_hours": d
    } for t, n, s, e, d in WEATHER_EVENTS]

    return {
        "start": start.isoformat(timespec="minutes"),
        "end": end.isoformat(timespec="minutes"),
        "duration_display": f"{days} days {rem // 60} hours {rem % 60} min",
        "total_discharged_mt": r2(total_mt),
        "avg_mt_per_day": r2(total_mt / elapsed_days if elapsed_days else 0),
        "total_downtime_hours": r2(sum(h["downtime_hours"] for h in hatches)),
        "total_weather_hours": r2(sum(w["duration_hours"] for w in weather)),
        "hatches": hatches,
        "weather": weather,
    }


# --------------------------------------------------------------------------- #
# receipts
# --------------------------------------------------------------------------- #
def receipts_block(receipts: pd.DataFrame) -> dict:
    if receipts is None or receipts.empty:
        return {"total_admin_mt": 0, "total_physical_mt": 0, "direct_mt": 0,
                "leasehold_mt": 0, "total_bags": 0, "damaged_bags": 0,
                "receipt_days": 0, "avg_mt_per_receipt_day": 0, "by_date": []}

    def type_mt(pattern: str) -> float:
        mask = receipts["receipt_type"].str.contains(pattern, case=False, na=False)
        return float(receipts.loc[mask, "received_mt_admin"].sum())

    by_date = (receipts.groupby("arrival_date")
               .agg(bags=("total_bags", "sum"),
                    admin_mt=("received_mt_admin", "sum"),
                    physical_mt=("received_mt_physical", "sum"))
               .reset_index().sort_values("arrival_date"))
    by_date["cum_admin_mt"] = by_date["admin_mt"].cumsum()

    return {
        "total_admin_mt": r2(receipts["received_mt_admin"].sum()),
        "total_physical_mt": r2(receipts["received_mt_physical"].sum()),
        "direct_mt": r2(type_mt("direct")),
        "leasehold_mt": r2(type_mt("leasehold")),
        "total_bags": int(receipts["total_bags"].sum()),
        "damaged_bags": int(pd.to_numeric(receipts["damaged_bags"], errors="coerce")
                            .fillna(0).sum()),
        "receipt_days": int(receipts["arrival_date"].nunique()),
        "avg_mt_per_receipt_day": r2(
            receipts["received_mt_physical"].sum() / max(receipts["arrival_date"].nunique(), 1)),
        "first_arrival": iso_d(receipts["arrival_date"].min()),
        "last_arrival": iso_d(receipts["arrival_date"].max()),
        "by_date": [{
            "date": iso_d(r.arrival_date), "bags": int(r.bags),
            "admin_mt": r2(r.admin_mt), "physical_mt": r2(r.physical_mt),
            "cum_admin_mt": r2(r.cum_admin_mt),
        } for r in by_date.itertuples()],
    }


# --------------------------------------------------------------------------- #
# dispatch / delivery to Safripol
# --------------------------------------------------------------------------- #
def dispatch_block(dispatch: pd.DataFrame, receipts_physical_mt: float,
                   drawdown: pd.DataFrame) -> dict:
    empty = {
        "delivered_mt": 0.0, "outstanding_mt": r2(TARGET_TOTAL_MT),
        "completion_pct": 0.0, "target_total_mt": TARGET_TOTAL_MT,
        "loads_total": 0, "avg_loads_per_day": 0.0, "max_loads_day": 0,
        "min_loads_day": 0, "avg_daily_offtake_mt": 0.0,
        "avg_mt_per_load": 0.0, "avg_mt_offloaded_per_tank": 0.0,
        "pta_balance_mt": r2(max(receipts_physical_mt, 0)),
        "projected_completion": "Pending", "projected_remaining_days": None,
        "required_rate_mt_per_day": None, "avg_remained_in_tank_kg": None,
        "tolerance_pct": 0.0, "variance_mt": 0.0,
        "first_delivery": None, "last_delivery": None,
        "by_date": [], "by_iso": [], "open_loads": 0,
    }
    if dispatch is None or dispatch.empty or "actual_delivery_date" not in dispatch:
        return empty

    done = dispatch[dispatch["actual_delivery_date"].notna()].copy()
    open_loads = int(len(dispatch) - len(done))
    if done.empty:
        empty["open_loads"] = open_loads
        return empty

    delivered_mt = float(done["offloaded_mt"].sum())
    outstanding = max(TARGET_TOTAL_MT - delivered_mt, 0.0)
    pta_balance = max(receipts_physical_mt - delivered_mt, 0.0)

    daily = (done.groupby("actual_delivery_date")
             .agg(loads=("offloaded_mt", "size"), mt=("offloaded_mt", "sum"))
             .reset_index().sort_values("actual_delivery_date"))
    daily["cum_mt"] = daily["mt"].cumsum()
    daily["cum_loads"] = daily["loads"].cumsum()

    first_d, last_d = daily["actual_delivery_date"].min(), daily["actual_delivery_date"].max()
    days_elapsed = (last_d - first_d).days + 1
    avg_daily_offtake = delivered_mt / days_elapsed if days_elapsed else 0.0

    projected = "Pending"
    remaining_days = None
    if avg_daily_offtake > 0:
        if outstanding <= 0:
            projected = "Completed"
        else:
            remaining_days = math.ceil(outstanding / avg_daily_offtake)
            base = max(date.today(), last_d)
            projected = (base + timedelta(days=remaining_days)).isoformat()

    # variance / tolerance on physically offloaded vs loaded, completed loads only
    tol_rows = done[done["actual_weight_offloaded_kg"].notna()
                    & done["loaded_weight_kg"].notna()] if \
        {"actual_weight_offloaded_kg", "loaded_weight_kg"}.issubset(done.columns) else pd.DataFrame()
    variance_kg = float((tol_rows["actual_weight_offloaded_kg"]
                         - tol_rows["loaded_weight_kg"]).sum()) if not tol_rows.empty else 0.0
    expected_kg = float(tol_rows["loaded_weight_kg"].sum()) if not tol_rows.empty else 0.0

    remained = None
    if "remained_in_tank_kg" in done:
        rt = pd.to_numeric(done["remained_in_tank_kg"], errors="coerce")
        rt = rt[rt.notna() & (rt >= 0)]
        remained = float(rt.mean()) if len(rt) else None

    by_iso = (done.groupby("iso")
              .agg(loads=("offloaded_mt", "size"), mt=("offloaded_mt", "sum"))
              .reset_index().sort_values("mt", ascending=False)) if "iso" in done else pd.DataFrame()

    # required rate to still hit the drawdown plan end date, if a plan exists
    required_rate = None
    if drawdown is not None and not drawdown.empty:
        plan_end = drawdown["date"].max()
        days_left = (plan_end - date.today()).days + 1
        if days_left > 0 and outstanding > 0:
            required_rate = outstanding / days_left

    return {
        "delivered_mt": r2(delivered_mt),
        "outstanding_mt": r2(outstanding),
        "completion_pct": r2(_pct(delivered_mt, TARGET_TOTAL_MT), 4),
        "target_total_mt": TARGET_TOTAL_MT,
        "loads_total": int(len(done)),
        "open_loads": open_loads,
        "avg_loads_per_day": r2(daily["loads"].mean()),
        "max_loads_day": int(daily["loads"].max()),
        "min_loads_day": int(daily["loads"].min()),
        "avg_daily_offtake_mt": r2(avg_daily_offtake),
        "avg_mt_per_load": r2(delivered_mt / len(done)),
        "avg_mt_offloaded_per_tank": r2(delivered_mt / len(done)),
        "pta_balance_mt": r2(pta_balance),
        "projected_completion": projected,
        "projected_remaining_days": remaining_days,
        "required_rate_mt_per_day": r2(required_rate),
        "avg_remained_in_tank_kg": r2(remained),
        "tolerance_pct": r2(_pct(variance_kg, expected_kg), 5),
        "variance_mt": r2(variance_kg / 1000.0, 3),
        "first_delivery": iso_d(first_d),
        "last_delivery": iso_d(last_d),
        "by_date": [{
            "date": iso_d(r.actual_delivery_date), "loads": int(r.loads),
            "mt": r2(r.mt), "cum_mt": r2(r.cum_mt), "cum_loads": int(r.cum_loads),
        } for r in daily.itertuples()],
        "by_iso": [{
            "iso": r.iso, "loads": int(r.loads), "mt": r2(r.mt),
            "actual_offloaded_mt": r2(r.mt),
        } for r in by_iso.itertuples() if r.iso],
    }


# --------------------------------------------------------------------------- #
# dwell + turnaround
# --------------------------------------------------------------------------- #
def dwell_block(visits: pd.DataFrame, turns: pd.DataFrame) -> dict:
    out = {"by_site": [], "by_iso": [], "trend": [], "turnaround": {}, "total_visits": 0}
    if visits is None or visits.empty:
        return out

    for site in (SITE_CONNECT, SITE_SAFRIPOL):
        s = visits[visits["site"] == site]
        if s.empty:
            continue
        breaches = int(pd.to_numeric(s["exceeds_site_target"], errors="coerce").fillna(0).sum())
        out["by_site"].append({
            "site": site,
            "visits": int(len(s)),
            "avg_hours": r2(s["dwell_hours"].mean()),
            "min_hours": r2(s["dwell_hours"].min()),
            "max_hours": r2(s["dwell_hours"].max()),
            "p90_hours": r2(s["dwell_hours"].quantile(0.90)),
            "p95_hours": r2(s["dwell_hours"].quantile(0.95)),
            "total_hours": r2(s["dwell_hours"].sum()),
            "target_hours": SITE_TARGET_HRS,
            "vs_target_hours": r2(s["dwell_hours"].mean() - SITE_TARGET_HRS),
            "breaches": breaches,
            "breach_pct": r2(_pct(breaches, len(s)), 4),
            "overstays": int(s["is_overstay_flag"].sum()),
            "overstay_pct": r2(_pct(s["is_overstay_flag"].sum(), len(s)), 4),
            "status": "Over Target" if s["dwell_hours"].mean() > SITE_TARGET_HRS
            else "Within Target",
        })

    piv = (visits.pivot_table(index="iso", columns="site", values="dwell_hours",
                              aggfunc="mean").reset_index())
    for r in piv.itertuples(index=False):
        row = r._asdict()
        out["by_iso"].append({
            "iso": row.get("iso"),
            "connect_avg_hours": r2(row.get(SITE_CONNECT)),
            "safripol_avg_hours": r2(row.get(SITE_SAFRIPOL)),
            "visits": int(len(visits[visits["iso"] == row.get("iso")])),
        })

    trend = (visits.groupby(["date", "site"])["dwell_hours"].mean().reset_index())
    for d in sorted(trend["date"].unique()):
        sub = trend[trend["date"] == d]
        out["trend"].append({
            "date": iso_d(d),
            "connect": r2(sub.loc[sub["site"] == SITE_CONNECT, "dwell_hours"].mean()),
            "safripol": r2(sub.loc[sub["site"] == SITE_SAFRIPOL, "dwell_hours"].mean()),
        })

    out["total_visits"] = int(len(visits))

    if turns is not None and not turns.empty:
        full = turns[turns.get("turnaround_hours").notna()] if "turnaround_hours" in turns \
            else pd.DataFrame()
        out["turnaround"] = {
            "cycles": int(len(turns)),
            "completed_cycles": int(len(full)),
            "avg_connect_dwell_hours": r2(_mean(turns["connect_dwell_hours"])),
            "avg_transit_out_hours": r2(_mean(turns["transit_out_hours"])),
            "avg_safripol_dwell_hours": r2(_mean(turns["safripol_dwell_hours"])),
            "avg_transit_return_hours": r2(_mean(turns.get("transit_return_hours"))),
            "avg_turnaround_hours": r2(_mean(full.get("turnaround_hours"))
                                       if not full.empty else None),
            "targets": {
                "connect_dwell": SITE_TARGET_HRS,
                "transit": TRANSIT_TARGET_HRS,
                "safripol_dwell": SITE_TARGET_HRS,
                "turnaround": SITE_TARGET_HRS * 2 + TRANSIT_TARGET_HRS * 2,
            },
            "trend": [{
                "date": iso_d(d),
                "turnaround": r2(_mean(g.get("turnaround_hours"))),
                "transit_out": r2(_mean(g["transit_out_hours"])),
                "safripol_dwell": r2(_mean(g["safripol_dwell_hours"])),
            } for d, g in turns.groupby("date")],
        }
    return out


# --------------------------------------------------------------------------- #
# decanting + shifts
# --------------------------------------------------------------------------- #
def decant_block(decant: pd.DataFrame, delays: pd.DataFrame,
                 plan: pd.DataFrame, staff: pd.DataFrame) -> dict:
    out = {
        "has_data": False, "total_decants": 0, "avg_decant_hours": None,
        "avg_interval_hours": None, "avg_productive_hours": None,
        "clean_pct": None, "time_lost_hours": 0.0, "delay_hours_logged": 0.0,
        "total_net_mt": 0.0, "decants_per_day": None,
        "by_day": [], "by_shift": [], "by_team": [], "by_iso": [],
        "delay_categories": [], "delay_owners": [], "hour_profile": [],
        "shifts": [], "shift_summary": {}, "staff": {},
    }
    if staff is not None and not staff.empty:
        by_team = staff.groupby("team").size().to_dict()
        out["staff"] = {
            "total": int(len(staff)),
            "by_team": [{"team": k, "headcount": int(v)} for k, v in sorted(by_team.items())],
            "by_company": [{"company": k, "headcount": int(v)} for k, v in
                           sorted(staff.groupby("company").size().to_dict().items())],
            "by_position": [{"position": k, "headcount": int(v)} for k, v in
                            sorted(staff.groupby("position").size().to_dict().items())],
            "trained_pct": r2(_pct(
                int(staff["trained"].str.lower().eq("yes").sum()) if "trained" in staff else 0,
                len(staff)), 4),
            "experienced_pct": r2(_pct(
                int(staff["worked_previous"].str.lower().eq("yes").sum())
                if "worked_previous" in staff else 0, len(staff)), 4),
        }

    if decant is None or decant.empty:
        return out

    out["has_data"] = True
    d = decant
    completed = d[d["decant_hours"].notna()]

    out["total_decants"] = int(len(completed))
    out["avg_decant_hours"] = r2(_mean(completed["decant_hours"]))
    out["avg_interval_hours"] = r2(_mean(d["decant_interval_hours"]))
    out["avg_productive_hours"] = r2(_mean(completed["productive_hours"]))
    out["allowed_decant_hours"] = ALLOWED_DECANT_HRS
    out["allowed_interval_hours"] = ALLOWED_INTERVAL_HRS
    clean = int((completed["performance"] == "Clean").sum())
    out["clean_pct"] = r2(_pct(clean, len(completed)), 4)
    over = completed[completed["time_variance_hours"] > 0]
    out["time_lost_hours"] = r2(float(over["time_variance_hours"].sum()))
    out["time_saved_hours"] = r2(abs(float(
        completed.loc[completed["time_variance_hours"] < 0, "time_variance_hours"].sum())))
    out["delay_hours_logged"] = r2(float(d["delay_minutes"].sum()) / 60.0)
    out["total_net_mt"] = r2(float(pd.to_numeric(d["net_product_mt"], errors="coerce")
                                   .fillna(0).sum()))
    days = d["shift_date"].nunique()
    out["decants_per_day"] = r2(len(completed) / days if days else None)
    out["avg_excess_kg"] = r2(_mean(d["excess_kg"]))
    out["avg_carry_over_kg"] = r2(_mean(d["carry_over_kg"]))

    daily = (d.groupby("shift_date")
             .agg(decants=("decant_hours", "count"),
                  net_mt=("net_product_mt", "sum"),
                  avg_decant=("decant_hours", "mean"),
                  avg_interval=("decant_interval_hours", "mean"),
                  delay_mins=("delay_minutes", "sum"))
             .reset_index().sort_values("shift_date"))
    daily["cum_mt"] = daily["net_mt"].cumsum()
    out["by_day"] = [{
        "date": iso_d(r.shift_date), "decants": int(r.decants), "net_mt": r2(r.net_mt),
        "cum_mt": r2(r.cum_mt), "avg_decant_hours": r2(r.avg_decant),
        "avg_interval_hours": r2(r.avg_interval), "delay_minutes": r2(r.delay_mins),
    } for r in daily.itertuples()]

    for key, label in (("shift", "by_shift"), ("team", "by_team"), ("iso", "by_iso")):
        if key not in d:
            continue
        grp = (d[d[key].notna() & (d[key].astype(str) != "")]
               .groupby(key)
               .agg(decants=("decant_hours", "count"),
                    avg_decant=("decant_hours", "mean"),
                    avg_interval=("decant_interval_hours", "mean"),
                    net_mt=("net_product_mt", "sum"),
                    delay_mins=("delay_minutes", "sum"))
               .reset_index())
        out[label] = [{
            key: getattr(r, key), "decants": int(r.decants),
            "avg_decant_hours": r2(r.avg_decant),
            "avg_interval_hours": r2(r.avg_interval),
            "net_mt": r2(r.net_mt), "delay_minutes": r2(r.delay_mins),
        } for r in grp.itertuples()]

    # time lost by reason - combines in-decant delay minutes and standalone delay log
    reasons: Counter = Counter()
    owners: Counter = Counter()
    for _, row in d.iterrows():
        mins = float(row.get("delay_minutes") or 0)
        if mins > 0:
            reasons[row.get("delay_category") or "Uncategorised"] += mins
            owners[row.get("delay_owner") or "Unassigned"] += mins
    if delays is not None and not delays.empty:
        for _, row in delays.iterrows():
            mins = float(row.get("delay_minutes") or 0)
            if mins > 0:
                reasons[row.get("delay_category") or "Uncategorised"] += mins
                owners[row.get("delay_owner") or "Unassigned"] += mins
    total_delay = sum(reasons.values())
    out["delay_categories"] = [{
        "category": k, "minutes": r2(v), "hours": r2(v / 60.0),
        "share": r2(_pct(v, total_delay), 4),
    } for k, v in reasons.most_common() if k and k != "None"]
    out["delay_owners"] = [{
        "owner": k, "minutes": r2(v), "hours": r2(v / 60.0),
        "share": r2(_pct(v, total_delay), 4),
    } for k, v in owners.most_common() if k]
    out["total_delay_hours"] = r2(total_delay / 60.0)

    prof: dict[int, list[float]] = defaultdict(list)
    for _, row in completed.iterrows():
        prof[row["start"].hour].append(row["decant_hours"])
    out["hour_profile"] = [{
        "hour": h, "decants": len(prof.get(h, [])),
        "avg_decant_hours": r2(_mean(prof[h])) if prof.get(h) else None,
    } for h in range(24)]

    # ---- shift level: planned vs actual -----------------------------------
    actual = (d.groupby("shift_key")
              .agg(shift_date=("shift_date", "first"), shift=("shift", "first"),
                   team=("team", "first"), heads=("heads", "max"),
                   actual_isotainers=("decant_hours", "count"),
                   actual_mt=("net_product_mt", "sum"),
                   avg_decant=("decant_hours", "mean"),
                   avg_interval=("decant_interval_hours", "mean"),
                   delay_mins=("delay_minutes", "sum"),
                   productive=("productive_hours", "sum"))
              .reset_index())

    planned_map: dict[str, dict] = {}
    if plan is not None and not plan.empty:
        for _, p in plan.iterrows():
            planned_map[p["shift_key"]] = {
                "planned_isotainers": p.get("planned_isotainers"),
                "planned_heads": p.get("planned_heads"),
                "notes": p.get("notes"),
            }

    shifts = []
    for r in actual.itertuples():
        p = planned_map.get(r.shift_key, {})
        planned = p.get("planned_isotainers")
        planned = float(planned) if planned not in (None, "") and not pd.isna(planned) else None
        attain = _pct(r.actual_isotainers, planned) if planned else None
        heads = r.heads if r.heads and not pd.isna(r.heads) else p.get("planned_heads")
        heads = float(heads) if heads not in (None, "") and not pd.isna(heads) else None
        shifts.append({
            "date": iso_d(r.shift_date), "shift": r.shift, "team": r.team,
            "heads": r2(heads, 0), "planned_isotainers": r2(planned, 0),
            "actual_isotainers": int(r.actual_isotainers),
            "variance_isotainers": r2((r.actual_isotainers - planned) if planned else None, 0),
            "attainment_pct": r2(attain, 4),
            "actual_mt": r2(r.actual_mt),
            "avg_decant_hours": r2(r.avg_decant),
            "avg_interval_hours": r2(r.avg_interval),
            "delay_minutes": r2(r.delay_mins),
            "productive_hours": r2(r.productive),
            "isotainers_per_head": r2(r.actual_isotainers / heads, 3) if heads else None,
            "notes": p.get("notes") if isinstance(p.get("notes"), str) else "",
        })
    shifts.sort(key=lambda s: (s["date"], s["shift"]), reverse=True)
    out["shifts"] = shifts

    rated = [s for s in shifts if s["attainment_pct"] is not None]
    out["shift_summary"] = {
        "shifts_recorded": len(shifts),
        "shifts_planned": len(rated),
        "avg_attainment_pct": r2(_mean([s["attainment_pct"] for s in rated]), 4)
        if rated else None,
        "shifts_on_target": sum(1 for s in rated if s["attainment_pct"] >= 1),
        "avg_isotainers_per_shift": r2(_mean([s["actual_isotainers"] for s in shifts])),
        "avg_heads_per_shift": r2(_mean([s["heads"] for s in shifts if s["heads"]])),
        "avg_isotainers_per_head": r2(
            _mean([s["isotainers_per_head"] for s in shifts if s["isotainers_per_head"]]), 3),
        "best_shift": max(shifts, key=lambda s: s["actual_isotainers"]) if shifts else None,
    }
    return out


# --------------------------------------------------------------------------- #
# plan vs actual
# --------------------------------------------------------------------------- #
def plan_block(drawdown: pd.DataFrame, dispatch_by_date: list[dict]) -> dict:
    if drawdown is None or drawdown.empty:
        return {"has_plan": False, "rows": [], "variance_mt": None, "days": 0}
    actual_map = {r["date"]: r["cum_mt"] for r in dispatch_by_date}
    running = None
    rows = []
    for r in drawdown.itertuples():
        d = iso_d(r.date)
        if d in actual_map:
            running = actual_map[d]
        rows.append({
            "date": d,
            "planned_mt": r2(r.planned_mt),
            "cum_planned_mt": r2(r.cum_planned_mt),
            "cum_actual_mt": r2(running),
            "variance_mt": r2((running - r.cum_planned_mt) if running is not None else None),
        })
    latest = next((x for x in reversed(rows) if x["cum_actual_mt"] is not None), None)
    return {
        "has_plan": True,
        "rows": rows,
        "plan_end": rows[-1]["date"] if rows else None,
        "plan_total_mt": rows[-1]["cum_planned_mt"] if rows else None,
        "variance_mt": latest["variance_mt"] if latest else None,
        "days": len(rows),
    }


# --------------------------------------------------------------------------- #
# exceptions - what management must action
# --------------------------------------------------------------------------- #
def exceptions_block(decant: pd.DataFrame, delays: pd.DataFrame,
                     dwell: dict, dispatch: dict, decant_stats: dict) -> list[dict]:
    items: list[dict] = []

    def add(severity, category, detail, when=None, owner="", value=None):
        items.append({"severity": severity, "category": category, "detail": detail,
                      "when": when, "owner": owner, "value": value})

    if delays is not None and not delays.empty:
        esc = delays[delays.get("escalate", False) == True]  # noqa: E712
        for _, r in esc.sort_values("shift_date", ascending=False).head(25).iterrows():
            add("high", r.get("delay_category") or "Escalation",
                (r.get("description") or "Escalated delay").strip(),
                iso_dt(r.get("start")) or iso_d(r.get("shift_date")),
                r.get("delay_owner") or "",
                f"{r.get('delay_minutes', 0):.0f} min")

    if decant is not None and not decant.empty and "escalate" in decant:
        esc = decant[decant["escalate"] == True]  # noqa: E712
        for _, r in esc.sort_values("shift_date", ascending=False).head(25).iterrows():
            add("high", r.get("delay_category") or "Decant escalation",
                (r.get("delay_detail") or r.get("comments") or "Escalated decant").strip(),
                iso_dt(r.get("start")), r.get("delay_owner") or "",
                f"{float(r.get('delay_minutes') or 0):.0f} min")

    for site in dwell.get("by_site", []):
        if site["status"] == "Over Target":
            add("medium", "Dwell over target",
                f"{site['site']} average dwell {site['avg_hours']} h against a "
                f"{site['target_hours']} h target across {site['visits']} visits.",
                None, "Connect" if site["site"] == SITE_CONNECT else "Safripol",
                f"+{site['vs_target_hours']} h")
        if site.get("breach_pct") and site["breach_pct"] > 0.25:
            add("medium", "Dwell breach rate",
                f"{site['site']} breached its dwell target on "
                f"{site['breaches']} of {site['visits']} visits.",
                None, site["site"], f"{site['breach_pct'] * 100:.0f}%")

    ta = dwell.get("turnaround") or {}
    if ta.get("avg_transit_out_hours") and ta["avg_transit_out_hours"] > TRANSIT_TARGET_HRS * 1.25:
        add("medium", "Transit over target",
            f"Average Connect to Safripol transit {ta['avg_transit_out_hours']} h "
            f"against a {TRANSIT_TARGET_HRS} h target.", None, "Transporter",
            f"+{r2(ta['avg_transit_out_hours'] - TRANSIT_TARGET_HRS)} h")

    if decant_stats.get("avg_decant_hours") and \
            decant_stats["avg_decant_hours"] > ALLOWED_DECANT_HRS:
        add("medium", "Decant over standard",
            f"Average decant {decant_stats['avg_decant_hours']} h against the "
            f"{ALLOWED_DECANT_HRS} h standard.", None, "Connect",
            f"+{r2(decant_stats['avg_decant_hours'] - ALLOWED_DECANT_HRS)} h")

    if decant_stats.get("avg_interval_hours") and \
            decant_stats["avg_interval_hours"] > ALLOWED_INTERVAL_HRS:
        add("low", "Interval over target",
            f"Average gap between decants {decant_stats['avg_interval_hours']} h "
            f"against a {ALLOWED_INTERVAL_HRS} h target.", None, "Connect",
            f"+{r2(decant_stats['avg_interval_hours'] - ALLOWED_INTERVAL_HRS)} h")

    for s in (decant_stats.get("shifts") or [])[:40]:
        if s["attainment_pct"] is not None and s["attainment_pct"] < 0.8:
            add("high", "Shift below plan",
                f"{s['date']} {s['shift']} team {s['team']}: "
                f"{s['actual_isotainers']} of {s['planned_isotainers']} isotainers.",
                s["date"], "Connect", f"{s['attainment_pct'] * 100:.0f}%")

    req = dispatch.get("required_rate_mt_per_day")
    avg = dispatch.get("avg_daily_offtake_mt")
    if req and avg and avg < req:
        add("high", "Behind drawdown pace",
            f"Current offtake {avg} MT/day is below the {r2(req)} MT/day needed to "
            f"finish on plan.", None, "Connect", f"-{r2(req - avg)} MT/day")

    top = (decant_stats.get("delay_categories") or [])[:3]
    for t in top:
        if t["hours"] and t["hours"] >= 1:
            add("low", "Top time loss", f"{t['category']} accounts for {t['hours']} h "
                f"({t['share'] * 100:.0f}%) of recorded time lost.", None, "", f"{t['hours']} h")

    order = {"high": 0, "medium": 1, "low": 2}
    items.sort(key=lambda i: (order.get(i["severity"], 9), i["category"]))
    return items
