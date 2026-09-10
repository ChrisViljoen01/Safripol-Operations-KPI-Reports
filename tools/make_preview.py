"""Dev harness: build a snapshot using the previous shipment's captured rows so the
whole UI can be validated before TAC IMOLA capture starts."""
import shutil, tempfile, json, random
from pathlib import Path
from datetime import timedelta
import pandas as pd
import ingest.sources as S
from ingest import metrics
from ingest.snapshot import write_snapshot
from ingest.config import SOURCES

base = Path(r"C:\Users\Christopher.Viljoen\Connect Logistics\Process Optimization and Development - Documents\Power BI\Safripol Reports")
tmp = Path(tempfile.gettempdir())
old = tmp/"old.xlsx"; shutil.copy2(base/"SAF Ops Tracking .xlsx", old)
dw  = tmp/"dw.xlsx";  shutil.copy2(base/"Safripol v Connect Dwells.xlsx", dw)
stf = tmp/"stf.xlsx"; shutil.copy2(base/"Saf Staff Roles .xlsx", stf)

orig=S._find_sheet
S._find_sheet=lambda p,*c: "Old Version" if "Decant Log" in c else orig(p,*c)
dec = S.parse_decant_log(old)
S._find_sheet=orig

# synthesise the delay + plan context the NEW tracker will capture, so every
# panel is exercised. Dev only - never used by the real refresh.
random.seed(7)
cats=["Bad / Damaged Bags","Breakdown - Forklift","Safripol Site Delay","WIFI / System Downtime",
      "Weighbridge Delay","Shift Handover","Weather - Rain","Isotainer Not Available"]
owners={"Bad / Damaged Bags":"Connect","Breakdown - Forklift":"Connect","Safripol Site Delay":"Safripol",
        "WIFI / System Downtime":"Third Party","Weighbridge Delay":"Connect","Shift Handover":"Connect",
        "Weather - Rain":"External / Weather","Isotainer Not Available":"Transporter"}
dm=[];cc=[];oo=[];es=[]
for _,r in dec.iterrows():
    if random.random()<0.28:
        c=random.choice(cats); m=random.choice([10,15,20,25,30,45,60,90])
        dm.append(m); cc.append(c); oo.append(owners[c]); es.append(m>=60 and random.random()<0.4)
    else:
        dm.append(0.0); cc.append("None"); oo.append(""); es.append(False)
dec["delay_minutes"]=dm; dec["delay_category"]=cc; dec["delay_owner"]=oo; dec["escalate"]=es
dec["delay_detail"]=["Logged during decant" if m else "" for m in dm]
dec["productive_hours"]=(dec["decant_hours"]-dec["delay_minutes"]/60).clip(lower=0)
dec["heads"]=[random.choice([10,11,12,13]) for _ in range(len(dec))]

plan=(dec.groupby("shift_key").agg(shift_date=("shift_date","first"),shift=("shift","first"),
      team=("team","first")).reset_index())
plan["planned_isotainers"]=[random.choice([5,6,7]) for _ in range(len(plan))]
plan["planned_heads"]=12; plan["notes"]=""

visits=S.parse_dwell_visits(dw); turns=S.build_turnaround(visits)
staff=S.parse_staff(stf)

# a drawdown plan and a dispatch history consistent with the decant record
days=sorted(dec["shift_date"].unique())
draw=pd.DataFrame({"date":days})
draw["planned_isotainers"]=11; draw["planned_mt"]=11*21600/1000
draw["cum_planned_mt"]=draw["planned_mt"].cumsum()

disp=[]
for _,r in dec[dec["decant_hours"].notna()].iterrows():
    disp.append({"actual_delivery_date":r["shift_date"],"iso":r["iso"],
                 "loaded_weight_kg":r["loaded_kg"],
                 "actual_weight_offloaded_kg":(r["loaded_kg"] or 0)-random.randint(0,120),
                 "remained_in_tank_kg":random.randint(0,180),"offloaded_mt":(r["loaded_kg"] or 0)/1000})
disp=pd.DataFrame(disp); disp["is_delivered"]=True
rec=pd.DataFrame({"arrival_date":days,"total_bags":250,"damaged_bags":0,
                  "total_received_kg":250*1200,"receipt_type":"Direct from Vessel"})
rec["received_mt_physical"]=rec["total_received_kg"]/1000
rec["received_mt_admin"]=rec["total_bags"]*1.2

recb=metrics.receipts_block(rec)
dispb=metrics.dispatch_block(disp,recb["total_physical_mt"],draw)
dwb=metrics.dwell_block(visits,turns)
decb=metrics.decant_block(dec,pd.DataFrame(),plan,staff)
plnb=metrics.plan_block(draw,dispb["by_date"])
excb=metrics.exceptions_block(dec,pd.DataFrame(),dwb,dispb,decb)

from ingest.snapshot import build_snapshot
import ingest.snapshot as SN
snap=SN.build_snapshot.__wrapped__ if hasattr(SN.build_snapshot,"__wrapped__") else None
from datetime import datetime,timezone
from ingest.config import SAST,SHIPMENT,VESSEL_LABEL,TARGET_TOTAL_MT
now=datetime.now(timezone.utc); ns=now.astimezone(SAST)
out={"schema_version":1,"meta":{"shipment":SHIPMENT+" (PREVIEW - previous shipment data)","vessel":VESSEL_LABEL,
 "client":"Safripol","operator":"Connect Logistics","product":"PTA",
 "generated_at_utc":now.isoformat(timespec="seconds"),"generated_at_sast":ns.isoformat(timespec="seconds"),
 "generated_display":ns.strftime("%d %b %Y %H:%M")+" SAST",
 "sources":[{"key":s.key,"label":s.label,"ok":True,"origin":"preview","last_modified":None,
             "size_kb":0,"required":s.required,"error":""} for s in SOURCES],"degraded":False},
 "headline":{"target_total_mt":TARGET_TOTAL_MT,"delivered_mt":dispb["delivered_mt"],
   "outstanding_mt":dispb["outstanding_mt"],"completion_pct":dispb["completion_pct"],
   "pta_balance_mt":dispb["pta_balance_mt"],"received_admin_mt":recb["total_admin_mt"],
   "received_physical_mt":recb["total_physical_mt"],"loads_total":dispb["loads_total"],
   "avg_daily_offtake_mt":dispb["avg_daily_offtake_mt"],"avg_loads_per_day":dispb["avg_loads_per_day"],
   "required_rate_mt_per_day":dispb["required_rate_mt_per_day"],"projected_completion":dispb["projected_completion"],
   "projected_remaining_days":dispb["projected_remaining_days"],"avg_decant_hours":decb["avg_decant_hours"],
   "avg_interval_hours":decb["avg_interval_hours"],"time_lost_hours":decb.get("total_delay_hours"),
   "decants_per_day":decb["decants_per_day"],"avg_attainment_pct":decb["shift_summary"].get("avg_attainment_pct"),
   "avg_turnaround_hours":(dwb.get("turnaround") or {}).get("avg_turnaround_hours"),
   "open_exceptions":len([e for e in excb if e["severity"]=="high"])},
 "vessel_discharge":metrics.vessel_block(),"receipts":recb,"dispatch":dispb,"dwell":dwb,
 "decant":decb,"plan":plnb,"exceptions":excb}
Path("preview").mkdir(exist_ok=True)
write_snapshot(out, Path("preview"))
print("preview snapshot written")
print("decants",decb["total_decants"],"shifts",len(decb["shifts"]),"exc",len(excb),
      "delaycats",len(decb["delay_categories"]),"attain",decb["shift_summary"]["avg_attainment_pct"])
