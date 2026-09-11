# Safripol Operations KPI Report — MV TAC IMOLA

Live operations dashboard for the **TAC IMOLA** PTA shipment, covering the full
chain from vessel discharge through Connect warehousing, decanting, transit and
delivery into Safripol.

Built for Safripol to self-serve, and for the Connect ops team to see where time
is being lost while there is still time to react.

---

## What it shows

| Tab | Answers |
|---|---|
| **Overview** | Where are we against the 20 400,00 MT drawdown, and will we finish on time? |
| **Drawdown** | Cumulative plan vs actual, offtake rate, receipts, split by isotainer. |
| **Decanting** | Average decant time, interval between decants, shift output vs planned isotainers, heads on site, hour-of-day profile. |
| **Turnaround** | Connect and Safripol dwell vs the 2,75 h target, transit legs, full 7,00 h cycle — all derived from the GPS feed. |
| **Time Lost** | Every delay by category and by owner, ranked by hours lost. |
| **Vessel** | Hatch-by-hatch discharge outturn and weather downtime. |
| **Team** | Shift resourcing linked to output — heads per team and productivity per head. |
| **Data** | Every source, when it was last read, and whether it is currently reachable. |

## How it stays live

```
SharePoint / OneDrive workbooks
        │  Microsoft Graph (read-only)
        ▼
   python -m ingest        ← runs every 30 min via a Windows Scheduled Task
        │  parses, recomputes every KPI, ranks exceptions
        ▼
   docs/data/snapshot.json  (committed and pushed)
        │  GitHub Pages
        ▼
   docs/index.html          ← polls version.json every 60 s, repaints on change
```

The page is plain HTML/CSS/JS with one vendored copy of Chart.js. No build step,
no framework, no CDN call, no tracking. It loads on a phone on site.

**Nothing writes back to SharePoint.** The Graph permission is read-only; the
source workbooks cannot be modified by this system.

> **Where the refresh runs, and why it isn't CI.** The tenant granted
> *Delegated* `Files.Read.All` rather than *Application*, so Graph tokens are
> tied to a signed-in user and cannot be issued on a GitHub runner. The refresh
> therefore runs on an operator PC under Task Scheduler and pushes the snapshot.
> `.github/workflows/refresh-data.yml` is ready to take over unchanged if an
> Application grant is approved later. See `ENTRA_SETUP.md`.

### First-time setup

```powershell
pip install -r requirements.txt
python -m ingest --login     # sign in once (device code)
python -m ingest --check     # confirm all four sources read
powershell -ExecutionPolicy Bypass -File tools\install_task.ps1 -Minutes 30
```

### Monitoring the refresh

Run the desktop shortcut command:

```text
C:\Users\Christopher.Viljoen\OneDrive - Connect Logistics\Desktop\Safripol Report Spot Check.cmd
```

It shows the scheduled task result, next run time, published snapshot timestamp,
headline tonnages, source status and the latest `refresh.log` lines. To force a
refresh immediately:

```powershell
Start-ScheduledTask -TaskName 'Safripol Ops Report Refresh'
```

### Sources

| Source | Owner | Used for |
|---|---|---|
| `Safripol Stock Report - TAC IMOLA.xlsx` | Warehouse | Receipts, dispatches, stock balance |
| `Safripol v Connect Dwells.xlsx` | Automated GPS feed | Dwell, transit, turnaround — **read-only, never edited** |
| `SAF Ops Tracking - TAC IMOLA.xlsx` | Yard / ops team | Decant times, delays, shift plan, heads |
| `Saf Staff Roles.xlsx` | Ops management | Team and role allocation |

## Running it locally

```powershell
pip install -r requirements.txt

# with live data
$env:AZURE_TENANT_ID="..."; $env:AZURE_CLIENT_ID="..."; $env:AZURE_CLIENT_SECRET="..."
python -m ingest

# or from a folder of downloaded workbooks
python -m ingest --local "C:\path\to\Safripol Reports"

# then
cd docs; python -m http.server 8099
```

Open <http://localhost:8099>.

Check Graph connectivity without building anything:

```powershell
python -m ingest --check
```

## Self-hosted (Docker)

Only needed if the report is served internally rather than from Pages. The
container refreshes on its own timer and serves the site.

```bash
docker compose up -d --build
# http://localhost:8080
```

## Repository layout

```
ingest/          data pipeline
  config.py        business constants, source URLs, targets   ← edit for a new shipment
  graph.py         Microsoft Graph auth + download, with local/cache fallback
  sources.py       workbook parsers
  metrics.py       every KPI (ported from the Power BI model, plus the ops KPIs)
  snapshot.py      assembles snapshot.json
docs/            the published site (GitHub Pages serves this folder)
tools/
  build_tracker.py regenerates the ops tracking workbook
ENTRA_SETUP.md   app registration, permissions, secrets
```

## Setting up a new shipment

Everything shipment-specific is in `ingest/config.py`: the vessel label, the
`TARGET_TOTAL_MT`, the four source URLs, the hatch table and the weather events.
Change those, re-run `tools/build_tracker.py` for a fresh capture workbook, and
the rest follows.

## Data handling

This repository contains Safripol operational data. Keep it **private** and
serve the site through an access-controlled host — see the hosting note in
`ENTRA_SETUP.md`. Credentials live only in GitHub Actions secrets and are never
committed.
