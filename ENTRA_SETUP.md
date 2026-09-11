# Entra (Azure AD) setup

> **Current state:** `Sites.Selected` (Application) has been **requested** and is
> awaiting admin consent. Until it lands, the report runs in
> [delegated mode](#delegated-mode-fallback) from an operator PC. The moment
> consent + the per-site grants are in place, run
> `python tools\check_app_auth.py` — when it prints **READY**, the refresh moves
> to GitHub Actions and no PC is involved.

| | App-only *(target)* | Delegated *(fallback, in use)* |
|---|---|---|
| Permission type | Application | Delegated |
| Who it reads as | the app itself | Christopher Viljoen |
| Sign-in | none, ever | once per machine |
| Can run on GitHub Actions | yes | no — needs the cached token |
| Where the refresh runs | CI runner | an operator PC, via Task Scheduler |

---

## App-only mode with `Sites.Selected` (target state)

`Sites.Selected` is the least-privilege option: consenting to it grants the app
**no data at all** until an admin separately names each site it may read. That is
why it is a two-step process, and why step 2 is the one people forget.

### Step 1 — Admin consent

Entra ID → **App registrations** → **Connect Logistics AI Hub**
(`207d9293-d311-4b5d-ba2b-5bbd3d3ade48`) → **API permissions**

Microsoft Graph → **Application permissions** → **`Sites.Selected`** →
**Grant admin consent for Connect Logistics**.

The status column must read **Granted**. Adding the permission without consenting
leaves the token's `roles` claim empty, which is exactly the 401 seen before.

### Step 2 — Grant the app access to the site

Consent alone does nothing. An admin must grant `read` on the site collection
that holds the sources. All four workbooks now live in one place, so this is a
single grant:

| Source | Lives in | Site collection |
|---|---|---|
| All four workbooks | team site | `connectlogisticscoza.sharepoint.com:/sites/ProcessOptimizationandDevelopment` |

Run these in **Graph Explorer** (`https://developer.microsoft.com/graph/graph-explorer`)
signed in as an admin. First resolve the site id:

```http
GET https://graph.microsoft.com/v1.0/sites/connectlogisticscoza.sharepoint.com:/sites/ProcessOptimizationandDevelopment
```

Then grant read on it, substituting the `id` returned above:

```http
POST https://graph.microsoft.com/v1.0/sites/{site-id}/permissions
Content-Type: application/json

{
  "roles": ["read"],
  "grantedToIdentities": [
    {
      "application": {
        "id": "207d9293-d311-4b5d-ba2b-5bbd3d3ade48",
        "displayName": "Connect Logistics AI Hub"
      }
    }
  ]
}
```

`"roles": ["read"]` is deliberate — the job never writes to SharePoint.

> **If the admin will not grant the OneDrive one:** move
> `Safripol Stock Report - TAC IMOLA.xlsx` onto the team site instead and update
> its URL in `ingest/config.py`. That is the better end state anyway — see
> *Hardening*.

### Step 3 — Verify

```powershell
python tools\check_app_auth.py
```

It decodes the token, prints the granted roles, and reads all four sources:

```
  [OK]  Application permissions granted: Sites.Selected
  [OK]  Safripol Stock Report - TAC IMOLA
  [OK]  Safripol v Connect Dwells
  [OK]  SAF Ops Tracking - TAC IMOLA
  [OK]  Saf Staff Roles
READY. The refresh can run in GitHub Actions without your PC.
```

### Step 4 — Move the refresh to the cloud

```bash
gh secret set AZURE_TENANT_ID
gh secret set AZURE_CLIENT_ID
gh secret set AZURE_CLIENT_SECRET
```

`.github/workflows/refresh-data.yml` is already scheduled every 15 minutes and
already sets `SAFRIPOL_AUTH_MODE: app`, so it starts working on its own once the
grants are live. Confirm two or three green runs, then retire the PC task:

```powershell
powershell -ExecutionPolicy Bypass -File tools\install_task.ps1 -Remove
```

> It is safe to leave both running during the changeover. The workflow refuses to
> publish a degraded snapshot, so a failed cloud run cannot blank the report.

### The client secret

**Certificates & secrets** → **New client secret** → expires **24 months**. Copy
the **Value**, not the Secret ID — it is shown once.

> ⚠️ Set a calendar reminder a month before expiry. When it lapses the report
> stops updating and shows a stale-data banner.

---

## Delegated mode (fallback)

Already configured, and what runs today. The app has Delegated `Files.Read.All`
and a public-client redirect URI, which is what the device-code sign-in needs.

```powershell
cd C:\Users\Christopher.Viljoen\source\repos\Safripol-Operations-KPI-Reports
python -m ingest --login          # sign in once, in a browser
python -m ingest --check          # confirm all four sources read
powershell -ExecutionPolicy Bypass -File tools\install_task.ps1 -Minutes 15
```

`--login` caches a refresh token at
`%USERPROFILE%\.safripol_report\token_cache.json`. It renews on every run, so it
keeps working as long as the task runs at least every ~90 days and the account
password does not change.

### What this means day to day

- The report reads **as you**, so it can only ever see what you can already see.
- The refresh runs while you are logged in to this PC. If it is off, the site
  keeps serving the last published snapshot and shows its age.
- Re-run `python -m ingest --login` after a password change or if the log shows
  `No usable cached sign-in`.

### Checking it is healthy

```powershell
Get-ScheduledTask -TaskName 'Safripol Ops Report Refresh' | Get-ScheduledTaskInfo
Get-Content refresh.log -Tail 20
```

`LastTaskResult : 0` means the last run published cleanly.

---

## Troubleshooting

| Symptom | Cause | Fix |
|---|---|---|
| `CERTIFICATE_VERIFY_FAILED` | corporate TLS proxy; Python does not trust the internal CA | handled automatically by `truststore`; ensure `pip install -r requirements.txt` has run |
| `401 generalException` with `roles: null` | permission added but **not consented** | complete Step 1 |
| `403 accessDenied` with `roles: Sites.Selected` | consented, but the site grant is missing | complete Step 2 for that site |
| `AADSTS7000215` invalid client secret | Secret ID copied instead of Value | recreate the secret, copy the **Value** |
| `404 itemNotFound` | file renamed or moved | update the URL in `ingest/config.py` |
| `No usable cached sign-in` | token cache missing or expired | `python -m ingest --login` |
| Report timestamp is old | refresh machine off, or task failed | check `refresh.log` and the task's `LastTaskResult` |

## Hardening (recommended follow-up)

1. ~~Move the stock report off personal OneDrive~~ — **done 11 Sep 2026.**
   `Safripol Stock Report - TAC IMOLA.xlsx` now lives on the **Process
   Optimization and Development** team site under `Safripol/2026/`, so all four
   sources sit in one site collection and only one `Sites.Selected` grant is
   needed. It also removes the single point of failure of a personal OneDrive.
2. Replace the client secret with **federated credentials (OIDC)** so GitHub
   Actions authenticates with nothing stored and nothing to expire.


