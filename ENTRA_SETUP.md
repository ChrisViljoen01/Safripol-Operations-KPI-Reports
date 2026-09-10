# Entra (Azure AD) setup

> **Current state:** this tenant granted **Delegated** `Files.Read.All` only, so
> the report runs in *delegated mode* — see [Delegated mode](#delegated-mode-current-setup)
> below, which is what is actually deployed. The app-only instructions are kept
> because they are the better end state if an Application grant is ever approved.

The two modes differ in one important way:

| | App-only | Delegated *(in use)* |
|---|---|---|
| Permission type | Application | Delegated |
| Who it reads as | the app itself | Christopher Viljoen |
| Sign-in | none, ever | once per machine |
| Can run on GitHub Actions | yes | no — needs the cached token |
| Where the refresh runs | CI runner | an operator PC, via Task Scheduler |

---

## Delegated mode (current setup)

Nothing further is needed in Entra. The app registration
**Connect Logistics AI Hub** (`207d9293-d311-4b5d-ba2b-5bbd3d3ade48`) already has
Delegated `Files.Read.All`, and its "public client" redirect URI allows the
device-code sign-in this uses.

### One-off setup on the machine that will refresh

```powershell
cd C:\Users\Christopher.Viljoen\source\repos\Safripol-Operations-KPI-Reports
python -m ingest --login          # sign in once, in a browser
python -m ingest --check          # confirm all four sources read
powershell -ExecutionPolicy Bypass -File tools\install_task.ps1 -Minutes 30
```

`--login` caches a refresh token at
`%USERPROFILE%\.safripol_report\token_cache.json`. It renews itself on every run,
so it keeps working indefinitely as long as the task runs at least every ~90 days
and the account's password does not change.

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

### Moving it off your PC

The dependency on one workstation is the weak point of delegated mode. Options,
best first:

1. **Get Application permission approved** (below) and move the refresh to
   GitHub Actions. No machine involved.
2. Run the refresh on an always-on server or VM, signed in once as a shared
   service account with read access to the four files.
3. Keep it here, but tell the ops team that the "last refreshed" stamp on the
   report is the thing to watch.

---

## App-only mode (preferred, needs admin approval)

Ask the tenant admin for:

> Microsoft Graph → **Application** permission → **`Files.Read.All`** → *Grant admin consent*
> on app `Connect Logistics AI Hub` (`207d9293-d311-4b5d-ba2b-5bbd3d3ade48`).
>
> It is read-only (not `ReadWrite`), it runs unattended so Delegated cannot work,
> and it is used to read four Excel files for a client operations dashboard.

To confirm a grant actually landed, decode the token — `roles` must contain
`Files.Read.All`. If `roles` is `null`, only Delegated was granted.

### 1. Create the app registration

Azure portal → **Microsoft Entra ID** → **App registrations** → **New registration**

| Field | Value |
|---|---|
| Name | `Safripol Ops Report – Data Refresh` |
| Supported account types | *Accounts in this organizational directory only* |
| Redirect URI | leave blank |

Copy from **Overview**:

- **Application (client) ID** → `AZURE_CLIENT_ID`
- **Directory (tenant) ID** → `AZURE_TENANT_ID`

### 2. Add the API permission

**API permissions** → **Add a permission** → **Microsoft Graph** →
**Application permissions** → tick **`Files.Read.All`** → **Add permissions**.

Then **Grant admin consent**. The status column must read *Granted*.

> **Why `Files.Read.All` and not something narrower?**
> `Sites.Selected` is least-privilege and preferable, but it only covers
> SharePoint **sites**. The stock report currently lives on a personal OneDrive,
> which `Sites.Selected` cannot reach. Move that file to the team site and the
> permission can be narrowed — see *Hardening*.

Do **not** add `Files.ReadWrite.All`. The job never writes to SharePoint.

### 3. Create the client secret

**Certificates & secrets** → **New client secret** → expires **24 months**.

Copy the **Value** (not the Secret ID) → `AZURE_CLIENT_SECRET`. Shown once only.

> ⚠️ Set a reminder a month before expiry. When it lapses the report stops
> updating and shows a stale-data banner.

### 4. Store the values

Locally, in a gitignored `.env` at the repo root (see `.env.example`):

```
AZURE_TENANT_ID=...
AZURE_CLIENT_ID=...
AZURE_CLIENT_SECRET=...
```

For GitHub Actions:

```bash
gh secret set AZURE_TENANT_ID
gh secret set AZURE_CLIENT_ID
gh secret set AZURE_CLIENT_SECRET
```

### 5. Verify

```powershell
python -m ingest --check
```

Expected:

```
Auth: app - signed in as application (app-only)

OK    Safripol Stock Report - TAC IMOLA   (2026-09-10T11:38:37Z, 244 KB)
OK    Safripol v Connect Dwells           (2026-09-10T13:01:01Z, 231 KB)
OK    SAF Ops Tracking - TAC IMOLA        (2026-09-10T12:20:42Z, 1049 KB)
OK    Saf Staff Roles                     (2026-09-10T11:17:20Z, 19 KB)
```

Then re-enable the schedule in `.github/workflows/refresh-data.yml` and remove
the local scheduled task with `tools\install_task.ps1 -Remove`.

---

## Troubleshooting

| Symptom | Cause | Fix |
|---|---|---|
| `CERTIFICATE_VERIFY_FAILED` | corporate TLS proxy; Python does not trust the internal CA | handled automatically by `truststore`; ensure `pip install -r requirements.txt` has run |
| `401 generalException` with `roles: null` | only Delegated was granted | use delegated mode, or get an Application grant |
| `AADSTS7000215` invalid client secret | Secret ID copied instead of Value | recreate the secret, copy the **Value** |
| `403 accessDenied` | admin consent not granted | grant consent |
| `404 itemNotFound` | file renamed or moved | update the URL in `ingest/config.py` |
| `No usable cached sign-in` | token cache missing or expired | `python -m ingest --login` |
| Report timestamp is old | refresh machine off, or task failed | check `refresh.log` and the task's `LastTaskResult` |

## Hardening (recommended follow-up)

1. Move `Safripol Stock Report - TAC IMOLA.xlsx` off personal OneDrive onto the
   **Process Optimization and Development** team site. A personal OneDrive is a
   single point of failure — if that account is disabled, the report loses its
   most important source.
2. Once moved, switch to **`Sites.Selected`** and grant the app `read` on just
   that site.
3. Replace the client secret with **federated credentials (OIDC)** so GitHub
   Actions authenticates with nothing stored and nothing to expire.

