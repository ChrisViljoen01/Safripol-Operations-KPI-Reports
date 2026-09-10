# Entra (Azure AD) setup

The refresh job signs in as an **application**, not as a person. No user is
prompted, nothing expires when someone leaves, and the report keeps working
overnight.

---

## 1. Create the app registration

Azure portal → **Microsoft Entra ID** → **App registrations** → **New registration**

| Field | Value |
|---|---|
| Name | `Safripol Ops Report – Data Refresh` |
| Supported account types | *Accounts in this organizational directory only* |
| Redirect URI | leave blank |

Register, then copy from the **Overview** page:

- **Application (client) ID** → `AZURE_CLIENT_ID`
- **Directory (tenant) ID** → `AZURE_TENANT_ID`

## 2. Add the API permission

**API permissions** → **Add a permission** → **Microsoft Graph** →
**Application permissions** → tick **`Files.Read.All`** → **Add permissions**.

Then click **Grant admin consent for Connect Logistics**. The status column must
read *Granted* — without this the refresh fails with `403`.

> **Why `Files.Read.All` and not something narrower?**
> `Sites.Selected` is the least-privilege option and would be preferable, but it
> only covers SharePoint **sites**. The stock report
> (`Safripol Stock Report - TAC IMOLA.xlsx`) currently lives on Richard's
> personal OneDrive, which `Sites.Selected` cannot reach.
>
> If that file is moved onto the Process Optimization team site, we can downgrade
> to `Sites.Selected` and grant the app access to that one site only. That is the
> recommended end state — see *Hardening* below.

Do **not** add `Files.ReadWrite.All`. The job never writes to SharePoint, and the
read-only grant is the guarantee that it cannot.

## 3. Create the client secret

**Certificates & secrets** → **Client secrets** → **New client secret**

- Description: `github-actions-refresh`
- Expires: **24 months**

Copy the **Value** immediately (not the Secret ID) → `AZURE_CLIENT_SECRET`.
It is only shown once.

> ⚠️ Put a calendar reminder in for one month before the expiry date. When the
> secret lapses the report silently stops updating — the page keeps serving the
> last snapshot and shows a stale-data banner.

## 4. Store the three values in GitHub

Repo → **Settings** → **Secrets and variables** → **Actions** → **New repository secret**

| Secret name | Value |
|---|---|
| `AZURE_TENANT_ID` | Directory (tenant) ID |
| `AZURE_CLIENT_ID` | Application (client) ID |
| `AZURE_CLIENT_SECRET` | the secret **Value** from step 3 |

Or from the CLI:

```bash
gh secret set AZURE_TENANT_ID
gh secret set AZURE_CLIENT_ID
gh secret set AZURE_CLIENT_SECRET
```

## 5. Verify

Locally:

```powershell
$env:AZURE_TENANT_ID="..."; $env:AZURE_CLIENT_ID="..."; $env:AZURE_CLIENT_SECRET="..."
python -m ingest --check
```

Expected:

```
OK    Safripol Stock Report - TAC IMOLA   (2026-09-10T09:14:02Z, 412 KB)
OK    Safripol v Connect Dwells           (2026-09-10T08:55:11Z, 288 KB)
OK    SAF Ops Tracking - TAC IMOLA        (2026-09-10T09:20:44Z, 196 KB)
OK    Saf Staff Roles                     (2026-08-02T11:02:19Z, 22 KB)
```

Then trigger the workflow: **Actions** → *Refresh report data* → **Run workflow**.

---

## Troubleshooting

| Symptom | Cause | Fix |
|---|---|---|
| `AADSTS7000215` invalid client secret | secret value wrong, or the Secret ID was copied instead of the Value | recreate the secret, copy the **Value** column |
| `403 accessDenied` | admin consent not granted | step 2, click *Grant admin consent* |
| `404 itemNotFound` | the file was renamed, moved, or the sharing URL changed | update the URL in `ingest/config.py` |
| Refresh works locally, fails in Actions | secrets not set on the repo | step 4 |
| `WARN` on `SAF Ops Tracking` | the rebuilt tracker has not been uploaded to SharePoint yet | upload it to the path in `config.py` |

## Hardening (recommended follow-up)

1. Move `Safripol Stock Report - TAC IMOLA.xlsx` from personal OneDrive onto the
   **Process Optimization and Development** team site. Personal OneDrive is a
   single point of failure — if that account is disabled, the report dies.
2. Once moved, switch the permission to **`Sites.Selected`** and grant the app
   `read` on just that site:
   ```
   POST /v1.0/sites/{site-id}/permissions
   { "roles": ["read"], "grantedToIdentities": [ { "application":
     { "id": "<client-id>", "displayName": "Safripol Ops Report" } } ] }
   ```
3. Replace the client secret with a **certificate** or, better, **federated
   credentials** (OIDC) so GitHub Actions authenticates with no stored secret at
   all and nothing to expire.
