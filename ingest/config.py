"""Configuration for the Safripol Operations KPI ingest.

Every source is addressed by its SharePoint / OneDrive sharing URL. Microsoft Graph
resolves those directly via the /shares endpoint, so no site-id or drive-id lookups
are needed and the URLs stay identical to the ones used by Power Query.
"""

from __future__ import annotations

import os
from dataclasses import dataclass, field
from pathlib import Path
from zoneinfo import ZoneInfo

REPO_ROOT = Path(__file__).resolve().parent.parent
DOCS_DIR = REPO_ROOT / "docs"
DATA_DIR = DOCS_DIR / "data"
CACHE_DIR = REPO_ROOT / ".cache"

SAST = ZoneInfo("Africa/Johannesburg")

# --- Business constants (mirror of the Power BI model) ----------------------
SHIPMENT = "TAC IMOLA"
VESSEL_LABEL = "MV TAC IMOLA"
TARGET_TOTAL_MT = 20400.00      # Safripol vessel outturn / drawdown basis
BAG_WEIGHT_MT = 1.2
TARGET_LOAD_KG = 21600
ALLOWED_DECANT_HRS = 1.75
ALLOWED_INTERVAL_HRS = 1.00
SITE_TARGET_HRS = 2.75
TRANSIT_TARGET_HRS = 0.75

# Vessel discharge outturn is reported off the hatch table, NOT off the delivery target.
HATCHES = [
    # hatch, start, end, volume_mt, gross_h, downtime_h, weather_h, net_h
    ("#1", "2026-09-01T02:45", "2026-09-02T00:05", 2405.40, 21.3, 0.9, 0.0, 20.5),
    ("#2", "2026-09-02T00:47", "2026-09-03T14:20", 4810.79, 37.5, 1.1, 0.0, 36.4),
    ("#3", "2026-09-01T02:38", "2026-09-03T00:45", 4814.40, 46.1, 2.2, 0.0, 43.9),
    ("#4", "2026-09-02T12:10", "2026-09-04T10:55", 4814.40, 46.8, 6.7, 5.8167, 40.0),
    ("#5", "2026-09-01T01:57", "2026-09-02T17:00", 3601.48, 39.0, 1.0, 0.0, 38.0),
]

WEATHER_EVENTS = [
    ("Rain", 1, "2026-09-04T02:54", "2026-09-04T08:43", 5.8167),
]

SITE_CONNECT = "Connect Logistics"
SITE_SAFRIPOL = "Safripol"

# Normalises the several spellings of the isotainer reference seen across sources.
ISO_ALIASES = {
    "ISOT 001": "ISO-01", "ISOT 002": "ISO-02", "ISOT 003": "ISO-03",
    "ISOT 004": "ISO-04", "ISOT 005": "ISO-05",
    "ISOT  001": "ISO-01", "ISOT  002": "ISO-02", "ISOT  003": "ISO-03",
    "ISOT  004": "ISO-04", "ISOT  005": "ISO-05",
    "ISO 01": "ISO-01", "ISO 02": "ISO-02", "ISO 03": "ISO-03",
    "ISO 04": "ISO-04", "ISO 05": "ISO-05",
}


@dataclass(frozen=True)
class Source:
    key: str
    label: str
    url: str
    required: bool = True
    local_hint: str = ""


SOURCES: list[Source] = [
    Source(
        key="stock",
        label="Safripol Stock Report - TAC IMOLA",
        url=(
            "https://connectlogisticscoza.sharepoint.com/sites/"
            "ProcessOptimizationandDevelopment/Shared%20Documents/Safripol/2026/"
            "Safripol%20Stock%20Report%20-%20TAC%20IMOLA.xlsx"
        ),
        local_hint="Safripol Stock Report - TAC IMOLA.xlsx",
    ),
    Source(
        key="dwells",
        label="Safripol v Connect Dwells",
        url=(
            "https://connectlogisticscoza.sharepoint.com/sites/"
            "ProcessOptimizationandDevelopment/Shared%20Documents/Power%20BI/"
            "Safripol%20Reports/Safripol%20v%20Connect%20Dwells.xlsx"
        ),
        local_hint="Safripol v Connect Dwells.xlsx",
    ),
    Source(
        key="tracking",
        label="SAF Ops Tracking - TAC IMOLA",
        url=(
            "https://connectlogisticscoza.sharepoint.com/sites/"
            "ProcessOptimizationandDevelopment/Shared%20Documents/Power%20BI/"
            "Safripol%20Reports/SAF%20Ops%20Tracking%20-%20TAC%20IMOLA.xlsx"
        ),
        required=False,
        local_hint="SAF Ops Tracking - TAC IMOLA.xlsx",
    ),
    Source(
        key="staff",
        label="Saf Staff Roles",
        url=(
            "https://connectlogisticscoza.sharepoint.com/sites/"
            "ProcessOptimizationandDevelopment/Shared%20Documents/Power%20BI/"
            "Safripol%20Reports/Saf%20Staff%20Roles%20.xlsx"
        ),
        required=False,
        local_hint="Saf Staff Roles .xlsx",
    ),
    Source(
        key="plan",
        label="PTA Drawdown Plan",
        url=(
            "https://connectlogisticscoza.sharepoint.com/sites/"
            "ProcessOptimizationandDevelopment/Shared%20Documents/Power%20BI/"
            "Safripol%20Reports/PTA_Drawdown_Plan.xlsx"
        ),
        required=False,
        local_hint="PTA_Drawdown_Plan.xlsx",
    ),
]

SOURCE_BY_KEY = {s.key: s for s in SOURCES}


def _load_dotenv() -> None:
    """Read key=value pairs from a local .env into the environment.

    Lets the credentials sit in one gitignored file for local runs, while
    GitHub Actions supplies the same names as real environment variables.
    Anything already set in the environment wins, so CI is never overridden.
    """
    path = REPO_ROOT / ".env"
    if not path.exists():
        return
    for raw in path.read_text(encoding="utf-8-sig").splitlines():
        line = raw.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, _, value = line.partition("=")
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        if key and key not in os.environ:
            os.environ[key] = value


_load_dotenv()


@dataclass
class Settings:
    tenant_id: str = field(default_factory=lambda: os.getenv("AZURE_TENANT_ID", ""))
    client_id: str = field(default_factory=lambda: os.getenv("AZURE_CLIENT_ID", ""))
    client_secret: str = field(default_factory=lambda: os.getenv("AZURE_CLIENT_SECRET", ""))
    # One or more folders, separated by ';'. Searched in order when Graph cannot
    # be reached, so the sources may live in different synced libraries.
    local_dir: str = field(default_factory=lambda: os.getenv("SAFRIPOL_LOCAL_DIR", ""))
    offline: bool = field(default_factory=lambda: os.getenv("SAFRIPOL_OFFLINE", "") == "1")
    # "app" for client credentials, "delegated" to run as a signed-in user, or
    # "" to pick automatically based on whether a client secret is present.
    auth_mode: str = field(
        default_factory=lambda: os.getenv("SAFRIPOL_AUTH_MODE", "").strip().lower()
    )
    # Device-code sign-in needs a human, so it is only offered when explicitly
    # allowed. Scheduled runs must rely on the cached refresh token.
    allow_interactive: bool = False

    @property
    def local_dirs(self) -> list[Path]:
        return [
            Path(p.strip())
            for p in self.local_dir.split(";")
            if p.strip()
        ]

    @property
    def authority(self) -> str:
        return f"https://login.microsoftonline.com/{self.tenant_id}"

    @property
    def has_graph_credentials(self) -> bool:
        return bool(self.tenant_id and self.client_id and self.client_secret)


settings = Settings()
