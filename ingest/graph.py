"""Microsoft Graph access for the Safripol ingest.

Authentication is app-only (client credentials) against an Entra app registration.
Files are fetched through the /shares endpoint, which accepts the ordinary SharePoint
or OneDrive URL and therefore works for both the personal-OneDrive stock report and
the team-site workbooks without any site/drive id resolution.

Required Graph application permission: Files.Read.All (admin consented).
Optional least-privilege alternative for team sites only: Sites.Selected.
"""

from __future__ import annotations

import base64
import logging
import os
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from pathlib import Path
from urllib.parse import quote, unquote, urlparse

import msal
import requests

from .config import CACHE_DIR, Source, settings

log = logging.getLogger(__name__)


def _use_system_trust_store() -> None:
    """Trust the certificates Windows already trusts.

    Connect's network terminates TLS at a corporate proxy, so Microsoft endpoints
    present a certificate signed by an internal CA. Windows trusts it; Python's
    bundled certifi store does not, which surfaces as CERTIFICATE_VERIFY_FAILED.
    Deferring to the OS store fixes that without ever disabling verification.
    Silently ignored where it is unnecessary, such as on the CI runner.
    """
    try:
        import truststore

        truststore.inject_into_ssl()
    except Exception as exc:  # noqa: BLE001
        log.debug("System trust store unavailable (%s); using certifi.", exc)


_use_system_trust_store()

GRAPH = "https://graph.microsoft.com/v1.0"
SCOPE = ["https://graph.microsoft.com/.default"]
# Delegated sign-in asks for the specific permission rather than /.default, so the
# consent prompt states plainly that this only ever reads files.
DELEGATED_SCOPE = ["Files.Read.All"]
TIMEOUT = 180

# Refresh tokens are what let the scheduled refresh run without a human. Keep the
# cache outside the repo so it can never be committed, and readable only by this
# Windows/Linux user account.
TOKEN_CACHE_PATH = Path(
    os.getenv("SAFRIPOL_TOKEN_CACHE")
    or (Path.home() / ".safripol_report" / "token_cache.json")
)


def _load_cache() -> msal.SerializableTokenCache:
    cache = msal.SerializableTokenCache()
    if TOKEN_CACHE_PATH.exists():
        try:
            cache.deserialize(TOKEN_CACHE_PATH.read_text(encoding="utf-8"))
        except Exception as exc:  # noqa: BLE001
            log.warning("Token cache unreadable (%s); a new sign-in is needed.", exc)
    return cache


def _save_cache(cache: msal.SerializableTokenCache) -> None:
    if not cache.has_state_changed:
        return
    TOKEN_CACHE_PATH.parent.mkdir(parents=True, exist_ok=True)
    TOKEN_CACHE_PATH.write_text(cache.serialize(), encoding="utf-8")
    try:
        os.chmod(TOKEN_CACHE_PATH, 0o600)
    except OSError:
        pass


class GraphError(RuntimeError):
    pass


@dataclass
class Fetched:
    key: str
    label: str
    path: Path
    last_modified: datetime | None
    size: int
    origin: str          # "graph" | "local" | "cache"
    ok: bool = True
    error: str = ""


def _encode_share_url(url: str) -> str:
    """Graph's share-id encoding: base64url of the URL, 'u!' prefixed, '=' stripped."""
    b64 = base64.urlsafe_b64encode(url.encode("utf-8")).decode("ascii")
    return "u!" + b64.rstrip("=")


def _split_sharepoint_url(url: str) -> tuple[str, str, str] | None:
    """Split a SharePoint/OneDrive URL into (hostname, site path, file path).

    Needed for Sites.Selected, where the app is entitled to named sites rather than
    to every file in the tenant. The /shares endpoint resolves a link without ever
    naming a site, so it can be refused under that model; addressing the site and
    then the drive path explicitly is what Sites.Selected is designed to allow.

    Handles team sites (/sites/<name>/<library>/<path>) and personal OneDrive
    (/personal/<user>/Documents/<path>), which is a site collection of its own.
    """
    parsed = urlparse(url)
    host = parsed.netloc
    parts = [unquote(p) for p in parsed.path.split("/") if p]
    if not host or len(parts) < 3 or parts[0].lower() not in {"sites", "teams", "personal"}:
        return None

    site_path = f"/{parts[0]}/{parts[1]}"
    rest = parts[2:]
    if len(rest) < 2:
        return None
    # The first remaining segment is the document library ("Shared Documents",
    # "Documents"); Graph's /drive already points at the default library.
    return host, site_path, "/".join(rest[1:])


class GraphClient:
    """Talks to Graph as either the application or a signed-in user.

    App-only is preferable for an unattended job, but it needs an *Application*
    permission grant. Where only *Delegated* permissions are available, the
    client falls back to signing a user in once via device code and reusing the
    cached refresh token from then on. Graph then sees exactly the access that
    user already has in SharePoint - nothing more.
    """

    def __init__(self, prefer: str = "") -> None:
        mode = prefer or settings.auth_mode
        if mode == "app" or (mode != "delegated" and settings.has_graph_credentials):
            self._mode = "app"
        else:
            self._mode = "delegated"

        if self._mode == "app":
            if not settings.has_graph_credentials:
                raise GraphError(
                    "Missing Entra credentials. Set AZURE_TENANT_ID, AZURE_CLIENT_ID "
                    "and AZURE_CLIENT_SECRET."
                )
            self._app = msal.ConfidentialClientApplication(
                client_id=settings.client_id,
                client_credential=settings.client_secret,
                authority=settings.authority,
            )
        else:
            if not (settings.tenant_id and settings.client_id):
                raise GraphError(
                    "Missing Entra details. Set AZURE_TENANT_ID and AZURE_CLIENT_ID."
                )
            self._cache = _load_cache()
            self._app = msal.PublicClientApplication(
                client_id=settings.client_id,
                authority=settings.authority,
                token_cache=self._cache,
            )
        self._token: str | None = None

    # -- authentication -----------------------------------------------------

    def _acquire(self) -> str:
        if self._token:
            return self._token
        result = (
            self._acquire_app() if self._mode == "app" else self._acquire_delegated()
        )
        if "access_token" not in result:
            raise GraphError(
                f"Token request failed: {result.get('error')} - "
                f"{result.get('error_description', '')[:400]}"
            )
        self._token = result["access_token"]
        return self._token

    def _acquire_app(self) -> dict:
        result = self._app.acquire_token_silent(SCOPE, account=None)
        return result or self._app.acquire_token_for_client(scopes=SCOPE)

    def _acquire_delegated(self) -> dict:
        accounts = self._app.get_accounts()
        if accounts:
            result = self._app.acquire_token_silent(DELEGATED_SCOPE, account=accounts[0])
            if result and "access_token" in result:
                _save_cache(self._cache)
                return result

        if not settings.allow_interactive:
            raise GraphError(
                "No usable cached sign-in. Run 'python -m ingest --login' on this "
                "machine to authorise it."
            )

        flow = self._app.initiate_device_flow(scopes=DELEGATED_SCOPE)
        if "user_code" not in flow:
            raise GraphError(
                f"Device flow failed: {flow.get('error')} - "
                f"{flow.get('error_description', '')[:300]}"
            )
        print("\n" + "=" * 68)
        print("  Authorise the Safripol report to read SharePoint on your behalf")
        print("=" * 68)
        print(f"\n  1. Open   : {flow['verification_uri']}")
        print(f"  2. Enter  : {flow['user_code']}")
        print("  3. Sign in as yourself and approve.\n")
        print("  Needed once per machine. The refresh token is then cached locally")
        print("  and renews itself on every run.\n")
        result = self._app.acquire_token_by_device_flow(flow)
        _save_cache(self._cache)
        return result

    def signed_in_as(self) -> str:
        if self._mode == "app":
            return "application (app-only)"
        accounts = self._app.get_accounts()
        return accounts[0].get("username", "unknown") if accounts else "not signed in"

    @property
    def mode(self) -> str:
        return self._mode

    def _headers(self) -> dict[str, str]:
        return {"Authorization": f"Bearer {self._acquire()}"}

    def item_metadata(self, url: str) -> dict:
        if self._mode == "app":
            direct = self._item_metadata_by_site_path(url)
            if direct is not None:
                return direct
        share_id = _encode_share_url(url)
        r = requests.get(
            f"{GRAPH}/shares/{share_id}/driveItem",
            headers=self._headers(),
            timeout=TIMEOUT,
        )
        if r.status_code != 200:
            raise GraphError(f"Metadata {r.status_code}: {r.text[:400]}")
        return r.json()

    def _item_metadata_by_site_path(self, url: str) -> dict | None:
        """Resolve an item as site -> drive -> path, the Sites.Selected-friendly route.

        Returns None when the URL is not addressable this way or the lookup fails,
        so the caller can still try /shares rather than losing the source outright.
        """
        split = _split_sharepoint_url(url)
        if split is None:
            return None
        host, site_path, file_path = split
        try:
            site = requests.get(
                f"{GRAPH}/sites/{host}:{site_path}",
                headers=self._headers(),
                timeout=TIMEOUT,
            )
            if site.status_code != 200:
                log.debug("Site lookup %s: %s", site.status_code, site.text[:200])
                return None
            site_id = site.json().get("id")
            item = requests.get(
                f"{GRAPH}/sites/{site_id}/drive/root:/{quote(file_path)}",
                headers=self._headers(),
                timeout=TIMEOUT,
            )
            if item.status_code != 200:
                log.debug("Item lookup %s: %s", item.status_code, item.text[:200])
                return None
            return item.json()
        except requests.RequestException as exc:
            log.debug("Site-path resolution failed: %s", exc)
            return None

    def download(self, source: Source, dest_dir: Path) -> Fetched:
        meta = self.item_metadata(source.url)
        modified = meta.get("lastModifiedDateTime")
        last_modified = (
            datetime.fromisoformat(modified.replace("Z", "+00:00")) if modified else None
        )

        # A resolved item carries its own pre-authorised download URL. Using it keeps
        # the fetch on whichever route already worked for metadata, which matters
        # under Sites.Selected where /shares may not be permitted.
        direct_url = meta.get("@microsoft.graph.downloadUrl")
        if direct_url:
            r = requests.get(direct_url, timeout=TIMEOUT, allow_redirects=True)
        else:
            share_id = _encode_share_url(source.url)
            r = requests.get(
                f"{GRAPH}/shares/{share_id}/driveItem/content",
                headers=self._headers(),
                timeout=TIMEOUT,
                allow_redirects=True,
            )
        if r.status_code != 200:
            raise GraphError(f"Download {r.status_code}: {r.text[:400]}")

        dest_dir.mkdir(parents=True, exist_ok=True)
        path = dest_dir / f"{source.key}.xlsx"
        path.write_bytes(r.content)
        log.info("Downloaded %s (%s bytes)", source.label, len(r.content))
        return Fetched(
            key=source.key,
            label=source.label,
            path=path,
            last_modified=last_modified,
            size=len(r.content),
            origin="graph",
        )


def _local_lookup(source: Source) -> Path | None:
    for base in settings.local_dirs:
        if not base.exists():
            continue
        candidate = base / source.local_hint
        if candidate.exists():
            return candidate
        # tolerate the trailing-space filenames that exist in the shared folder
        stem = source.local_hint.strip().lower().replace(" ", "")
        for p in base.glob("*.xlsx"):
            if p.name.strip().lower().replace(" ", "") == stem:
                return p
    return None


def _copy_local_source(source: Source, local: Path) -> Fetched:
    import shutil

    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    dest = CACHE_DIR / f"{source.key}.xlsx"
    shutil.copy2(local, dest)   # copy: OneDrive locks the original
    stat = local.stat()
    return Fetched(
        key=source.key,
        label=source.label,
        path=dest,
        last_modified=datetime.fromtimestamp(stat.st_mtime, tz=timezone.utc),
        size=stat.st_size,
        origin="local",
    )


def fetch_all(sources: list[Source]) -> dict[str, Fetched]:
    """Fetch every source, preferring the freshest available copy.

    A failure on an optional source is recorded rather than raised, so the report
    still refreshes when, for example, the ops tracker has not been uploaded yet.
    When a OneDrive-synced local workbook is newer than Graph metadata, use it:
    SharePoint/Graph can lag the desktop sync client by several minutes.
    """
    out: dict[str, Fetched] = {}
    client: GraphClient | None = None
    if not settings.offline and settings.has_graph_credentials:
        try:
            client = GraphClient()
        except GraphError as exc:
            log.warning("Graph unavailable, falling back to local files: %s", exc)

    for source in sources:
        fetched: Fetched | None = None
        if client is not None:
            try:
                fetched = client.download(source, CACHE_DIR)
            except Exception as exc:  # noqa: BLE001 - we degrade rather than fail
                log.warning("Graph fetch failed for %s: %s", source.label, exc)

        local = _local_lookup(source)
        if local is not None:
            local_mtime = datetime.fromtimestamp(local.stat().st_mtime, tz=timezone.utc)
            if fetched is None:
                fetched = _copy_local_source(source, local)
            elif fetched.last_modified and local_mtime > fetched.last_modified + timedelta(seconds=60):
                log.info(
                    "Using newer local copy for %s (local %s > graph %s)",
                    source.label, local_mtime.isoformat(), fetched.last_modified.isoformat(),
                )
                fetched = _copy_local_source(source, local)

        if fetched is None:
            cached = CACHE_DIR / f"{source.key}.xlsx"
            if cached.exists():
                stat = cached.stat()
                fetched = Fetched(
                    key=source.key, label=source.label, path=cached,
                    last_modified=datetime.fromtimestamp(stat.st_mtime, tz=timezone.utc),
                    size=stat.st_size, origin="cache",
                )

        if fetched is None:
            fetched = Fetched(
                key=source.key, label=source.label, path=Path(),
                last_modified=None, size=0, origin="none", ok=False,
                error="Not reachable via Graph and no local or cached copy found.",
            )
            if source.required:
                log.error("Required source unavailable: %s", source.label)

        out[source.key] = fetched
    return out
