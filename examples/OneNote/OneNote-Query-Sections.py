#!/usr/bin/env python3

"""Query OneNote notebook sections with Microsoft Graph.

This script supports two query modes:

1. Personal notebooks through ``/me/onenote/notebooks/{id}/sections``
2. SharePoint-hosted team notebooks through ``/sites/{site-id}/onenote/sections``

It can also resolve a SharePoint site ID from hostname and server-relative path
before querying the site's OneNote sections.

Authentication uses MSAL public-client interactive sign-in with a local token
cache stored next to this script in ``.onenote-token-cache.json``. On each run,
the script tries to reuse a cached account silently before opening an
interactive sign-in flow.

Required environment variables:
    TENANT_ID             Entra tenant ID or tenant domain
    CLIENT_ID             App registration client ID

Optional environment variables:
    LOGIN_HINT            Preferred account for interactive sign-in
    ONENOTE_NOTEBOOK_ID   Notebook ID to query by default
    ONENOTE_SITE_ID       SharePoint site ID for team notebooks
    ONENOTE_SITE_HOSTNAME SharePoint hostname for resolving a site ID
    ONENOTE_SITE_PATH     SharePoint server-relative path for resolving a site ID

Required delegated Microsoft Graph permissions:
    - Notes.Read.All
    - Sites.Read.All when using --site-hostname and --site-path

The app registration must also allow public client interactive sign-in.

Examples:
    python examples/OneNote/OneNote-Query-Sections.py
    python examples/OneNote/OneNote-Query-Sections.py --notebook-name "Team Notes"
    python examples/OneNote/OneNote-Query-Sections.py --notebook-id <notebook-id>
    python examples/OneNote/OneNote-Query-Sections.py --site-id <site-id>
    python examples/OneNote/OneNote-Query-Sections.py --site-id <site-id> --notebook-name "General"
    python examples/OneNote/OneNote-Query-Sections.py --site-hostname contoso.sharepoint.com --site-path /sites/TeamA
    python examples/OneNote/OneNote-Query-Sections.py --login-hint user@contoso.com
"""

from __future__ import annotations

import argparse
import os
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional
from urllib.parse import quote

try:
    import requests
    from dotenv import load_dotenv
    from msal import PublicClientApplication, SerializableTokenCache
except ImportError as exc:
    package_name = getattr(exc, "name", "a required package")
    raise SystemExit(
        "Missing dependency: "
        f"{package_name}. Install with: pip install msal requests python-dotenv"
    ) from exc


GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"
NOTES_GRAPH_SCOPES = ["Notes.Read.All"]
SITE_RESOLUTION_SCOPES = ["Sites.Read.All"]
TOKEN_CACHE_PATH = Path(__file__).with_name(".onenote-token-cache.json")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="List the sections of a OneNote notebook using Microsoft Graph."
    )
    parser.add_argument("--login-hint", dest="login_hint", default=os.getenv("LOGIN_HINT"))
    parser.add_argument("--verbose", action="store_true")
    parser.add_argument(
        "--site-id",
        default=os.getenv("ONENOTE_SITE_ID"),
        help="SharePoint site ID for team notebooks. Defaults to ONENOTE_SITE_ID.",
    )
    parser.add_argument(
        "--site-hostname",
        default=os.getenv("ONENOTE_SITE_HOSTNAME") or os.getenv("DEFAULT_HOST"),
        help="SharePoint hostname used to resolve a site ID. Defaults to ONENOTE_SITE_HOSTNAME or DEFAULT_HOST.",
    )
    parser.add_argument(
        "--site-path",
        default=os.getenv("ONENOTE_SITE_PATH") or os.getenv("SITE_PATH"),
        help="SharePoint server-relative path used to resolve a site ID. Defaults to ONENOTE_SITE_PATH or SITE_PATH.",
    )
    parser.add_argument(
        "--notebook-id",
        default=os.getenv("ONENOTE_NOTEBOOK_ID"),
        help="Notebook ID to query. Defaults to ONENOTE_NOTEBOOK_ID.",
    )
    parser.add_argument(
        "--notebook-name",
        help="Notebook display name to resolve before listing sections.",
    )
    return parser.parse_args()


def _stderr(msg: str) -> None:
    sys.stderr.write(msg + "\n")


def load_token_cache() -> SerializableTokenCache:
    cache = SerializableTokenCache()
    if TOKEN_CACHE_PATH.exists():
        cache.deserialize(TOKEN_CACHE_PATH.read_text(encoding="utf-8"))
    return cache


def save_token_cache(cache: SerializableTokenCache) -> None:
    if cache.has_state_changed:
        TOKEN_CACHE_PATH.write_text(cache.serialize(), encoding="utf-8")


def get_required_env(name: str) -> str:
    value = os.getenv(name)
    if not value:
        raise SystemExit(f"Missing required environment variable: {name}")
    return value


def normalize_site_path(site_path: str) -> str:
    normalized_path = site_path.strip()
    if not normalized_path:
        raise SystemExit("Site path cannot be empty.")
    if not normalized_path.startswith("/"):
        normalized_path = f"/{normalized_path}"
    if normalized_path != "/":
        normalized_path = normalized_path.rstrip("/")
    return normalized_path


def graph_get(access_token: str, path: str, params: dict[str, str] | None = None) -> dict:
    response = requests.get(
        f"{GRAPH_BASE_URL}{path}",
        headers={
            "Authorization": f"Bearer {access_token}",
            "Accept": "application/json",
        },
        params=params,
        timeout=30,
    )

    if not response.ok:
        try:
            error_payload = response.json()
        except ValueError:
            error_payload = response.text
        raise SystemExit(f"Graph request failed ({response.status_code}): {error_payload}")

    return response.json()


def list_notebooks(access_token: str) -> list[dict]:
    payload = graph_get(
        access_token,
        "/me/onenote/notebooks",
        params={"$select": "id,displayName,self"},
    )
    return payload.get("value", [])


def list_site_notebooks(access_token: str, site_id: str) -> list[dict]:
    payload = graph_get(
        access_token,
        f"/sites/{site_id}/onenote/notebooks",
        params={"$select": "id,displayName,self"},
    )
    return payload.get("value", [])


def resolve_site(access_token: str, site_hostname: str, site_path: str) -> dict:
    normalized_path = normalize_site_path(site_path)
    encoded_path = quote(normalized_path, safe="/")
    payload = graph_get(
        access_token,
        f"/sites/{site_hostname}:{encoded_path}",
        params={"$select": "id,displayName,webUrl"},
    )
    payload["resolvedPath"] = normalized_path
    return payload


def get_sections(access_token: str, notebook_id: str) -> list[dict]:
    payload = graph_get(
        access_token,
        f"/me/onenote/notebooks/{notebook_id}/sections",
        params={"$select": "id,displayName,createdDateTime,lastModifiedDateTime"},
    )
    return payload.get("value", [])


def get_site_sections(access_token: str, site_id: str) -> list[dict]:
    payload = graph_get(
        access_token,
        f"/sites/{site_id}/onenote/sections",
        params={
            "$select": "id,displayName,createdDateTime,lastModifiedDateTime",
            "$expand": "parentNotebook($select=id,displayName)",
        },
    )
    return payload.get("value", [])


def filter_sections_by_notebook(
    sections: list[dict], notebook_id: str | None, notebook_name: str | None
) -> list[dict]:
    if notebook_id:
        return [
            section
            for section in sections
            if section.get("parentNotebook", {}).get("id") == notebook_id
        ]

    if notebook_name:
        return [
            section
            for section in sections
            if section.get("parentNotebook", {}).get("displayName") == notebook_name
        ]

    return sections


def resolve_notebook(
    access_token: str,
    notebook_id: str | None,
    notebook_name: str | None,
    site_id: str | None,
) -> dict | None:
    notebooks = list_site_notebooks(access_token, site_id) if site_id else list_notebooks(access_token)

    if notebook_id:
        for notebook in notebooks:
            if notebook.get("id") == notebook_id:
                return notebook
        raise SystemExit(f"Notebook ID not found: {notebook_id}")

    if notebook_name:
        matches = [
            notebook for notebook in notebooks if notebook.get("displayName") == notebook_name
        ]
        if not matches:
            raise SystemExit(f"Notebook name not found: {notebook_name}")
        if len(matches) > 1:
            raise SystemExit(
                f"Notebook name is ambiguous: {notebook_name}. Use --notebook-id instead."
            )
        return matches[0]

    if not notebooks:
        if site_id:
            print("No OneNote notebooks were returned for the specified SharePoint site.")
        else:
            print("No OneNote notebooks were returned for the signed-in user.")
        return None

    if site_id:
        print("Available site notebooks:")
    else:
        print("Available notebooks:")
    for notebook in notebooks:
        print(f"- {notebook['displayName']}: {notebook['id']}")
    print("\nRe-run with --notebook-id or --notebook-name to list sections.")
    return None


def print_sections(sections: list[dict], include_parent_notebook: bool = False) -> int:
    print("Sections:")

    if not sections:
        print("- No sections found.")
        return 0

    for section in sections:
        parent_notebook = section.get("parentNotebook", {})
        notebook_prefix = ""
        if include_parent_notebook:
            notebook_prefix = f"notebook={parent_notebook.get('displayName', 'n/a')} | "

        print(
            "- "
            f"{section['displayName']} | "
            f"{notebook_prefix}"
            f"id={section['id']} | "
            f"created={section.get('createdDateTime', 'n/a')} | "
            f"modified={section.get('lastModifiedDateTime', 'n/a')}"
        )

    return 0


# -----------------------------
# Auth
# -----------------------------
@dataclass
class AuthConfig:
    tenant_id: str
    client_id: str
    scopes: List[str]
    login_hint: Optional[str] = None
    verbose: bool = False


class MgAuth:
    """Acquire delegated Graph tokens with MSAL and a persisted local cache."""

    def __init__(self, cfg: AuthConfig) -> None:
        self.cfg = cfg
        cache_dir = os.path.dirname(TOKEN_CACHE_PATH)
        os.makedirs(cache_dir, exist_ok=True)
        self.cache = SerializableTokenCache()
        # Load existing cache
        if os.path.exists(TOKEN_CACHE_PATH):
            try:
                self.cache.deserialize(Path(TOKEN_CACHE_PATH).read_text())
            except Exception:
                # Corrupt cache: start fresh
                self.cache = SerializableTokenCache()
        authority = f"https://login.microsoftonline.com/{cfg.tenant_id}"
        self.app = PublicClientApplication(
            client_id=cfg.client_id,
            authority=authority,
            token_cache=self.cache,
        )

    def _persist_cache_if_changed(self) -> None:
        if self.cache.has_state_changed:
            Path(TOKEN_CACHE_PATH).write_text(self.cache.serialize())

    def acquire_token(self) -> str:
        scopes = self.cfg.scopes
        login_hint = self.cfg.login_hint

        # Try silent first (if any account is present)
        accounts = self.app.get_accounts(username=login_hint) if login_hint else self.app.get_accounts()
        if accounts:
            result = self.app.acquire_token_silent(scopes, account=accounts[0])
            if result and "access_token" in result:
                self._persist_cache_if_changed()
                return result["access_token"]

        # Fallback to interactive
        extra_args: Dict[str, Any] = {"login_hint": login_hint} if login_hint else {}
        result = self.app.acquire_token_interactive(scopes=scopes, **extra_args)
        if "access_token" not in result:
            raise RuntimeError(f"Token acquisition failed: {result.get('error_description') or result}")

        self._persist_cache_if_changed()
        return result["access_token"]


def main() -> int:
    load_dotenv()
    args = parse_args()
    needs_site_resolution = not args.site_id and bool(args.site_hostname or args.site_path)
    if (args.site_hostname and not args.site_path) or (args.site_path and not args.site_hostname):
        raise SystemExit("Provide both --site-hostname and --site-path when resolving a site.")

    scopes = NOTES_GRAPH_SCOPES.copy()
    if needs_site_resolution:
        scopes.extend(SITE_RESOLUTION_SCOPES)

    # Build config
    tenant_id = get_required_env("TENANT_ID")
    client_id = get_required_env("CLIENT_ID")

    cfg = AuthConfig(
        tenant_id=tenant_id,
        client_id=client_id,
        scopes=scopes,
        login_hint=args.login_hint,
        verbose=args.verbose,
    )

    try:
        access_token = MgAuth(cfg).acquire_token()
    except Exception as ex:
        _stderr(f"Auth error: {ex}")
        return 1

    effective_site_id = args.site_id

    if needs_site_resolution:
        resolved_site = resolve_site(access_token, args.site_hostname, args.site_path)
        effective_site_id = resolved_site["id"]
        print(f"Resolved site: {args.site_hostname}{resolved_site['resolvedPath']}")
        print(f"Resolved site name: {resolved_site.get('displayName', 'n/a')}")
        print(f"Resolved site URL: {resolved_site.get('webUrl', 'n/a')}")
        print(f"Site ID: {effective_site_id}")

    if effective_site_id:
        sections = get_site_sections(access_token, effective_site_id)
        filtered_sections = filter_sections_by_notebook(
            sections, args.notebook_id, args.notebook_name
        )

        if not needs_site_resolution:
            print(f"Site ID: {effective_site_id}")
        if args.notebook_id:
            print(f"Notebook ID filter: {args.notebook_id}")
        if args.notebook_name:
            print(f"Notebook name filter: {args.notebook_name}")

        return print_sections(filtered_sections, include_parent_notebook=True)

    notebook = resolve_notebook(
        access_token,
        args.notebook_id,
        args.notebook_name,
        effective_site_id,
    )
    if notebook is None:
        return 0

    sections = get_sections(access_token, notebook["id"])
    print(f"Notebook: {notebook['displayName']}")
    print(f"Notebook ID: {notebook['id']}")
    return print_sections(sections)


if __name__ == "__main__":
    sys.exit(main())
