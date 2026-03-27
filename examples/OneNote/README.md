# OneNote Examples

This directory contains a Microsoft Graph example for querying OneNote notebook sections.

## Files

- `OneNote-Query-Sections.py`
  Query sections from a personal notebook or a SharePoint-hosted team notebook.
- `.onenote-token-cache.json`
  Local MSAL token cache created after interactive sign-in.

## What The Script Supports

- Query sections from a personal OneNote notebook through `/me/onenote/notebooks/{id}/sections`
- Query sections from a SharePoint team site through `/sites/{site-id}/onenote/sections`
- Resolve a SharePoint site ID from hostname and server-relative path through `/sites/{hostname}:/{path}`
- Filter returned site sections by notebook ID or notebook name

## Authentication Model

The script uses MSAL public-client authentication with a local token cache.

Current behavior:

- Attempts silent token acquisition from the local cache first
- Falls back to interactive browser sign-in if needed
- Stores tokens in `examples/OneNote/.onenote-token-cache.json`

## Required Graph Permissions

The current script requests these delegated Microsoft Graph permissions:

- `Notes.Read.All`
- `Sites.Read.All` when resolving a SharePoint site from hostname and path

Your app registration also needs:

- public client flow enabled for interactive sign-in

## Python Dependencies

Install the dependencies used by the script:

```bash
pip install msal requests python-dotenv
```

If you are using the repo virtual environment on macOS:

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install msal requests python-dotenv
```

## Environment Variables

The script reads values from `.env` via `python-dotenv`.

Required:

- `TENANT_ID`
- `CLIENT_ID`

Optional:

- `LOGIN_HINT`
- `ONENOTE_NOTEBOOK_ID`
- `ONENOTE_SITE_ID`
- `ONENOTE_SITE_HOSTNAME`
- `ONENOTE_SITE_PATH`
- `DEFAULT_HOST`
- `SITE_PATH`

The repo template already includes the OneNote-specific variables in [example.env](/Users/jb/Source/wg_gh/M365-dev/example.env).

## Usage

Run from the repo root:

```bash
python examples/OneNote/OneNote-Query-Sections.py --help
```

### Personal Notebook

List available notebooks:

```bash
python examples/OneNote/OneNote-Query-Sections.py
```

Query a notebook by name:

```bash
python examples/OneNote/OneNote-Query-Sections.py --notebook-name "Team Notes"
```

Query a notebook by ID:

```bash
python examples/OneNote/OneNote-Query-Sections.py --notebook-id "<notebook-id>"
```

### SharePoint Team Notebook By Site ID

```bash
python examples/OneNote/OneNote-Query-Sections.py --site-id "<site-id>"
```

Filter site sections to a notebook name:

```bash
python examples/OneNote/OneNote-Query-Sections.py --site-id "<site-id>" --notebook-name "General"
```

### Resolve SharePoint Site By Hostname And Path

```bash
python examples/OneNote/OneNote-Query-Sections.py \
  --site-hostname contoso.sharepoint.com \
  --site-path /sites/TeamA
```

You can also combine that with notebook filters:

```bash
python examples/OneNote/OneNote-Query-Sections.py \
  --site-hostname contoso.sharepoint.com \
  --site-path /sites/TeamA \
  --notebook-name "General"
```

### Optional Login Hint

If your tenant has multiple cached accounts, you can steer sign-in with:

```bash
python examples/OneNote/OneNote-Query-Sections.py --login-hint user@contoso.com
```

## Example Output

For site resolution, the script prints the resolved SharePoint site details first:

```text
Resolved site: contoso.sharepoint.com/sites/TeamA
Resolved site name: TeamA
Resolved site URL: https://contoso.sharepoint.com/sites/TeamA
Site ID: contoso.sharepoint.com,site-collection-guid,site-guid
Sections:
- General | notebook=General | id=... | created=... | modified=...
```

For personal notebook queries, the script prints the notebook details and then the section list.

## Notes

- The local token cache file is machine-specific and should not be committed.
- Site-level section queries can return sections from multiple notebooks, so `--notebook-id` or `--notebook-name` is useful when the site contains more than one notebook.
- If a notebook name is ambiguous, use the notebook ID instead.