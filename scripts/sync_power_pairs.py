"""
Fetches the Power_pair tab from SharePoint Excel via Microsoft Graph API
and updates the Power Pairs <tbody> in index.html.

Auth: Client Credentials (app-only) — no user login needed.
Required env vars (set as GitHub Secrets):
  TENANT_ID, CLIENT_ID, CLIENT_SECRET
"""

import os
import re
import sys
import requests

# ── Config ────────────────────────────────────────────────────────────────────

TENANT_ID     = os.environ["TENANT_ID"]
CLIENT_ID     = os.environ["CLIENT_ID"]
CLIENT_SECRET = os.environ["CLIENT_SECRET"]

SITE_ID    = "990ca898-242d-45b0-89e8-07d655473f4d"
# Drive path relative to the site's default document library
FILE_PATH  = "NEW - AS OF 2025/10_Team/30_Team Folders/Management Accounting & Revenue Recognition/AI/All-in On AI Wrokshop/AI_UseCase_Agents and Assistants.xlsx"
SHEET_NAME = "Power_pair"

HTML_FILE  = os.path.join(os.path.dirname(__file__), "..", "index.html")

# ── Status → badge CSS class mapping ─────────────────────────────────────────

STATUS_BADGE = {
    "ga":             ("badge-ga",    "GA"),
    "generally available": ("badge-ga", "GA"),
    "beta":           ("badge-beta",  "Beta"),
    "in development": ("badge-dev",   "In Development"),
    "dev":            ("badge-dev",   "In Development"),
    "prototyping":    ("badge-proto", "Prototyping"),
    "proto":          ("badge-proto", "Prototyping"),
    "on hold":        ("badge-hold",  "On Hold"),
    "hold":           ("badge-hold",  "On Hold"),
    "ideation":       ("badge-idea",  "Ideation"),
    "idea":           ("badge-idea",  "Ideation"),
}

def status_badge(raw):
    key = (raw or "").strip().lower()
    cls, label = STATUS_BADGE.get(key, ("badge-idea", raw.strip() if raw else "—"))
    return f'<span class="badge {cls}">{label}</span>'

# ── Led-By badge ──────────────────────────────────────────────────────────────

def led_badge(raw):
    val = (raw or "").strip()
    val_low = val.lower()
    if val_low in ("puc", "public cloud", "☁️ puc"):
        return '<span class="badge badge-dev" style="font-size:11px;">PuC</span>'
    if val_low in ("pce", "private cloud", "🔒 pce"):
        return '<span class="badge badge-proto" style="font-size:11px;">PCE</span>'
    return f'<span style="font-size:11px;font-weight:600;color:#888;">{val or "TBD"}</span>'

def group_key(raw):
    val = (raw or "").strip().lower()
    if val in ("puc", "public cloud", "☁️ puc"):
        return "puc"
    if val in ("pce", "private cloud", "🔒 pce"):
        return "pce"
    return "tbd"

# ── HTML escaping ─────────────────────────────────────────────────────────────

def esc(s):
    return (s or "—").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

# ── Step 1: Get access token ──────────────────────────────────────────────────

def get_token():
    url  = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    data = {
        "grant_type":    "client_credentials",
        "client_id":     CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope":         "https://graph.microsoft.com/.default",
    }
    r = requests.post(url, data=data, timeout=30)
    r.raise_for_status()
    return r.json()["access_token"]

# ── Step 2: Get the Drive item ID for the Excel file ─────────────────────────

def get_file_drive_item(token):
    encoded = requests.utils.quote(FILE_PATH, safe="")
    url = f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/drive/root:/{encoded}"
    r = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=30)
    r.raise_for_status()
    return r.json()["id"]

# ── Step 3: Read rows from the Power_pair worksheet ───────────────────────────

def get_sheet_rows(token, item_id):
    # usedRange returns all non-empty cells as a 2-D array
    url = (
        f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/drive/items/{item_id}"
        f"/workbook/worksheets/{SHEET_NAME}/usedRange"
    )
    r = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=30)
    r.raise_for_status()
    values = r.json()["values"]
    if not values:
        print("ERROR: Sheet returned no data", file=sys.stderr)
        sys.exit(1)

    headers = [str(h).strip() for h in values[0]]
    print(f"Excel columns found: {headers}")

    # Map Excel column names → indices (case-insensitive)
    col = {h.lower(): i for i, h in enumerate(headers)}

    def ci(name):
        """Return column index by name (case-insensitive)."""
        k = name.lower()
        if k not in col:
            print(f"WARNING: column '{name}' not found in sheet. Available: {headers}", file=sys.stderr)
            return None
        return col[k]

    idx_agent   = ci("Agent")
    idx_led     = ci("Led By")
    idx_puc_pm  = ci("PUC PM")
    idx_puc_dm  = ci("PUC Dev Mgr")
    idx_pce_pm  = ci("PCE PM")
    idx_pce_dm  = ci("PCE Dev Mgr")
    idx_status  = ci("Status")

    rows = []
    for row in values[1:]:
        def cell(idx):
            if idx is None:
                return ""
            return str(row[idx]).strip() if idx < len(row) else ""

        agent = cell(idx_agent)
        if not agent:          # skip blank rows
            continue
        rows.append({
            "agent":   agent,
            "led_by":  cell(idx_led),
            "puc_pm":  cell(idx_puc_pm),
            "puc_dm":  cell(idx_puc_dm),
            "pce_pm":  cell(idx_pce_pm),
            "pce_dm":  cell(idx_pce_dm),
            "status":  cell(idx_status),
        })

    print(f"Rows fetched: {len(rows)}")
    return rows

# ── Step 4: Build replacement <tbody> HTML ────────────────────────────────────

GROUP_META = {
    "puc": ("▸ PuC-Led Agents",              "#1565C0"),
    "tbd": ("▸ Still Alignment With PUC/PCE", "#1565C0"),
    "pce": ("▸ PCE-Led Agents",               "#1565C0"),
}
GROUP_ORDER = ["puc", "tbd", "pce"]

def build_tbody(rows):
    # Bucket rows by group
    groups = {"puc": [], "tbd": [], "pce": []}
    for r in rows:
        groups[group_key(r["led_by"])].append(r)

    lines = []
    for gkey in GROUP_ORDER:
        group_rows = groups[gkey]
        if not group_rows:
            continue
        label, color = GROUP_META[gkey]
        lines.append(
            f'            <tr style="background:#F3F8FF;">'
            f'<td colspan="7" style="font-weight:700;color:{color};'
            f'padding:6px 10px;font-size:12px;letter-spacing:.5px;">{label}</td></tr>'
        )
        for r in group_rows:
            lines.append(
                f'            <tr>'
                f'<td>{esc(r["agent"])}</td>'
                f'<td style="text-align:center;">{led_badge(r["led_by"])}</td>'
                f'<td>{esc(r["pce_pm"])}</td>'
                f'<td>{esc(r["pce_dm"])}</td>'
                f'<td>{esc(r["puc_pm"])}</td>'
                f'<td>{esc(r["puc_dm"])}</td>'
                f'<td style="text-align:center;">{status_badge(r["status"])}</td>'
                f'</tr>'
            )

    return "\n".join(lines)

# ── Step 5: Inject into index.html ───────────────────────────────────────────

# Markers — the script replaces everything between these two comments
START_MARKER = "<!-- POWER-PAIRS-START -->"
END_MARKER   = "<!-- POWER-PAIRS-END -->"

def inject_html(tbody_html):
    with open(HTML_FILE, "r", encoding="utf-8") as f:
        html = f.read()

    if START_MARKER not in html or END_MARKER not in html:
        print(
            f"ERROR: Markers not found in index.html.\n"
            f"Add these comments around the Power Pairs <tbody>:\n"
            f"  {START_MARKER}\n"
            f"  {END_MARKER}",
            file=sys.stderr,
        )
        sys.exit(1)

    pattern = re.compile(
        re.escape(START_MARKER) + r".*?" + re.escape(END_MARKER),
        re.DOTALL,
    )
    replacement = f"{START_MARKER}\n{tbody_html}\n          {END_MARKER}"
    new_html, n = pattern.subn(replacement, html)

    if n != 1:
        print("ERROR: Could not replace markers (found 0 or multiple matches)", file=sys.stderr)
        sys.exit(1)

    with open(HTML_FILE, "w", encoding="utf-8") as f:
        f.write(new_html)

    print("index.html updated successfully.")

# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    print("Authenticating...")
    token = get_token()

    print("Locating Excel file in SharePoint...")
    item_id = get_file_drive_item(token)
    print(f"File item ID: {item_id}")

    print(f"Reading sheet: {SHEET_NAME}")
    rows = get_sheet_rows(token, item_id)

    tbody = build_tbody(rows)
    inject_html(tbody)

if __name__ == "__main__":
    main()
