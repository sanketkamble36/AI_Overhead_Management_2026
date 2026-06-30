# AI Overhead Management Dashboard

A production-quality GitHub Pages application that displays live data from an Excel workbook stored in SharePoint Online.

No backend server. No build step. Pure ES Modules running directly in the browser.

---

## Architecture

```
/index.html          Main page — all existing content + live Power Pairs section
/css/styles.css      Styles for the live data section only
/js/config.js        Central configuration (Client ID, Tenant ID, SharePoint details)
/js/auth.js          MSAL Browser authentication module
/js/graph.js         Microsoft Graph API wrapper
/js/sharepoint.js    SharePoint data layer (drive/file resolution + worksheet read)
/js/dashboard.js     Table rendering (sort, search, badges)
/js/ui.js            DOM helpers (spinner, error banner, auth UI)
/js/app.js           Entry point — wires everything together
```

---

## Authentication Flow

1. Page loads → MSAL initialises, checks `sessionStorage` for a cached account
2. If cached → silently acquires a fresh token → loads data immediately
3. If not cached → user clicks **Sign in with Microsoft**
4. MSAL opens a popup → user signs in with their SAP Microsoft account
5. On success → token stored in `sessionStorage` → data is fetched
6. On token expiry → MSAL silently refreshes; if interaction is required, a popup is shown automatically

> **Why no Client Secret in the browser?**
> A secret in browser code is visible to anyone via DevTools — it is not a secret.
> This app uses **Delegated Authentication** (OAuth 2.0 Authorization Code + PKCE).
> The signed-in user's identity is the credential. No secret is needed or used.

---

## SharePoint / Graph Flow

1. `sharepoint.js` calls `GET /sites/{siteId}/drive` → resolves the Drive ID
2. Calls `GET /drives/{driveId}/root:/{filePath}` → resolves the Excel file's Item ID
3. Calls `GET /drives/{driveId}/items/{itemId}/workbook/worksheets/Power_pair/usedRange`
4. Graph returns a 2-D array of cell values — no binary download, no SheetJS needed
5. `valuesToRows()` converts the array into plain JS objects keyed by column headers

Drive ID and File Item ID are cached in module scope for the session lifetime.
Click **Refresh** to re-resolve and re-fetch.

---

## Configuration

All values live in `js/config.js`:

| Key | Value |
|---|---|
| `auth.clientId` | Azure App Registration Application (client) ID |
| `auth.tenantId` | Azure Active Directory Directory (tenant) ID |
| `auth.redirectUri` | Must exactly match the SPA redirect URI in Azure |
| `sharepoint.siteId` | SharePoint site ID |
| `sharepoint.filePath` | Path to the Excel file relative to the document library root |
| `sharepoint.sheetName` | Worksheet name to read |

---

## Azure App Registration Requirements

| Setting | Required value |
|---|---|
| Platform type | Single-page application (SPA) |
| Redirect URI | `https://sanketkamble36.github.io/AI_Overhead_Management_2026/` |
| Supported account types | Accounts in this organizational directory only |
| `User.Read` | Delegated |
| `Files.Read` | Delegated |
| `Files.Read.All` | Delegated |
| `Sites.Read.All` | Delegated |
| Client Secret | Not used by this app (server-side only) |

---

## Deployment

1. Push to the `main` branch of `sanketkamble36/AI_Overhead_Management_2026`
2. GitHub Pages serves from `main` / root automatically
3. Site is live at: `https://sanketkamble36.github.io/AI_Overhead_Management_2026/`

---

## Troubleshooting

| Symptom | Cause | Fix |
|---|---|---|
| Popup blocked | Browser blocked the login popup | Allow popups for the site |
| `redirect_uri_mismatch` | Redirect URI in Azure doesn't match | Update Azure → Authentication → SPA redirect URI |
| `InteractionRequiredAuthError` | Consent not yet granted | Sign in again — consent screen will appear |
| 403 Forbidden | Missing permission or no access to file | Confirm delegated `Sites.Read.All` in Azure; confirm user has SharePoint access |
| 404 Not Found | Wrong file path | Check `sharepoint.filePath` in `config.js` |
| Blank table | Sheet name wrong or sheet is empty | Check `sharepoint.sheetName` in `config.js` |
| CORS error | MSAL CDN version mismatch | Ensure `msal-browser@3` is loaded in index.html |
