/**
 * config.js
 * Central configuration for the AI Overhead Management Dashboard.
 *
 * SECURITY NOTE:
 * Only Client ID and Tenant ID belong here — both are public-facing identifiers
 * that are safe to expose in browser code. They are visible in every OAuth 2.0
 * authorization request by design.
 *
 * The Client Secret MUST NEVER appear here. A client secret in browser code
 * is not a secret — anyone can read it from DevTools. The secret is only used
 * by server-side processes (e.g. GitHub Actions) where it is stored as an
 * encrypted repository secret.
 *
 * This app uses Delegated Authentication (MSAL Browser popup) which does
 * not require a client secret — the signed-in user's identity is the credential.
 */

export const CONFIG = {

  // ── Azure AD / MSAL ─────────────────────────────────────────────────────────
  auth: {
    clientId:    '7a405a92-f548-4f7e-87dc-b2cf3c4a6bd4',
    tenantId:    '42f7676c-f455-423c-82f6-dc2d99791af7',
    // Must exactly match a Redirect URI registered in Azure App Registration
    // as platform type "Single-page application (SPA)"
    redirectUri: 'https://sanketkamble36.github.io/AI_Overhead_Management_2026/',
  },

  // ── Microsoft Graph scopes (delegated) ──────────────────────────────────────
  // Request only the minimum scopes required. Users see a consent screen
  // listing exactly these permissions.
  scopes: [
    'User.Read',       // Read the signed-in user's profile (name, email)
    'Files.Read',      // Read files the user has access to
    'Files.Read.All',  // Read all files the user can access
    'Sites.Read.All',  // Read items in all site collections
  ],

  // ── SharePoint / Excel ──────────────────────────────────────────────────────
  sharepoint: {
    // Known Site ID — stable, saves one extra Graph API call at startup
    siteId:   '789d099d-bb6c-4b54-8ae9-ff5d5d139f0f',

    // Path to the Excel file relative to the site's default document library root
    filePath: 'NEW - AS OF 2025/10_Team/30_Team Folders/Management Accounting & Revenue Recognition/AI/All-in On AI Wrokshop/AI_UseCase_Agents and Assistants.xlsx',

    // Worksheet name to read
    sheetName: 'Power_pair',
  },

};
