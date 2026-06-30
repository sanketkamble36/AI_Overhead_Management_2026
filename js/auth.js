/**
 * auth.js
 * MSAL Browser authentication module.
 *
 * Exposes:
 *   login()           — popup sign-in
 *   logout()          — sign out
 *   getAccessToken()  — silent acquire, falls back to popup
 *   isAuthenticated() — true if a cached account exists
 *   getAccount()      — returns the active MSAL account object
 *   waitForInit()     — resolves after MSAL handles redirect/cache on page load
 *
 * Uses Authorization Code Flow with PKCE — no client secret required.
 */

import { CONFIG } from './config.js';

// ── MSAL instance ─────────────────────────────────────────────────────────────
// `msal` is the global exposed by the CDN script in index.html

const msalConfig = {
  auth: {
    clientId:    CONFIG.auth.clientId,
    authority:   `https://login.microsoftonline.com/${CONFIG.auth.tenantId}`,
    redirectUri: CONFIG.auth.redirectUri,
  },
  cache: {
    // sessionStorage clears on tab close — safer on shared/corporate machines
    cacheLocation: 'sessionStorage',
    storeAuthStateInCookie: false,
  },
  system: {
    loggerOptions: {
      loggerCallback(level, message, containsPii) {
        if (containsPii) return;
        switch (level) {
          case msal.LogLevel.Error:   console.error('[MSAL]', message); break;
          case msal.LogLevel.Warning: console.warn('[MSAL]', message);  break;
          case msal.LogLevel.Info:    console.info('[MSAL]', message);  break;
          case msal.LogLevel.Verbose: console.debug('[MSAL]', message); break;
        }
      },
      logLevel: msal.LogLevel.Warning,
    },
  },
};

const msalInstance = new msal.PublicClientApplication(msalConfig);

// handleRedirectPromise() must be called on every page load — required by MSAL
// even when using popup mode. Clears any pending redirect state from the URL.
const _initPromise = msalInstance.handleRedirectPromise().then(response => {
  if (response) {
    msalInstance.setActiveAccount(response.account);
    console.info('[Auth] Redirect response handled, account set.');
  }
}).catch(err => {
  console.error('[Auth] handleRedirectPromise error:', err);
});

// ── Public API ────────────────────────────────────────────────────────────────

/**
 * Sign the user in via a popup window.
 * @returns {Promise<AuthenticationResult|null>} null if user cancelled
 */
export async function login() {
  await _initPromise;
  try {
    const result = await msalInstance.loginPopup({
      scopes: CONFIG.scopes,
      prompt: 'select_account',
    });
    msalInstance.setActiveAccount(result.account);
    console.info('[Auth] Signed in as:', result.account.username);
    return result;
  } catch (err) {
    if (err.errorCode === 'user_cancelled') {
      console.info('[Auth] Login cancelled by user.');
      return null;
    }
    console.error('[Auth] Login error:', err);
    throw err;
  }
}

/**
 * Sign the user out and clear the MSAL cache.
 */
export async function logout() {
  await _initPromise;
  const account = msalInstance.getActiveAccount();
  await msalInstance.logoutPopup({
    account,
    postLogoutRedirectUri: CONFIG.auth.redirectUri,
  });
}

/**
 * Acquire an access token silently from cache or silent refresh.
 * Falls back to a popup if silent acquisition fails (expired, consent needed).
 * @returns {Promise<string>} Bearer access token
 */
export async function getAccessToken() {
  await _initPromise;

  const account = msalInstance.getActiveAccount();
  if (!account) throw new Error('No authenticated account. Call login() first.');

  const request = { scopes: CONFIG.scopes, account };

  try {
    const result = await msalInstance.acquireTokenSilent(request);
    console.debug('[Auth] Token acquired silently.');
    return result.accessToken;
  } catch (silentErr) {
    console.warn('[Auth] Silent token failed, falling back to popup:', silentErr.errorCode);
    if (silentErr instanceof msal.InteractionRequiredAuthError) {
      const result = await msalInstance.acquireTokenPopup(request);
      console.info('[Auth] Token acquired via popup.');
      return result.accessToken;
    }
    throw silentErr;
  }
}

/**
 * Returns true if there is a signed-in account in the MSAL cache.
 */
export function isAuthenticated() {
  return msalInstance.getAllAccounts().length > 0;
}

/**
 * Returns the active MSAL account, or null if not signed in.
 */
export function getAccount() {
  return msalInstance.getActiveAccount()
    || msalInstance.getAllAccounts()[0]
    || null;
}

/**
 * Wait for MSAL init and restore active account from cache on page refresh.
 */
export async function waitForInit() {
  await _initPromise;
  if (!msalInstance.getActiveAccount()) {
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length > 0) msalInstance.setActiveAccount(accounts[0]);
  }
}
