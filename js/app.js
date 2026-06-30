/**
 * app.js
 * Entry point — wires auth, data fetching, and UI together.
 *
 * Flow:
 *   Page load  → waitForInit() → if cached session: load data immediately
 *              → else: show Sign In button
 *   Sign In    → loginPopup → showSignedIn → loadDashboard
 *   Refresh    → clearCache → loadDashboard
 *   Sign Out   → logout → showSignedOut → reset UI
 */

import { login, logout, isAuthenticated, getAccount, waitForInit } from './auth.js';
import { fetchSheetData, clearCache }                               from './sharepoint.js';
import { loadData, initSearch }                                     from './dashboard.js';
import {
  showSpinner, hideSpinner,
  showError,   hideError,
  showSignedIn, showSignedOut, showDataSection,
  setLastUpdated,
} from './ui.js';

// ── Load data from SharePoint and render ──────────────────────────────────────

async function loadDashboard() {
  hideError();
  showSpinner('Connecting to SharePoint…');
  try {
    const { headers, rows } = await fetchSheetData();
    loadData(headers, rows);
    showDataSection();
    setLastUpdated();
    console.info('[App] Dashboard loaded —', rows.length, 'rows.');
  } catch (err) {
    console.error('[App] Load failed:', err);
    let msg = 'Failed to load data from SharePoint.';
    if (err.status === 401) msg = 'Your session expired. Please sign in again.';
    else if (err.status === 403) msg = 'Access denied. You may not have permission to read this file.';
    else if (err.status === 404) msg = 'Excel file not found in SharePoint. Check the file path in config.js.';
    else if (err.message)        msg = err.message;
    showError(msg);
  } finally {
    hideSpinner();
  }
}

// ── Button handlers ───────────────────────────────────────────────────────────

async function handleSignIn() {
  hideError();
  showSpinner('Opening sign-in…');
  try {
    const result = await login();
    if (!result) { hideSpinner(); return; } // user cancelled popup
    showSignedIn(result.account);
    await loadDashboard();
  } catch (err) {
    hideSpinner();
    showError(`Sign in failed: ${err.message || err}`);
  }
}

async function handleSignOut() {
  try { await logout(); } catch { /* non-critical */ }
  showSignedOut();
}

async function handleRefresh() {
  clearCache();
  await loadDashboard();
}

// ── Bootstrap ─────────────────────────────────────────────────────────────────

async function init() {
  await waitForInit();
  initSearch();

  document.getElementById('btn-sign-in') ?.addEventListener('click', handleSignIn);
  document.getElementById('btn-sign-out')?.addEventListener('click', handleSignOut);
  document.getElementById('btn-refresh') ?.addEventListener('click', handleRefresh);

  if (isAuthenticated()) {
    showSignedIn(getAccount());
    await loadDashboard();
  } else {
    showSignedOut();
  }
}

document.addEventListener('DOMContentLoaded', init);
