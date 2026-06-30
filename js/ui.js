/**
 * ui.js
 * Low-level UI helpers — show/hide elements, update labels.
 * No business logic. Called by app.js and dashboard.js.
 */

const el = id => document.getElementById(id);

// ── Spinner ───────────────────────────────────────────────────────────────────

export function showSpinner(message = 'Loading…') {
  const spinner = el('live-spinner');
  const text    = el('live-spinner-text');
  if (spinner) spinner.style.display = 'flex';
  if (text)    text.textContent = message;
}

export function hideSpinner() {
  const spinner = el('live-spinner');
  if (spinner) spinner.style.display = 'none';
}

// ── Error banner ──────────────────────────────────────────────────────────────

export function showError(message) {
  const banner = el('live-error-banner');
  const text   = el('live-error-text');
  if (text)   text.textContent = message;
  if (banner) banner.style.display = 'flex';
}

export function hideError() {
  const banner = el('live-error-banner');
  if (banner) banner.style.display = 'none';
}

// ── Auth UI ───────────────────────────────────────────────────────────────────

export function showSignedIn(account) {
  const out    = el('auth-signed-out');
  const inn    = el('auth-signed-in');
  const name   = el('auth-user-name');
  const status = el('connection-status');

  if (out)    out.style.display  = 'none';
  if (inn)    inn.style.display  = 'flex';
  if (name)   name.textContent   = account.name || account.username || 'User';
  if (status) {
    status.textContent = 'Connected';
    status.className   = 'connection-status connected';
  }
}

export function showSignedOut() {
  const out     = el('auth-signed-out');
  const inn     = el('auth-signed-in');
  const status  = el('connection-status');
  const section = el('live-data-section');

  if (out)     out.style.display    = 'flex';
  if (inn)     inn.style.display    = 'none';
  if (section) section.style.display = 'none';
  if (status) {
    status.textContent = 'Not connected';
    status.className   = 'connection-status disconnected';
  }
}

export function showDataSection() {
  const section = el('live-data-section');
  if (section) section.style.display = 'block';
}

// ── Timestamp ─────────────────────────────────────────────────────────────────

export function setLastUpdated(date = new Date()) {
  const stamp = el('last-updated');
  if (stamp) stamp.textContent = `Last updated: ${date.toLocaleString()}`;
}

// ── Search value ──────────────────────────────────────────────────────────────

export function getSearchValue() {
  const input = el('live-search');
  return input ? input.value.trim().toLowerCase() : '';
}
