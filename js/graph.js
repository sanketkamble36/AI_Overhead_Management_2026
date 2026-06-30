/**
 * graph.js
 * Reusable Microsoft Graph API wrapper.
 *
 * All requests go through graphRequest() which:
 *   - Attaches the Bearer token automatically
 *   - Retries once on 429 (rate limiting) after Retry-After delay
 *   - Throws descriptive GraphError for all failure responses
 */

import { getAccessToken } from './auth.js';

const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';

// ── Custom error class ────────────────────────────────────────────────────────

export class GraphError extends Error {
  constructor(status, message, body) {
    super(message);
    this.name   = 'GraphError';
    this.status = status;
    this.body   = body;
  }
}

// ── Core request ──────────────────────────────────────────────────────────────

/**
 * Make an authenticated Graph API request.
 * @param {string} path     - e.g. '/me' or '/sites/{id}/drive'
 * @param {object} options  - fetch options (method, headers, body)
 * @param {number} _retry   - internal retry counter, do not pass externally
 */
export async function graphRequest(path, options = {}, _retry = 0) {
  const token = await getAccessToken();
  const url   = path.startsWith('http') ? path : `${GRAPH_BASE}${path}`;

  const response = await fetch(url, {
    ...options,
    headers: {
      'Authorization': `Bearer ${token}`,
      'Accept':        'application/json',
      'Content-Type':  'application/json',
      ...(options.headers || {}),
    },
  });

  // Rate limited — wait and retry once
  if (response.status === 429 && _retry === 0) {
    const wait = parseInt(response.headers.get('Retry-After') || '5', 10);
    console.warn(`[Graph] Rate limited. Retrying after ${wait}s…`);
    await new Promise(r => setTimeout(r, wait * 1000));
    return graphRequest(path, options, 1);
  }

  if (response.status === 204) return null;

  let body;
  try { body = await response.json(); } catch { body = null; }

  if (!response.ok) {
    const msg  = body?.error?.message || response.statusText || 'Unknown error';
    const code = body?.error?.code    || String(response.status);
    console.error(`[Graph] ${response.status} (${code}): ${msg}`);
    throw new GraphError(response.status, `[${code}] ${msg}`, body);
  }

  return body;
}

/**
 * GET shorthand — most common Graph operation.
 */
export async function graphGet(path) {
  return graphRequest(path, { method: 'GET' });
}
