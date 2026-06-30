/**
 * sharepoint.js
 * SharePoint data layer — resolves Drive/File IDs at runtime and reads
 * worksheet data via the Graph Workbook API (no binary download needed).
 *
 * Public API:
 *   fetchSheetData(sheetName?)  — returns { headers, rows }
 *   clearCache()                — force full re-resolve on next call
 */

import { CONFIG }   from './config.js';
import { graphGet } from './graph.js';

// Session-scoped cache — avoids redundant API calls within one page load
let _driveId    = null;
let _fileItemId = null;

// ── Step 1: Get the default Drive ID for the configured site ──────────────────

async function getDriveId() {
  if (_driveId) return _driveId;
  console.debug('[SharePoint] Fetching drive ID…');
  const data = await graphGet(`/sites/${CONFIG.sharepoint.siteId}/drive`);
  _driveId = data.id;
  console.debug('[SharePoint] Drive ID:', _driveId);
  return _driveId;
}

// ── Step 2: Resolve the Excel file's Drive Item ID from its path ──────────────

async function getFileItemId() {
  if (_fileItemId) return _fileItemId;
  const driveId = await getDriveId();

  // Encode path but preserve forward slashes as path separators
  const encoded = CONFIG.sharepoint.filePath
    .split('/')
    .map(segment => encodeURIComponent(segment))
    .join('/');

  console.debug('[SharePoint] Resolving file:', CONFIG.sharepoint.filePath);
  const data = await graphGet(`/drives/${driveId}/root:/${encoded}`);
  _fileItemId = data.id;
  console.debug('[SharePoint] File item ID:', _fileItemId);
  return _fileItemId;
}

// ── Step 3: Read usedRange from the worksheet ─────────────────────────────────
// Graph Workbook API returns all non-empty cells as a 2-D array — no parsing needed.

async function getUsedRange(itemId, sheetName) {
  const driveId = await getDriveId();
  const encoded = encodeURIComponent(sheetName);
  console.debug('[SharePoint] Reading sheet:', sheetName);
  const data = await graphGet(
    `/drives/${driveId}/items/${itemId}/workbook/worksheets/${encoded}/usedRange`
  );
  return data.values; // 2-D array: [ [header1, header2, ...], [val, val, ...], ... ]
}

// ── Step 4: Convert 2-D array → array of plain objects ───────────────────────

function valuesToRows(values) {
  if (!values || values.length < 2) {
    console.warn('[SharePoint] Sheet is empty or has only a header row.');
    return [];
  }

  const headers = values[0].map(h => String(h).trim());
  console.debug('[SharePoint] Headers:', headers);

  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const raw = values[i];
    // Skip fully blank rows
    if (!raw.some(cell => String(cell).trim() !== '')) continue;

    const obj = {};
    headers.forEach((h, idx) => {
      obj[h] = idx < raw.length ? String(raw[idx]).trim() : '';
    });
    rows.push(obj);
  }

  console.info(`[SharePoint] Parsed ${rows.length} data rows.`);
  return rows;
}

// ── Public API ────────────────────────────────────────────────────────────────

/**
 * Fetch all data rows from the configured worksheet.
 * @param {string} [sheetName] defaults to CONFIG.sharepoint.sheetName
 * @returns {Promise<{ headers: string[], rows: object[] }>}
 */
export async function fetchSheetData(sheetName = CONFIG.sharepoint.sheetName) {
  const itemId = await getFileItemId();
  const values = await getUsedRange(itemId, sheetName);

  const headers = values?.[0]?.map(h => String(h).trim()).filter(Boolean) ?? [];
  const rows    = valuesToRows(values);

  return { headers, rows };
}

/**
 * Clear cached IDs so the next fetchSheetData() does a full re-resolve.
 */
export function clearCache() {
  _driveId    = null;
  _fileItemId = null;
}
