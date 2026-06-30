/**
 * dashboard.js
 * Renders the live Power Pairs data table.
 *
 * Features:
 *   - Dynamic columns (no fixed column names assumed)
 *   - Search/filter across all columns
 *   - Sortable headers (click to sort, click again to reverse)
 *   - Status and Led-By badges matching the existing page style
 *   - Sticky header, row count
 */

import { getSearchValue } from './ui.js';

// ── Module state ──────────────────────────────────────────────────────────────

let _headers = [];
let _allRows = [];
let _sortCol = null;
let _sortAsc = true;

// ── Badge renderers ───────────────────────────────────────────────────────────

const STATUS_MAP = {
  'ga':                  ['badge-ga',    'GA'],
  'generally available': ['badge-ga',    'GA'],
  'beta':                ['badge-beta',  'Beta'],
  'in development':      ['badge-dev',   'In Development'],
  'dev':                 ['badge-dev',   'In Development'],
  'prototyping':         ['badge-proto', 'Prototyping'],
  'proto':               ['badge-proto', 'Prototyping'],
  'on hold':             ['badge-hold',  'On Hold'],
  'hold':                ['badge-hold',  'On Hold'],
  'ideation':            ['badge-idea',  'Ideation'],
  'idea':                ['badge-idea',  'Ideation'],
  'na':                  ['badge-hold',  'N/A'],
};

function esc(s) {
  return String(s || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}

function statusBadge(raw) {
  const key = (raw || '').trim().toLowerCase();
  const [cls, label] = STATUS_MAP[key] || ['badge-idea', raw || '—'];
  return `<span class="badge ${cls}">${esc(label)}</span>`;
}

function ledBadge(raw) {
  const val = (raw || '').trim();
  const low = val.toLowerCase();
  if (['puc','public cloud','☁️ puc'].includes(low))
    return '<span class="badge badge-dev live-badge-sm">PuC</span>';
  if (['pce','private cloud','🔒 pce'].includes(low))
    return '<span class="badge badge-proto live-badge-sm">PCE</span>';
  return `<span class="live-led-other">${esc(val) || 'TBD'}</span>`;
}

function renderCell(header, value) {
  const h = header.toLowerCase();
  if (h === 'status') return statusBadge(value);
  if (h === 'led by') return ledBadge(value);
  return esc(value);
}

// ── Grouping ──────────────────────────────────────────────────────────────────

const GROUP_ORDER = ['puc', 'tbd', 'pce'];
const GROUP_LABELS = {
  'puc': '▸ PuC-Led Agents',
  'tbd': '▸ Still Alignment With PUC/PCE',
  'pce': '▸ PCE-Led Agents',
};

function groupRows(rows) {
  const groups = {};
  rows.forEach(row => {
    const key = (row['Led By'] || '').trim().toLowerCase();
    if (!groups[key]) groups[key] = [];
    groups[key].push(row);
  });

  // Build ordered output: known groups first, then any others
  const ordered = [];
  GROUP_ORDER.forEach(k => { if (groups[k]) ordered.push([k, groups[k]]); });
  Object.keys(groups).forEach(k => {
    if (!GROUP_ORDER.includes(k)) ordered.push([k, groups[k]]);
  });
  return ordered;
}

// ── Sort + filter ─────────────────────────────────────────────────────────────

function getVisibleRows() {
  const q = getSearchValue();
  let rows = q
    ? _allRows.filter(r => Object.values(r).some(v => String(v).toLowerCase().includes(q)))
    : [..._allRows];

  if (_sortCol !== null) {
    const key = _headers[_sortCol];
    rows.sort((a, b) => {
      const av = (a[key] || '').toLowerCase();
      const bv = (b[key] || '').toLowerCase();
      return _sortAsc ? av.localeCompare(bv) : bv.localeCompare(av);
    });
  }
  return rows;
}

// ── Render ────────────────────────────────────────────────────────────────────

export function render() {
  const tableEl = document.getElementById('live-table');
  if (!tableEl) return;

  const rows = getVisibleRows();

  // thead
  let thead = tableEl.tHead;
  if (!thead) { thead = tableEl.createTHead(); }
  thead.innerHTML = '';
  const tr = thead.insertRow();
  _headers.forEach((h, i) => {
    const th = document.createElement('th');
    th.textContent = h;
    th.classList.add('sortable');
    if (_sortCol === i) th.classList.add(_sortAsc ? 'sort-asc' : 'sort-desc');
    th.addEventListener('click', () => {
      _sortCol === i ? (_sortAsc = !_sortAsc) : (_sortCol = i, _sortAsc = true);
      render();
    });
    tr.appendChild(th);
  });

  // tbody
  let tbody = tableEl.tBodies[0];
  if (!tbody) tbody = tableEl.createTBody();
  tbody.innerHTML = '';

  if (rows.length === 0) {
    const empty = tbody.insertRow();
    empty.innerHTML = `<td colspan="${_headers.length}" class="live-empty">No matching rows found.</td>`;
  } else {
    // Only group when not actively sorting (sorting overrides grouping)
    const useGroups = _sortCol === null;
    if (useGroups) {
      groupRows(rows).forEach(([key, groupRows]) => {
        const label = GROUP_LABELS[key] || `▸ ${key.toUpperCase()}`;
        const headerRow = tbody.insertRow();
        headerRow.innerHTML = `<td colspan="${_headers.length}" style="background:#F3F8FF;font-weight:700;color:#1565C0;padding:6px 10px;font-size:12px;letter-spacing:.5px;">${label}</td>`;
        groupRows.forEach(row => {
          const r = tbody.insertRow();
          _headers.forEach(h => {
            const td = r.insertCell();
            td.innerHTML = renderCell(h, row[h] || '');
          });
        });
      });
    } else {
      rows.forEach(row => {
        const r = tbody.insertRow();
        _headers.forEach(h => {
          const td = r.insertCell();
          td.innerHTML = renderCell(h, row[h] || '');
        });
      });
    }
  }

  // Row count
  const counter = document.getElementById('live-row-count');
  if (counter) counter.textContent = `${rows.length} of ${_allRows.length} rows`;
}

// ── Public API ────────────────────────────────────────────────────────────────

/**
 * Load new data and render.
 */
export function loadData(headers, rows) {
  _headers = headers;
  _allRows = rows;
  _sortCol = null;
  _sortAsc = true;
  render();
}

/**
 * Wire the search input to live-filter the table.
 */
export function initSearch() {
  const input = document.getElementById('live-search');
  if (input) input.addEventListener('input', render);
}
