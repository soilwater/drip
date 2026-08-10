/**
 * IrrigationSpreadsheet.js
 * A standalone, zero-dependency spreadsheet library for irrigation applications.
 * Features: Hairline paper grid, sticky freeze headers, NEXRAD/NWS/MODIS/ColorBrewer colormaps,
 * threshold alerts, sparkline trends, read-only locking, and Excel-compatible paste.
 *
 * ── Fixes applied in this copy (see chat for full detail) ──
 * 1. FATAL: removed the trailing `export default IrrigationSpreadsheet;`. That statement is only
 *    legal inside an ES module. Loaded via a plain `<script src="...">` tag — this library's
 *    primary, intended environment — it throws `SyntaxError: Unexpected token 'export'` and the
 *    whole file fails to parse, so `window.IrrigationSpreadsheet` is never defined. Verified with
 *    `vm.Script` (same non-module grammar a classic script tag uses). Replaced with a
 *    runtime-guarded CommonJS export (see bottom of file) — `export` can't be guarded since it's
 *    parse-time-only, but `module.exports` is just a normal property write and is safe to leave
 *    in a browser context where `module` doesn't exist.
 * 2. onCellChange fired on the `input` event — once per keystroke, the same granularity as a
 *    native text input firing on every character. The docs' own Example 1 calls
 *    `sheet.setData(...)` directly inside `onCellChange`, which would replace the whole `<tbody>`
 *    mid-keystroke and kick the user out of the cell after the first character typed. Cell edits
 *    now commit once — on Enter (which no longer inserts a line break) or on blur — and
 *    `onCellChange` only fires from that single commit path. The per-keystroke `input` listener
 *    (which also recomputed conditional formatting, sparklines, and summary row on every
 *    character) is removed entirely; all of that now runs once, at commit, too.
 * 3. Removed the built-in toast widget. A grid module has no business owning notification UI —
 *    the host app should have one place to surface transient messages consistently, not one
 *    per widget. `copyColumnToClipboard`'s "copied!" message (and anything else that would have
 *    used `showToast`) now goes through an optional `onNotify(message)` constructor option;
 *    without one, it just logs to console instead of failing silently.
 * 4. showToast() (before its removal above) used `document.getElementById('irrigation-toast')`,
 *    unscoped to the instance's own container — harmless with one instance on a page, but it
 *    would have broken the moment a second instance was mounted.
 * 5. Default `sortState` was `{ colIndex: 0, asc: true }`, so a ▲ appeared on the first column
 *    before any sort had actually happened — misleading, since setData() does not sort. Changed
 *    to `{ colIndex: -1, asc: true }`, with the header render guarded so no indicator shows until
 *    a header is actually clicked.
 * 6. Removed the per-column copy-to-clipboard button. It shared header space with the sort-click
 *    target and the drag-resizer, and in practice was easy to hit by accident while trying to
 *    sort or resize a column instead. `copyColumnToClipboard(colIndex)` still exists as a public
 *    method for anyone who wants to wire it up elsewhere (e.g. a right-click menu); it's just no
 *    longer rendered as a header button. CSV export already covers "get this data into Excel".
 */

// ==========================================================================
// 1. UNIFIED COLORMAPS REGISTRY
// ==========================================================================
const COLORMAPS = {
  nexrad_base_reflectivity: [
    { "min": -30, "max": 4,   "hex": "#000000" },
    { "min": 5,   "max": 9,   "hex": "#404040" },
    { "min": 10,  "max": 14,  "hex": "#9C9C9C" },
    { "min": 15,  "max": 19,  "hex": "#00ECEC" },
    { "min": 20,  "max": 24,  "hex": "#01A0F6" },
    { "min": 25,  "max": 29,  "hex": "#0000F6" },
    { "min": 30,  "max": 34,  "hex": "#00FF00" },
    { "min": 35,  "max": 39,  "hex": "#00C800" },
    { "min": 40,  "max": 44,  "hex": "#009000" },
    { "min": 45,  "max": 49,  "hex": "#FFFF00" },
    { "min": 50,  "max": 54,  "hex": "#E7C000" },
    { "min": 55,  "max": 59,  "hex": "#FF9000" },
    { "min": 60,  "max": 64,  "hex": "#FF0000" },
    { "min": 65,  "max": 69,  "hex": "#D00000" },
    { "min": 70,  "max": 74,  "hex": "#A00000" },
    { "min": 75,  "max": 79,  "hex": "#FF00FF" },
    { "min": 80,  "max": 100, "hex": "#FFFFFF" }
  ],

  nws_precipitation_accumulation: [
    { "min": 0.00, "max": 0.00, "hex": "#00000000" },
    { "min": 0.01, "max": 0.09, "hex": "#7FFF7F" },
    { "min": 0.10, "max": 0.24, "hex": "#00C800" },
    { "min": 0.25, "max": 0.49, "hex": "#009000" },
    { "min": 0.50, "max": 0.74, "hex": "#01A0F6" },
    { "min": 0.75, "max": 0.99, "hex": "#0000F6" },
    { "min": 1.00, "max": 1.49, "hex": "#000090" },
    { "min": 1.50, "max": 1.99, "hex": "#FFFF00" },
    { "min": 2.00, "max": 2.99, "hex": "#E7C000" },
    { "min": 3.00, "max": 3.99, "hex": "#FF9000" },
    { "min": 4.00, "max": 4.99, "hex": "#FF0000" },
    { "min": 5.00, "max": 6.99, "hex": "#D00000" },
    { "min": 7.00, "max": 9.99, "hex": "#A00000" },
    { "min": 10.0, "max": 14.9, "hex": "#FF00FF" },
    { "min": 15.0, "max": 19.9, "hex": "#B000B0" },
    { "min": 20.0, "max": 99.0, "hex": "#FFFFFF" }
  ],

  plant_growth_modis: [
    { "hex": "#CE7E45" }, { "hex": "#DF923D" }, { "hex": "#F1B555" }, { "hex": "#FCD163" }, 
    { "hex": "#99B718" }, { "hex": "#74A901" }, { "hex": "#66A000" }, { "hex": "#529400" }, 
    { "hex": "#3E8601" }, { "hex": "#207401" }, { "hex": "#0A6B05" }, { "hex": "#005C00" }, 
    { "hex": "#004D00" }, { "hex": "#004200" }, { "hex": "#003B00" }, { "hex": "#003300" }
  ],

  water_deficit_ylorrd: [
    { "hex": "#ffffcc" }, { "hex": "#ffeda0" }, { "hex": "#fed976" }, { "hex": "#feb24c" }, 
    { "hex": "#fd8d3c" }, { "hex": "#fc4e2a" }, { "hex": "#e31a1c" }, { "hex": "#bd0026" }, { "hex": "#800026" }
  ],

  irrigation_blues: [
    { "hex": "#f7fbff" }, { "hex": "#deebf7" }, { "hex": "#c6dbef" }, { "hex": "#9ecae1" }, 
    { "hex": "#6baed6" }, { "hex": "#4292c6" }, { "hex": "#2171b5" }, { "hex": "#08519c" }, { "hex": "#08306b" }
  ]
};

// ==========================================================================
// 2. HELPER FUNCTIONS
// ==========================================================================
function getContrastTextColor(hexColor) {
  if (!hexColor || hexColor === 'transparent') return '#0f172a';
  let hex = hexColor.replace('#', '');
  if (hex.length === 8 && hex.substring(6, 8) === '00') return '#0f172a';
  if (hex.length >= 6) hex = hex.substring(0, 6);

  const r = parseInt(hex.substring(0, 2), 16);
  const g = parseInt(hex.substring(2, 4), 16);
  const b = parseInt(hex.substring(4, 6), 16);
  const yiq = ((r * 299) + (g * 587) + (b * 114)) / 1000;
  return (yiq >= 140) ? '#0f172a' : '#ffffff';
}

// ==========================================================================
// 3. AUTO-INJECTED CSS STYLES
// ==========================================================================
const SPREADSHEET_CSS = `
  :root {
    --bg-module: #ffffff;
    --bg-header-paper: #f8fafc;
    --bg-th: #f1f5f9;
    --bg-tfoot: #f1f5f9;
    --bg-readonly: #f8fafc;
    --border-dark: #0f172a;
    --border-grid: #e2e8f0;
    --text-main: #0f172a;
    --text-muted: #64748b;
    --accent-color: #0284c7;
    --alert-red: #dc2626;

    --status-optimal-bg: #dcfce7;
    --status-optimal-text: #14532d;
    --status-warning-bg: #fef9c3;
    --status-warning-text: #713f12;
    --status-critical-bg: #fee2e2;
    --status-critical-text: #7f1d1d;
  }

  .irrigation-module * { 
    box-sizing: border-box; 
    font-family: ui-sans-serif, -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
    border-radius: 0 !important; 
  }

  .irrigation-module {
    width: 100%;
    background: var(--bg-module);
    border: 1.5px solid var(--border-dark);
    box-shadow: 4px 4px 0px rgba(15, 23, 42, 0.08);
    padding: 24px;
  }

  .irrigation-module .field-header {
    text-align: left;
    margin-bottom: 16px;
    border-bottom: 2px solid var(--border-dark);
    padding-bottom: 12px;
  }

  .irrigation-module .title-row-container {
    display: flex;
    justify-content: space-between;
    align-items: center;
    margin-bottom: 8px;
  }

  .irrigation-module .field-title {
    font-size: 1.25rem;
    font-weight: 800;
    color: var(--text-main);
    margin: 0;
    text-transform: uppercase;
    letter-spacing: 0.05em;
  }

  .irrigation-module .field-status-pill {
    font-size: 0.75rem;
    font-weight: 800;
    padding: 4px 10px;
    text-transform: uppercase;
    letter-spacing: 0.05em;
    border: 1px solid var(--border-dark);
  }

  .irrigation-module .field-status-pill.status-optimal { background: var(--status-optimal-bg); color: var(--status-optimal-text); }
  .irrigation-module .field-status-pill.status-warning { background: var(--status-warning-bg); color: var(--status-warning-text); }
  .irrigation-module .field-status-pill.status-critical { background: var(--status-critical-bg); color: var(--status-critical-text); }

  .irrigation-module .metadata-subheader {
    display: flex;
    flex-wrap: wrap;
    align-items: center;
    gap: 16px;
    font-size: 0.8rem;
    color: var(--text-muted);
    background: var(--bg-header-paper);
    padding: 8px 12px;
    border: 1px solid var(--border-grid);
  }

  .irrigation-module .meta-tag {
    display: inline-flex;
    align-items: center;
    gap: 5px;
  }

  .irrigation-module .meta-tag span.label {
    font-weight: 700;
    color: var(--border-dark);
    text-transform: uppercase;
    font-size: 0.7rem;
    letter-spacing: 0.03em;
  }

  .irrigation-module .meta-divider { color: #cbd5e1; }

  .irrigation-module .spreadsheet-toolbar {
    display: flex;
    justify-content: space-between;
    align-items: center;
    gap: 10px;
    margin-bottom: 10px;
    background: #fafaf9;
    padding: 6px 12px;
    border: 1px solid var(--border-grid);
    font-size: 0.75rem;
    font-weight: 600;
    color: var(--text-muted);
  }

  .irrigation-module .table-container {
    overflow: auto;
    max-height: 440px;
    border: 1.5px solid var(--border-dark);
    background: #ffffff;
    position: relative;
  }

  .irrigation-module .irrigation-table {
    width: 100%;
    border-collapse: separate;
    border-spacing: 0;
    text-align: left;
    font-size: 0.825rem;
    table-layout: fixed;
  }

  .irrigation-module .irrigation-table th, 
  .irrigation-module .irrigation-table td {
    border-right: 1px solid var(--border-grid);
    border-bottom: 1px solid var(--border-grid);
    padding: 6px 10px;
    min-width: 110px;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
  }

  .irrigation-module .irrigation-table thead th {
    position: sticky;
    top: 0;
    z-index: 20;
    background: var(--bg-th);
    color: var(--text-main);
    font-weight: 700;
    font-size: 0.75rem;
    text-transform: uppercase;
    letter-spacing: 0.04em;
    border-bottom: 2px solid var(--border-dark) !important;
    cursor: pointer;
  }

  .irrigation-module .irrigation-table thead th:hover { background: #e2e8f0; }

  .irrigation-module .th-content {
    display: flex;
    flex-direction: column;
    gap: 2px;
  }

  .irrigation-module .th-top-row {
    display: flex;
    justify-content: space-between;
    align-items: center;
    min-height: 16px;
  }

  .irrigation-module .th-title-row {
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 4px;
    min-width: 0; /* required for the ellipsis rule below to actually take effect in a flex child */
  }

  .irrigation-module .th-title-row .th-label {
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
    min-width: 0;
  }

  .irrigation-module .sparkline-box {
    width: 50px;
    height: 14px;
  }

  .irrigation-module .header-alert-badge {
    background: var(--alert-red);
    color: #ffffff;
    font-size: 0.65rem;
    font-weight: 800;
    padding: 1px 4px;
    letter-spacing: 0.02em;
  }

  .irrigation-module .irrigation-table tfoot td {
    position: sticky;
    bottom: 0;
    z-index: 20;
    background: var(--bg-tfoot);
    color: var(--text-main);
    font-weight: 700;
    font-size: 0.8rem;
    border-top: 2px solid var(--border-dark) !important;
  }

  .irrigation-module .irrigation-table th:first-child,
  .irrigation-module .irrigation-table td:first-child {
    position: sticky;
    left: 0;
    z-index: 15;
    background: #ffffff;
    border-right: 2px solid var(--border-dark) !important;
  }

  .irrigation-module .irrigation-table thead th:first-child { z-index: 30; background: var(--bg-th); }
  .irrigation-module .irrigation-table tfoot td:first-child { z-index: 30; background: var(--bg-tfoot); }

  .irrigation-module .irrigation-table td.read-only-cell {
    background-color: var(--bg-readonly);
    color: var(--text-muted);
    cursor: default;
  }

  .irrigation-module .irrigation-table td:focus {
    box-shadow: inset 0 0 0 2px var(--border-dark);
    z-index: 10;
  }

  .irrigation-module .col-resizer {
    position: absolute;
    right: 0;
    top: 0;
    bottom: 0;
    width: 5px;
    cursor: col-resize;
    user-select: none;
    z-index: 25;
  }

  .irrigation-module .col-resizer:hover, 
  .irrigation-module .col-resizer.resizing {
    background-color: var(--border-dark);
  }

  .irrigation-module .sort-indicator {
    font-size: 0.7rem;
    color: var(--accent-color);
  }

  .irrigation-module .hint {
    font-size: 0.75rem;
    color: var(--text-muted);
    margin-top: 10px;
  }
`;

function injectStyles() {
  if (typeof document !== 'undefined' && !document.getElementById('irrigation-sheet-styles')) {
    const styleTag = document.createElement('style');
    styleTag.id = 'irrigation-sheet-styles';
    styleTag.textContent = SPREADSHEET_CSS;
    document.head.appendChild(styleTag);
  }
}

// ==========================================================================
// 4. MAIN CLASS IMPLEMENTATION
// ==========================================================================
class IrrigationSpreadsheet {
  constructor(containerId, options = {}) {
    injectStyles(); // Auto-inject styles into document head

    this.container = document.getElementById(containerId);
    if (!this.container) return;

    this.metadata = Object.assign({
      fieldName: "North Field - Sector A",
      soilType: "Silty Clay Loam",
      crop: "Corn (Maize)",
      plantingDate: "2026-04-15",
      latitude: "40.7128",
      longitude: "-74.0060",
      statusLabel: "STATUS: ADEQUATE MOISTURE",
      statusLevel: "optimal"
    }, options.metadata || {});

    this.columns = options.columns || ["Date", "DAP", "FAW", "NDVI", "Net Bal (mm)", "Precip Accum (in)", "Irrigation (mm)", "Notes"];

    this.sparklineColumns = options.sparklineColumns || ["FAW", "NDVI", "Net Bal (mm)", "Precip Accum (in)", "Irrigation (mm)"];
    this.readOnlyColumns = options.readOnlyColumns || ["Date", "DAP", "Net Bal (mm)", "Precip Accum (in)"];

    this.precision = options.precision || {
      "DAP": 0,
      "FAW": 2,
      "NDVI": 2,
      "Net Bal (mm)": 1,
      "Precip Accum (in)": 2,
      "Irrigation (mm)": 0
    };

    this.thresholdAlerts = options.thresholdAlerts || [
      { column: "FAW", threshold: 0.30, condition: "below", consecutiveDays: 2, label: "MAD Deficit Alert (<0.30 for 2+ days)" }
    ];

    this.columnAggregates = options.columnAggregates || {
      "FAW": "avg",
      "NDVI": "avg",
      "Net Bal (mm)": "sum",
      "Precip Accum (in)": "sum",
      "Irrigation (mm)": "sum"
    };

    this.conditionalFormatting = options.conditionalFormatting || [
      { column: "FAW", palette: "water_deficit_ylorrd", min: 0.0, max: 0.5, invert: true },
      { column: "NDVI", palette: "plant_growth_modis", min: 0.1, max: 0.9 },
      { column: "Precip Accum (in)", palette: "nws_precipitation_accumulation" },
      { column: "Irrigation (mm)", palette: "irrigation_blues", min: 0, max: 30 }
    ];

    this.onCellChange = options.onCellChange || null;
    this.onNotify = options.onNotify || null;
    this.sortState = { colIndex: -1, asc: true };
    this.activeCell = null;

    this.init();

    if (options.data) {
      this.setData(options.data);
    }
  }

  init() {
    this.render();
    this.bindEvents();
    this.setupResizers();
  }

  render() {
    this.container.innerHTML = `
      <div class="irrigation-module">
        <div class="field-header">
          <div class="title-row-container">
            <h2 class="field-title" id="field-title-text">${this.escapeHtml(this.metadata.fieldName)}</h2>
            <div id="field-status-badge" class="field-status-pill status-${this.metadata.statusLevel || 'optimal'}">
              ${this.escapeHtml(this.metadata.statusLabel || 'STATUS: ADEQUATE MOISTURE')}
            </div>
          </div>
          <div class="metadata-subheader" id="metadata-subheader-text">
            ${this.getMetadataHtml()}
          </div>
        </div>

        <div class="spreadsheet-toolbar">
          <div class="toolbar-group">
            <span>🔒 Locked Model Outputs — read-only columns are computed, not editable</span>
          </div>
        </div>

        <div class="table-container">
          <table class="irrigation-table" id="sheet-table">
            <thead>
              <tr id="table-headers">
                ${this.columns.map((col, idx) => `
                  <th data-col-index="${idx}" style="position: relative;">
                    <div class="th-content">
                      <div class="th-top-row">
                        <div class="sparkline-box" id="sparkline-col-${idx}"></div>
                        <div id="alert-box-col-${idx}"></div>
                      </div>
                      <div class="th-title-row">
                        <span class="th-label" title="${col}">${col}</span>
                        <span>
                          ${idx === this.sortState.colIndex && this.sortState.colIndex >= 0 ? `<span class="sort-indicator">${this.sortState.asc ? '▲' : '▼'}</span>` : ''}
                        </span>
                      </div>
                    </div>
                  </th>
                `).join('')}
              </tr>
            </thead>
            <tbody id="table-body"></tbody>
            <tfoot id="table-footer"></tfoot>
          </table>
        </div>
        <div class="hint">💡 Frozen headers & Date column. Paste vertical columns directly from Excel.</div>
      </div>
    `;
  }

  getMetadataHtml() {
    return `
      <div class="meta-tag"><span class="label">Soil:</span> ${this.escapeHtml(this.metadata.soilType)}</div>
      <span class="meta-divider">|</span>
      <div class="meta-tag"><span class="label">Crop:</span> ${this.escapeHtml(this.metadata.crop)}</div>
      <span class="meta-divider">|</span>
      <div class="meta-tag"><span class="label">Planting Date:</span> ${this.escapeHtml(this.metadata.plantingDate)}</div>
      <span class="meta-divider">|</span>
      <div class="meta-tag"><span class="label">Location:</span> ${this.escapeHtml(this.metadata.latitude)}°, ${this.escapeHtml(this.metadata.longitude)}°</div>
    `;
  }

  bindEvents() {
    const table = this.container.querySelector('#sheet-table');

    table.addEventListener('focusin', (e) => {
      if (e.target.tagName === 'TD' && e.target.contentEditable === "true") {
        this.activeCell = e.target;
      }
    });

    // Cell edits commit once — on Enter or on blur — never mid-keystroke.
    // A contenteditable cell fires no event equivalent to a text input's
    // "change"; the two paths that mean "the user is done with this cell"
    // are focus leaving it, or pressing Enter (which, left alone, inserts
    // a line break in a contenteditable rather than committing — Enter is
    // intercepted below and turned into a blur so both paths land here).
    const commitCell = (cell) => {
      if (cell.tagName !== 'TD' || cell.contentEditable !== 'true') return;
      this.formatCellPrecision(cell);
      this.applyConditionalFormattingToCell(cell);
      this.updateSummaryRow();
      this.updateHeaderIndicators();

      if (typeof this.onCellChange === 'function') {
        const rowIndex = cell.parentElement.rowIndex - 1;
        const colIndex = cell.cellIndex;
        this.onCellChange({
          rowIndex,
          colIndex,
          columnName: this.columns[colIndex],
          value: cell.textContent.trim()
        });
      }
    };

    table.addEventListener('focusout', (e) => commitCell(e.target));

    table.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' && e.target.tagName === 'TD' && e.target.contentEditable === 'true') {
        e.preventDefault();   // don't insert a line break
        e.target.blur();      // triggers the focusout handler above
      }
    });

    const headers = this.container.querySelectorAll('#table-headers th');
    headers.forEach(th => {
      th.addEventListener('click', (e) => {
        if (e.target.classList.contains('col-resizer')) return;
        const colIdx = parseInt(th.getAttribute('data-col-index'), 10);
        this.sortSpreadsheet(colIdx);
      });
    });

    table.addEventListener('paste', (e) => this.handlePaste(e));
  }

  copyColumnToClipboard(colIndex) {
    const tbody = this.container.querySelector('#table-body');
    if (!tbody) return;

    const rows = Array.from(tbody.querySelectorAll('tr'));
    const colValues = rows.map(r => r.children[colIndex] ? r.children[colIndex].textContent.trim() : '');
    const colName = this.columns[colIndex];
    const tsvData = colValues.join('\r\n');

    navigator.clipboard.writeText(tsvData).then(() => {
      this.showToast(`Copied entire '${colName}' column to clipboard!`);
    }).catch(err => {
      console.error("Copy failed", err);
    });
  }

  // A grid module has no business owning its own toast UI — the host app
  // already has (or should have) a place to surface transient messages
  // consistently across its whole UI, not just this one widget. Pass
  // `onNotify(message)` in the constructor options to hook this up to
  // whatever notification system the host app uses; if it's not provided,
  // this just logs, so nothing breaks silently.
  showToast(msg) {
    if (typeof this.onNotify === 'function') this.onNotify(msg);
    else console.log('[IrrigationSpreadsheet]', msg);
  }

  setupResizers() {
    const headers = this.container.querySelectorAll('#table-headers th');
    headers.forEach((th) => {
      let resizer = th.querySelector('.col-resizer');
      if (!resizer) {
        resizer = document.createElement('div');
        resizer.className = 'col-resizer';
        th.appendChild(resizer);
      }

      let startX, startWidth;

      const onMouseDown = (e) => {
        e.preventDefault();
        e.stopPropagation();
        startX = e.pageX;
        startWidth = th.offsetWidth;
        resizer.classList.add('resizing');

        const onMouseMove = (moveEvent) => {
          const diffX = moveEvent.pageX - startX;
          const newWidth = Math.max(65, startWidth + diffX);
          th.style.width = `${newWidth}px`;
          th.style.minWidth = `${newWidth}px`;
        };

        const onMouseUp = () => {
          resizer.classList.remove('resizing');
          document.removeEventListener('mousemove', onMouseMove);
          document.removeEventListener('mouseup', onMouseUp);
        };

        document.addEventListener('mousemove', onMouseMove);
        document.addEventListener('mouseup', onMouseUp);
      };

      resizer.addEventListener('mousedown', onMouseDown);
    });
  }

  formatCellPrecision(cell) {
    const colName = this.columns[cell.cellIndex];
    const decimals = this.precision[colName];

    if (decimals !== undefined) {
      const num = parseFloat(cell.textContent.replace(/[^0-9.-]/g, ''));
      if (!isNaN(num)) {
        cell.textContent = num.toFixed(decimals);
      }
    }
  }

  updateHeaderIndicators() {
    const tbody = this.container.querySelector('#table-body');
    if (!tbody) return;
    const rows = Array.from(tbody.querySelectorAll('tr'));

    this.columns.forEach((colName, colIdx) => {
      const rawValues = rows.map(r => {
        const cell = r.children[colIdx];
        return cell ? parseFloat(cell.textContent.replace(/[^0-9.-]/g, '')) : NaN;
      });

      const validNums = rawValues.filter(v => !isNaN(v));

      const sparklineBox = this.container.querySelector(`#sparkline-col-${colIdx}`);
      if (sparklineBox) {
        if (this.sparklineColumns.includes(colName) && validNums.length > 1) {
          const min = Math.min(...validNums);
          const max = Math.max(...validNums);
          const range = (max - min) || 1;
          const width = 50;
          const height = 12;

          const points = validNums.map((val, i) => {
            const x = (i / (validNums.length - 1)) * width;
            const y = height - ((val - min) / range) * (height - 2) - 1;
            return `${x.toFixed(1)},${y.toFixed(1)}`;
          }).join(' ');

          sparklineBox.innerHTML = `
            <svg width="100%" height="100%" viewBox="0 0 50 14" preserveAspectRatio="none">
              <polyline points="${points}" fill="none" style="stroke:var(--accent-color, #0284c7);" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/>
            </svg>
          `;
        } else {
          sparklineBox.innerHTML = '';
        }
      }

      const alertBox = this.container.querySelector(`#alert-box-col-${colIdx}`);
      if (alertBox) {
        const alertRule = this.thresholdAlerts.find(a => a.column === colName);
        let alertTriggered = false;

        if (alertRule) {
          let consecutiveCount = 0;
          for (let v of rawValues) {
            if (!isNaN(v) && ((alertRule.condition === "below" && v < alertRule.threshold) || (alertRule.condition === "above" && v > alertRule.threshold))) {
              consecutiveCount++;
              if (consecutiveCount >= alertRule.consecutiveDays) {
                alertTriggered = true;
                break;
              }
            } else {
              consecutiveCount = 0;
            }
          }
        }

        if (alertTriggered) {
          alertBox.innerHTML = `<span class="header-alert-badge" title="${alertRule.label}">⚠️ ALERT</span>`;
        } else {
          alertBox.innerHTML = '';
        }
      }
    });
  }

  updateSummaryRow() {
    const tfoot = this.container.querySelector('#table-footer');
    const tbody = this.container.querySelector('#table-body');
    if (!tbody || !tfoot) return;

    const rows = Array.from(tbody.querySelectorAll('tr'));
    if (rows.length === 0) {
      tfoot.innerHTML = '';
      return;
    }

    const footTr = document.createElement('tr');

    this.columns.forEach((colName, colIdx) => {
      const td = document.createElement('td');
      const aggType = this.columnAggregates[colName];

      if (colIdx === 0) {
        td.textContent = "SUMMARY";
      } else if (aggType) {
        let sum = 0;
        let count = 0;

        rows.forEach(r => {
          const valCell = r.children[colIdx];
          if (valCell) {
            const num = parseFloat(valCell.textContent.replace(/[^0-9.-]/g, ''));
            if (!isNaN(num)) {
              sum += num;
              count++;
            }
          }
        });

        const decimals = this.precision[colName] !== undefined ? this.precision[colName] : 1;

        if (count > 0) {
          if (aggType === 'sum') {
            td.textContent = `SUM ${sum.toFixed(decimals)}`;
          } else if (aggType === 'avg') {
            td.textContent = `AVG ${(sum / count).toFixed(decimals)}`;
          }
        } else {
          td.textContent = '-';
        }
      } else {
        td.textContent = '';
      }

      footTr.appendChild(td);
    });

    tfoot.innerHTML = '';
    tfoot.appendChild(footTr);
  }

  applyAllConditionalFormatting() {
    const tbody = this.container.querySelector('#table-body');
    if (!tbody) return;
    const rows = tbody.querySelectorAll('tr');

    rows.forEach(row => {
      Array.from(row.children).forEach(cell => {
        this.applyConditionalFormattingToCell(cell);
      });
    });
  }

  applyConditionalFormattingToCell(cell) {
    const colIndex = cell.cellIndex;
    const colName = this.columns[colIndex];
    const rule = this.conditionalFormatting.find(r => r.column === colName);

    if (!rule) return;

    const rawVal = parseFloat(cell.textContent.replace(/[^0-9.-]/g, ''));
    if (isNaN(rawVal)) {
      cell.style.backgroundColor = '';
      cell.style.color = '';
      return;
    }

    const cmap = COLORMAPS[rule.palette];
    if (!cmap || !cmap.length) return;

    let matchedHex = null;

    if (cmap[0].min !== undefined && cmap[0].max !== undefined) {
      const matchedEntry = cmap.find(item => rawVal >= item.min && rawVal <= item.max);
      if (matchedEntry) matchedHex = matchedEntry.hex;
      else if (rawVal < cmap[0].min) matchedHex = cmap[0].hex;
      else if (rawVal > cmap[cmap.length - 1].max) matchedHex = cmap[cmap.length - 1].hex;
    } 
    else if (rule.min !== undefined && rule.max !== undefined) {
      let pct = (rawVal - rule.min) / (rule.max - rule.min);
      pct = Math.max(0, Math.min(1, pct));
      if (rule.invert) pct = 1 - pct;

      const colorIndex = Math.min(cmap.length - 1, Math.floor(pct * cmap.length));
      matchedHex = cmap[colorIndex].hex;
    }

    if (matchedHex) {
      if (matchedHex === '#00000000') {
        cell.style.backgroundColor = 'transparent';
        cell.style.color = '#0f172a';
      } else {
        cell.style.backgroundColor = matchedHex;
        cell.style.color = getContrastTextColor(matchedHex);
      }
      cell.style.fontSize = '0.825rem';
      cell.style.fontWeight = '500';
    }
  }

  exportCSV() {
    let csvContent = "data:text/csv;charset=utf-8,";
    csvContent += `Field Name,${this.metadata.fieldName}\n`;
    csvContent += `Soil Type,${this.metadata.soilType}\n`;
    csvContent += `Crop,${this.metadata.crop}\n`;
    csvContent += `Planting Date,${this.metadata.plantingDate}\n`;
    csvContent += `Coordinates,${this.metadata.latitude},${this.metadata.longitude}\n\n`;

    csvContent += this.columns.join(",") + "\n";

    const tbody = this.container.querySelector('#table-body');
    const rows = tbody.querySelectorAll('tr');

    rows.forEach(row => {
      const rowVals = Array.from(row.children).map(td => `"${td.textContent.trim().replace(/"/g, '""')}"`);
      csvContent += rowVals.join(",") + "\n";
    });

    const encodedUri = encodeURI(csvContent);
    const link = document.createElement("a");
    link.setAttribute("href", encodedUri);
    link.setAttribute("download", `${this.metadata.fieldName.replace(/[^a-z0-9]/gi, '_')}_irrigation_data.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  }

  sortSpreadsheet(colIndex) {
    if (this.sortState.colIndex === colIndex) {
      this.sortState.asc = !this.sortState.asc;
    } else {
      this.sortState.colIndex = colIndex;
      this.sortState.asc = true;
    }

    const tbody = this.container.querySelector('#table-body');
    const rows = Array.from(tbody.querySelectorAll('tr'));

    rows.sort((rowA, rowB) => {
      const valA = rowA.children[colIndex].textContent.trim();
      const valB = rowB.children[colIndex].textContent.trim();

      const timeA = Date.parse(valA);
      const timeB = Date.parse(valB);

      let comparison = 0;
      if (!isNaN(timeA) && !isNaN(timeB)) {
        comparison = timeA - timeB;
      } else {
        const numA = parseFloat(valA);
        const numB = parseFloat(valB);
        if (!isNaN(numA) && !isNaN(numB)) {
          comparison = numA - numB;
        } else {
          comparison = valA.localeCompare(valB);
        }
      }

      return this.sortState.asc ? comparison : -comparison;
    });

    rows.forEach(r => tbody.appendChild(r));
    this.updateSortHeaders();
  }

  updateSortHeaders() {
    const headers = this.container.querySelectorAll('#table-headers th');
    headers.forEach((th, idx) => {
      let indicator = th.querySelector('.sort-indicator');
      if (idx === this.sortState.colIndex) {
        if (!indicator) {
          indicator = document.createElement('span');
          indicator.className = 'sort-indicator';
          th.querySelector('.th-title-row').appendChild(indicator);
        }
        indicator.textContent = this.sortState.asc ? ' ▲' : ' ▼';
      } else if (indicator) {
        indicator.remove();
      }
    });
  }

  setMetadata(newMeta = {}) {
    this.metadata = Object.assign(this.metadata, newMeta);
    this.container.querySelector('#field-title-text').textContent = this.metadata.fieldName;
    this.container.querySelector('#metadata-subheader-text').innerHTML = this.getMetadataHtml();

    const badge = this.container.querySelector('#field-status-badge');
    if (badge && (newMeta.statusLabel || newMeta.statusLevel)) {
      badge.textContent = this.metadata.statusLabel || 'STATUS: ADEQUATE MOISTURE';
      badge.className = `field-status-pill status-${this.metadata.statusLevel || 'optimal'}`;
    }
  }

  setData(dataRows = []) {
    this.clearRows();
    dataRows.forEach(row => this.addRow(row));
    this.applyAllConditionalFormatting();
    this.updateSummaryRow();
    this.updateHeaderIndicators();
  }

  addRow(rowData = []) {
    const tbody = this.container.querySelector('#table-body');
    if (!tbody) return;
    const tr = document.createElement('tr');

    this.columns.forEach((colName, cIdx) => {
      const td = document.createElement('td');
      const isReadOnly = this.readOnlyColumns.includes(colName);

      td.contentEditable = isReadOnly ? "false" : "true";
      if (isReadOnly) td.classList.add('read-only-cell');

      let val = rowData[cIdx] !== undefined ? rowData[cIdx] : '';

      if (this.precision[colName] !== undefined && val !== '') {
        const num = parseFloat(val.toString().replace(/[^0-9.-]/g, ''));
        if (!isNaN(num)) val = num.toFixed(this.precision[colName]);
      }

      td.textContent = val;
      tr.appendChild(td);
    });

    tbody.appendChild(tr);
  }

  clearRows() {
    const tbody = this.container.querySelector('#table-body');
    if (tbody) tbody.innerHTML = '';
  }

  handlePaste(e) {
    e.preventDefault();
    const clipboardData = e.clipboardData || window.clipboardData;
    const pastedText = clipboardData.getData('text');
    if (!pastedText) return;

    const rows = pastedText.split(/\r\n|\n|\r/).map(row => row.split('\t'));
    let startRowIndex = this.activeCell ? this.activeCell.parentElement.rowIndex - 1 : 0;
    let startColIndex = this.activeCell ? this.activeCell.cellIndex : 0;

    const tbody = this.container.querySelector('#table-body');
    const tableRows = tbody.querySelectorAll('tr');

    rows.forEach((rowValues, rOffset) => {
      if (rowValues.length === 1 && rowValues[0] === "") return;

      let targetRow = tableRows[startRowIndex + rOffset];
      if (!targetRow) {
        this.addRow();
        targetRow = tbody.querySelectorAll('tr')[startRowIndex + rOffset];
      }

      rowValues.forEach((val, cOffset) => {
        const targetCellIndex = startColIndex + cOffset;
        let targetCell = targetRow.children[targetCellIndex];
        
        if (targetCell && targetCell.contentEditable === "true") {
          targetCell.textContent = val.trim();
          this.formatCellPrecision(targetCell);
          this.applyConditionalFormattingToCell(targetCell);
        }
      });
    });

    this.updateSummaryRow();
    this.updateHeaderIndicators();
  }

  escapeHtml(str) {
    return String(str || '').replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  }
}

// This library is written to be loaded via a plain <script src="IrrigationSpreadsheet.js">
// tag — that's its primary, intended environment. The block below adds CommonJS support
// (`require('./IrrigationSpreadsheet.js')`) as a courtesy for newer/bundled JS setups, using
// a runtime `typeof` check rather than an `export` statement: `export` is parse-time-only and
// can't be feature-detected, so an actual ES `export` here would break the plain-script-tag
// case outright (see file header). CommonJS's `module.exports` has no such restriction — it's
// just a normal object write, safe to leave in a browser context where `module` doesn't exist.
if (typeof module !== 'undefined' && module.exports) {
  module.exports = IrrigationSpreadsheet;
}
if (typeof window !== 'undefined') {
  window.IrrigationSpreadsheet = IrrigationSpreadsheet;
}