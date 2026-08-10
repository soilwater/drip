# `IrrigationSpreadsheet` Documentation

A lightweight, zero-dependency **vanilla JavaScript, HTML, and CSS** spreadsheet component optimized for irrigation planning, agronomic modeling, and soil water balance web applications.

It provides a paper-like table UI with hairline gridlines, Excel copy/paste support, sticky freeze headers, per-column sparklines, threshold alerts, precision formatting, and colormap formatting (NEXRAD, NWS, MODIS, ColorBrewer).

---

## Table of Contents
- [Architectural Philosophy](#architectural-philosophy)
- [Module Loading](#module-loading)
- [Quick Start](#quick-start)
- [Configuration Options](#configuration-options)
- [Cell Editing Behavior](#cell-editing-behavior)
- [Unified Colormap Registry](#unified-colormap-registry)
- [Instance Methods](#instance-methods)
- [Integration Examples](#integration-examples)
- [Guidelines for AI Coding Assistants](#guidelines-for-ai-coding-assistants)

---

## Architectural Philosophy

This library enforces a strict **Separation of Concerns**:

1. **Agronomic Model Layer (External)**:
   * Responsible for all physical calculations: $ET_c$, $K_c$, Days After Planting ($DAP$), soil water balance equations ($P + I - ET_c$), and deficit calculations.
2. **UI Spreadsheet Layer (`IrrigationSpreadsheet`)**:
   * Responsible purely for visual presentation: rendering grid lines, conditional formatting, text-contrast calculations, number precision formatting, threshold alert badges, summary totals/averages, Excel clipboard interactions, and CSV export.

This separation extends to notifications: the component has no opinion about how your app surfaces a message like "column copied" — see [`onNotify`](#configuration-options) below.

---

## Module Loading

This library is written to be loaded via a plain script tag — that's its primary, intended environment:

```html
<script src="IrrigationSpreadsheet.js"></script>
<script>
  const sheet = new IrrigationSpreadsheet('irrigation-sheet-wrapper', { /* ... */ });
</script>
```

That tag must **not** have `type="module"`. The file ends with a runtime-guarded assignment to `window.IrrigationSpreadsheet`, plus a CommonJS-compatible `module.exports` for bundler/Node environments (`require('./IrrigationSpreadsheet.js')`). It does **not** use an ES `export` statement, and it shouldn't gain one: `export` is parse-time-only syntax — a JS engine rejects the whole file with `SyntaxError: Unexpected token 'export'` before any runtime `typeof` check even runs, so it can't be feature-detected or guarded the way `module.exports` can. Adding `export default IrrigationSpreadsheet;` (even behind what looks like a safe conditional) will break every plain `<script src>` consumer of this file outright. If you need real ESM `import`/`export`, wrap this file from a separate build entry point rather than editing it here.

---

## Quick Start

### 1. HTML Container
Include a container `<div>` where you want the spreadsheet to render:

```html
<div id="irrigation-sheet-wrapper"></div>
```

### 2. Instantiate and Load Data
```javascript
// Initialize Component
const sheet = new IrrigationSpreadsheet('irrigation-sheet-wrapper', {
  metadata: {
    fieldName: "North Field - Pivot 1",
    soilType: "Silty Clay Loam",
    crop: "Corn (Maize)",
    plantingDate: "2026-04-15",
    latitude: "40.7128",
    longitude: "-74.0060",
    statusLabel: "STATUS: ADEQUATE MOISTURE",
    statusLevel: "optimal" // 'optimal' | 'warning' | 'critical'
  },
  columns: ["Date", "DAP", "FAW", "NDVI", "Reflectivity (dBZ)", "Precip Accum (in)", "Irrigation (mm)", "Notes"],
  sparklineColumns: ["FAW", "NDVI", "Reflectivity (dBZ)", "Precip Accum (in)", "Irrigation (mm)"],
  readOnlyColumns: ["Date", "DAP", "Reflectivity (dBZ)", "Precip Accum (in)"],
  precision: { "DAP": 0, "FAW": 2, "NDVI": 2, "Reflectivity (dBZ)": 0, "Precip Accum (in)": 2, "Irrigation (mm)": 0 }
});

// Pass computed model data
sheet.setData([
  ["2026-05-01", "16", "0.85", "0.15", "12", "0.00", "0", "Early emergence"],
  ["2026-05-02", "17", "0.60", "0.28", "22", "0.15", "0", "V2 Stage"],
  ["2026-05-03", "18", "0.42", "0.42", "38", "0.60", "0", "Entering mild stress"]
]);
```

---

## Configuration Options

When instantiating `new IrrigationSpreadsheet(containerId, options)`, pass an `options` object with the following properties:

| Option | Type | Default | Description |
| :--- | :--- | :--- | :--- |
| `metadata` | `Object` | `{}` | Field metadata displayed in the header (`fieldName`, `soilType`, `crop`, `plantingDate`, `latitude`, `longitude`, `statusLabel`, `statusLevel`). |
| `columns` | `Array<string>` | `[...]` | Column header titles in left-to-right order. |
| `readOnlyColumns` | `Array<string>` | `[]` | Array of column names that cannot be edited by the user (sets `contenteditable="false"`). |
| `sparklineColumns` | `Array<string>` | `[]` | List of column names that should render a mini SVG trend chart in their header. |
| `precision` | `Object` | `{}` | Fixed decimal places per column. Example: `{ "FAW": 2, "Irrigation (mm)": 0 }`. |
| `thresholdAlerts` | `Array<Object>` | `[]` | Rule definitions to trigger red header warning badges. |
| `conditionalFormatting` | `Array<Object>` | `[]` | Colormap rules matching column names to palettes. |
| `columnAggregates` | `Object` | `{}` | Defines whether a summary footer cell calculates `"sum"` or `"avg"`. Leave empty (or omit) to hide the summary footer row entirely. |
| `onCellChange` | `Function` | `null` | Called once per **committed** edit — see [Cell Editing Behavior](#cell-editing-behavior). Not called on paste (paste has no per-cell commit step; call `sheet.setData()` yourself afterward if you need to react to it). |
| `onNotify` | `Function` | `null` | Called with a single string message whenever the component wants to tell the user something (currently: `copyColumnToClipboard`'s "Copied ... to clipboard!"). The component has no built-in toast/notification UI of its own — wire this to whatever your app already uses for transient messages. If omitted, messages just go to `console.log`. |

---

## Cell Editing Behavior

An editable cell is a `contenteditable` `<td>`, not a form input, so "the user is done editing" isn't a single native event the way it is for `<input>`. This component defines it as **either of**:

- **Blur** — the user clicks or tabs away from the cell.
- **Enter** — intercepted so it commits the cell instead of inserting a line break (a bare contenteditable's default behavior for Enter).

Both paths converge on one internal commit step that, in order: reformats the cell to its configured `precision`, re-applies `conditionalFormatting`, refreshes the summary row and sparklines/threshold badges, and — only then — calls `onCellChange`. Nothing in that list runs mid-keystroke. This matters because it makes it safe to do expensive work — including calling `sheet.setData()` — directly inside `onCellChange` (see [Example 1](#example-1-listening-for-user-edits-to-re-run-soil-water-balance-model)): it can only run once the user has actually finished typing, not on every character.

Locked cells (`readOnlyColumns`, or any cell you've set `contentEditable = 'false'` on directly, e.g. to lock a specific row) never reach the commit step at all — Enter and blur are no-ops on them.

---

## Unified Colormap Registry

All colormaps are registered inside `COLORMAPS`.

```javascript
// Rule format examples
conditionalFormatting: [
  // 1. Normalized Continuous Colormaps (min -> max mapping)
  { column: "FAW", palette: "water_deficit_ylorrd", min: 0.0, max: 0.5, invert: true },
  { column: "NDVI", palette: "plant_growth_modis", min: 0.1, max: 0.9 },
  { column: "Irrigation (mm)", palette: "irrigation_blues", min: 0, max: 30 },

  // 2. Bracketed Colormaps (auto-matches raw cell value against min/max ranges)
  { column: "Reflectivity (dBZ)", palette: "nexrad_base_reflectivity" },
  { column: "Precip Accum (in)", palette: "nws_precipitation_accumulation" }
]
```

Conditional formatting is a fixed colormap/bracket lookup keyed on a single cell's own numeric value — it has no visibility into other columns or external state. If a column's status depends on more than its own value (e.g. a threshold that varies per row, or a flag computed elsewhere in your model), color that column yourself after `setData()` by writing directly to the cell's `style.backgroundColor`/`style.color` instead of trying to express it as a `conditionalFormatting` rule.

### Available Palettes in `COLORMAPS`:
* `nexrad_base_reflectivity`: Radar reflectivity ranges (-30 to +80 dBZ).
* `nws_precipitation_accumulation`: NWS accumulated rainfall (0.00 to 99.0 inches). Includes transparency support (`#00000000`).
* `plant_growth_modis`: 16-step MODIS EVI/NDVI vegetation palette (tuned to terminate in deep forest green).
* `water_deficit_ylorrd`: ColorBrewer Yellow-Orange-Red 9-class palette.
* `irrigation_blues`: ColorBrewer Blues 9-class palette.
* `soil_moisture_ylgnbu`: ColorBrewer Yellow-Green-Blue 9-class palette.

---

## Instance Methods

### `sheet.setData(dataRows)`
Overwrites the current table data with new rows from your model run and re-evaluates all conditional formatting, summary aggregates, sparklines, and status badges.
```javascript
sheet.setData([
  ["2026-05-01", "16", "0.85", "0.15", "12", "0.00", "0", "V1 Stage"],
  ["2026-05-02", "17", "0.60", "0.28", "22", "0.15", "0", "V2 Stage"]
]);
```
Note: this fully rebuilds `<tbody>`. If called from inside `onCellChange`, that's fine — see [Cell Editing Behavior](#cell-editing-behavior) — but don't call it from anything that fires more often than a user's own commit action, or you'll rebuild the table out from under whatever the user is doing.

### `sheet.setMetadata(newMetadata)`
Updates field metadata text and executive field status badge.
```javascript
sheet.setMetadata({
  fieldName: "South Pivot - Sector B",
  statusLabel: "STATUS: CRITICAL WATER DEFICIT",
  statusLevel: "critical" // 'optimal' | 'warning' | 'critical'
});
```

### `sheet.addRow(rowData)`
Appends a single row to the bottom of the table.
```javascript
sheet.addRow(["2026-05-03", "18", "0.42", "0.42", "38", "0.60", "25", "Irrigated"]);
```

### `sheet.clearRows()`
Removes all data rows from `<tbody>`.

### `sheet.exportCSV()`
Generates and downloads a `.csv` file containing field metadata and formatted spreadsheet rows.

### `sheet.copyColumnToClipboard(colIndex)`
Programmatically copies a full column's dataset (as vertical TSV lines) directly to the user's system clipboard for instant pasting into Excel. On success, calls `onNotify` (if provided) with a confirmation message.

---

## Integration Examples

### Example 1: Listening for User Edits to Re-Run Soil Water Balance Model
`onCellChange` only fires once the user has committed the edit (Enter or blur — see [Cell Editing Behavior](#cell-editing-behavior)), so it's safe to call `sheet.setData()` directly inside it without fighting the user's own typing.
```javascript
const sheet = new IrrigationSpreadsheet('irrigation-sheet-wrapper', {
  columns: ["Date", "FAW", "ETc (mm)", "Irrigation (mm)"],
  readOnlyColumns: ["Date", "FAW", "ETc (mm)"], // Only 'Irrigation (mm)' is editable

  onCellChange: (info) => {
    console.log(`User committed ${info.columnName} at Row ${info.rowIndex} to ${info.value}`);

    // Re-run your agronomic model using updated user input
    const updatedModelOutput = myAgronomicModel.recalculate({
      editedRow: info.rowIndex,
      newIrrigationValue: parseFloat(info.value)
    });

    // Pass updated simulation rows back into spreadsheet
    sheet.setData(updatedModelOutput.rows);
    sheet.setMetadata({
      statusLabel: updatedModelOutput.statusLabel,
      statusLevel: updatedModelOutput.statusLevel
    });
  }
});
```

### Example 2: Configuring Threshold Alert Badges
```javascript
const sheet = new IrrigationSpreadsheet('irrigation-sheet-wrapper', {
  thresholdAlerts: [
    // Displays red ⚠️ ALERT badge if FAW < 0.30 for 2+ consecutive days
    { 
      column: "FAW", 
      threshold: 0.30, 
      condition: "below", 
      consecutiveDays: 2, 
      label: "Management Allowed Depletion (MAD) Exceeded" 
    }
  ]
});
```

### Example 3: Routing Notifications Through Your App's Own Toast
The component has no toast UI of its own. Wire `onNotify` to whatever your app already uses so messages look and behave consistently across the whole UI, not just inside this widget.
```javascript
const sheet = new IrrigationSpreadsheet('irrigation-sheet-wrapper', {
  onNotify: (message) => myApp.showToast(message)
});

// Later, e.g. from a "copy" button elsewhere in your UI:
sheet.copyColumnToClipboard(3); // -> myApp.showToast("Copied entire 'Irrigation (mm)' column to clipboard!")
```

---

## Guidelines for AI Coding Assistants

When using AI agents (Cursor, GitHub Copilot, ChatGPT, Claude) to build features around this component, follow these rules:

1. **Do NOT write soil water balance math inside this component**:
   * All daily $ET_c$, $DAP$, $K_c$, and soil moisture balance calculations must remain in an external model module/class.
   * Simply pass the calculated outputs into `sheet.setData()`.
2. **Data Structure**:
   * Rows passed to `setData()` must be 2D arrays where indices match the `columns` array order.
3. **Pasting Data**:
   * `handlePaste` handles tab-separated vertical/horizontal Excel pastes natively. Do not add external clipboard parsing libraries.
4. **Colormap Additions**:
   * If adding new palettes, append them directly to the `COLORMAPS` global registry using either hex arrays or range objects (`min`, `max`, `hex`). The text contrast calculation (`getContrastTextColor`) handles text color automatically.
5. **Do NOT add an ES `export` statement to this file**:
   * This library ships as a plain `<script src="...">` global (`window.IrrigationSpreadsheet`), with `module.exports` provided as a courtesy for CommonJS consumers. `export`/`export default` is parse-time-only syntax with no runtime guard possible — adding it breaks the plain-script-tag case outright, for every consumer, not just bundler users. If ESM is genuinely needed somewhere, wrap this file from that build's own entry point instead of editing it here.
6. **Do NOT hook expensive or DOM-rebuilding work to anything that fires per keystroke**:
   * `onCellChange` already only fires once per commit (Enter or blur) — see [Cell Editing Behavior](#cell-editing-behavior) — so this is handled correctly out of the box. The rule is for anyone tempted to *add* a live/`input`-level listener back in for some new feature: don't, unless it's cheap, purely visual, and idempotent. Anything that calls `setData()`, touches the DOM tree structurally, or does real computation belongs on the commit path only.
7. **Notifications go through `onNotify`, not a new toast implementation**:
   * If a new feature needs to tell the user something, call `this.showToast(message)` internally (which delegates to `onNotify`) rather than building another notification element. This component should never own its own persistent, always-in-the-DOM notification UI.
