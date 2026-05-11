# @libraz/formulon-cell

Spreadsheet UI for the [formulon](https://github.com/libraz/formulon) WASM
calc engine.

- **Desktop-spreadsheet-style** chrome out of the box (formula bar, status bar,
  context menu, sheet tabs).
- **Canvas-rendered** grid with theme tokens — `paper` (light) and `ink`
  (dark) ship in the box; bring your own with the documented CSS variables.
- **Extension-based** API: built-ins are controlled with feature flags, and
  replaceable pieces (find/replace, format dialog, paste-special, hyperlink
  dialog, hover comments, View toolbar, Quick Analysis, PivotTable creation, …) are available as
  extension factories you can compose into the mount call.
- **Runtime i18n** — swap locales without re-mounting; `ja` and `en` ship
  by default, register more at runtime.
- **Headless option** — keep just the canvas + store and provide your own
  chrome.

> **Status:** v0.1 — public API stabilizing. Until v1.0 minor bumps may
> reshape extension contracts. Pin a version range you can update on
> purpose.

## Install

```sh
npm install @libraz/formulon-cell zustand
# or yarn / pnpm
```

`zustand` is a peer dependency — exposed because consumers can read from
the same store the chrome subscribes to.

The WASM engine ships pthread-enabled and requires a
[crossOriginIsolated context](https://developer.mozilla.org/docs/Web/API/crossOriginIsolated)
(`Cross-Origin-Opener-Policy: same-origin` + `Cross-Origin-Embedder-Policy:
require-corp`). Without it, formulon-cell falls back to an in-memory stub
engine — the UI keeps working, formulas degrade gracefully.

## Bundler integration (Vite, webpack, esbuild)

formulon-cell re-uses `@libraz/formulon`'s pthread-enabled WASM module, so
the bundler hygiene rules from the engine package apply here too. Four
things matter:

**1. Workers must ship as ES modules.** The recalc scheduler runs on Web
Workers spawned by Emscripten with
`new Worker(new URL(...), { type: 'module' })`. Bundlers default to
classic (IIFE) workers and must be told otherwise:

```ts
// vite.config.ts
export default defineConfig({
  worker: { format: 'es' },
});
```

webpack 5 picks up `{ type: 'module' }` automatically when
`output.module: true`. esbuild needs `--format=esm` for the worker chunk.

**2. Top-level await + dynamic node imports need an es2022 target.** The
engine factory uses TLA and conditional `await import('node:...')`. Lift
both the main and worker target:

```ts
// vite.config.ts
export default defineConfig({
  build: { target: 'es2022' },
});
```

**3. Keep the engine out of dependency pre-bundling.** formulon-cell imports
`@libraz/formulon`, whose Emscripten wrapper owns the worker/WASM asset
resolution. Keep both packages out of dependency pre-bundling so those
assets stay under the app bundler's control:

```ts
// vite.config.ts
export default defineConfig({
  optimizeDeps: { exclude: ['@libraz/formulon-cell', '@libraz/formulon'] },
});
```

**4. SharedArrayBuffer requires cross-origin isolation.** Serve your page
with `Cross-Origin-Opener-Policy: same-origin` + `Cross-Origin-Embedder-Policy:
require-corp`. Without these headers, `SharedArrayBuffer` is undefined and
formulon-cell drops to an in-memory **stub engine**: the canvas, formula
bar, and editing affordances all keep working, but formula evaluation,
recalc, and xlsx round-trip degrade to no-ops. Detect at runtime via
`crossOriginIsolated` or via `isUsingStub()` after `WorkbookHandle.createDefault()`:

```ts
import { WorkbookHandle, isUsingStub } from '@libraz/formulon-cell';

const wb = await WorkbookHandle.createDefault();
if (isUsingStub()) {
  console.warn('formulon-cell: running on stub engine — recalc disabled');
}
```

## Quick start

```ts
import { Spreadsheet, WorkbookHandle, presets } from '@libraz/formulon-cell';
import '@libraz/formulon-cell/styles.css';

const host = document.getElementById('sheet')!;
const wb = await WorkbookHandle.createDefault();
const sheet = await Spreadsheet.mount(host, {
  workbook: wb,
  features: presets.full(),
  locale: 'en',
});

sheet.i18n.setLocale('ja');     // runtime locale swap
sheet.setTheme('ink');           // dark mode
```

## Presets

| preset | what's in it |
|--------|--------------|
| `presets.minimal()`  | formula bar, status bar, basic keymap |
| `presets.standard()` | + View toolbar, Quick Analysis, session chart overlays, workbook object inspector, context menu, find/replace, clipboard, format painter, wheel scroll |
| `presets.full()`    | + format dialog, paste-special, conditional formatting, iterative calculation settings, Go To Special, page setup, named ranges, hyperlink dialog, PivotTable creation, validation, autocomplete, hover comments, spreadsheet keymap |

Compose your own:

```ts
import {
  Spreadsheet,
  contextMenu,
  findReplace,
  pasteSpecial,
  presets,
  quickAnalysis,
  statusBar,
  viewToolbar,
  workbookObjects,
} from '@libraz/formulon-cell';

await Spreadsheet.mount(host, {
  features: {
    ...presets.minimal(),
    statusBar: false,
    viewToolbar: false,
    workbookObjects: false,
    quickAnalysis: false,
    contextMenu: false,
    findReplace: false,
    pasteSpecial: false,
  },
  extensions: [
    statusBar(),
    workbookObjects(),
    viewToolbar(),
    quickAnalysis(),
    contextMenu(),
    findReplace(),
    pasteSpecial(),
  ],
});
```

`allBuiltIns()` returns the default-on replaceable built-ins as extension
factories. Default-off panels such as `watchWindow()` and `slicer()` are
exported separately so apps can opt into them deliberately.

## Spreadsheet compatibility

formulon-cell aims for a full-featured spreadsheet surface, not a
pixel-for-pixel clone. The current implementation covers the reusable core
workflows most host apps need:

- workbook-backed grid painting, editing, formula bar, name box, sheet tabs,
  View toolbar, selection, fill handle, undo/redo, copy/cut/paste, paste
  special
- spreadsheet-like keyboard shortcuts for navigation, formatting, fill down/right,
  recalc, R1C1 toggle, comments, hyperlink, and dialogs
- number formats, font/fill/borders/alignment/wrap, named cell styles,
  conditional-format evaluation/authoring, data validation, hyperlinks,
  comments, merged cells, freeze panes, zoom, hidden rows/columns, outlines
- find/replace, Go To / Go To Special, Format Cells, Named Ranges, Watch
  Window, iterative calculation settings, external links summary
- toolbar-ready View commands plus built-in View toolbar for gridlines,
  headings, show formulas, R1C1, freeze panes, zoom, Sheet View save/restore,
  and status-bar aggregates
- read-only workbook object inspector for preserved charts, drawings, pivot
  parts, media, threaded comments, query tables, and loaded spreadsheet Tables
- session chart overlays from Quick Analysis for lightweight column/line chart
  previews while chart writeback remains engine-gated
- xlsx metadata round-trip surfaces when the formulon engine exposes them,
  including sheet views, protection, loaded spreadsheet Table visuals, passthrough
  OOXML parts, and read-only PivotTable layout projection
- `summarizeSpreadsheetCompatibility(workbook)` reports which surfaces are writable,
  read-only, session-only, or unsupported for the current engine bundle,
  including formatting, validation, hyperlinks, comments, defined names, sheet
  protection, views, tables, PivotTable surfaces, charts/drawings, and external
  links
- `spreadsheetCompatibilityStatus(summary, id)`,
  `isSpreadsheetFeatureWritable(summary, id)`, and
  `isSpreadsheetFeatureAvailable(summary, id)` let custom toolbars and menus use
  the same compatibility decisions as the built-in object inspector
- `createSessionChart(store, range, options)` gives host chrome a reusable
  command for creating the same UI-owned column/line chart overlays used by
  Quick Analysis
- `enabledQuickAnalysisActions(input)`, `quickAnalysisActionById(actions, id)`,
  and `isQuickAnalysisActionEnabled(input, id)` let custom Quick Analysis
  surfaces reuse the built-in availability rules
- `listSessionCharts(state)`, `sessionChartById(state, id)`,
  `sessionChartSeries(state, chartOrRange)`, `setSparkline(store, addr, spec)`,
  and `listSparklines(state)` expose session chart data and sparkline state for
  custom chart panes
- `pageSetupForSheet(state, sheet)`, `setPageSetup(store, sheet, patch)`,
  `setPrintTitleRows(store, sheet, rows)`, and `listPageSetups(state)` expose
  the print/page setup state used by the built-in dialog
- `saveSheetView(store, id, name)`, `activateSheetView(store, id)`, and
  `deleteSheetView(store, id)` let host chrome reuse the same session Sheet
  View state used by the built-in View toolbar
- `listDefinedNames(workbook)`, `upsertDefinedName(workbook, name, formula)`,
  and `deleteDefinedName(workbook, name)` provide a headless Name Manager API
  matching the built-in named-range dialog
- `ignoreCellError(store, addr)`, `restoreCellErrorIndicator(store, addr)`,
  and `clearIgnoredCellErrors(store)` expose the same session error-indicator
  suppression used by the built-in error menu
- `tracePrecedents(store, workbook, addr)`, `traceDependents(store, workbook, addr)`,
  and `clearTraceArrows(store)` expose the same session trace-arrow state used
  by the built-in instance methods
- `visibleStatusAggregates(state)`, `statusAggregateValue(key, stats)`, and
  `STATUS_AGGREGATE_KEYS` expose the same selection-summary logic used by the
  built-in status bar
- `createSlicer(store, workbook, options)`, `setSlicerSelected(store, id, values)`,
  and `recomputeSlicerFilters(store, workbook)` expose the built-in slicer
  state/filter behavior for custom host chrome
- `setProtectedSheet(store, sheet, on, { workbook, password })` and
  `toggleProtectedSheet(store, sheet, options)` expose the same sheet-protection
  state and optional engine writeback used by the built-in instance methods
- `listExternalLinks(workbook)` and `summarizeExternalLinks(workbook)` expose
  the same read-only external-reference inventory used by the built-in link
  inspector
- `listWorkbookObjects(workbook)`, `workbookObjectsByKind(objects)`,
  `workbookObjectKindCounts(objects)`, `workbookObjectKindLabel(kind)`, and
  `WORKBOOK_OBJECT_KINDS` expose preserved workbook object parts as classified
  records for custom object browsers and compatibility badges, including
  charts, drawings, media, comments, slicers, timelines, connections, external
  links, controls, print settings, custom XML, and macro project parts
- `setHyperlink(store, addr, target, workbook)`, `clearHyperlink(store, addr)`,
  and `listHyperlinks(state, sheet)` expose the same hyperlink state used by
  the built-in hyperlink dialog
- `listComments(state, sheet)` exposes cell comments for custom sidebars and
  review panes alongside `setComment()` / `clearComment()`
- `addConditionalRule(store, rule)`, `listConditionalRules(state)`, and
  `clearConditionalRulesInRange(store, range)` expose session conditional
  formatting rules for custom rule managers
- `listTableOverlays(state)`, `tableOverlayAt(state, sheet, row, col)`, and
  `updateTableOverlay(store, id, patch)` expose loaded read-only table overlays
  and session Format-as-Table overlays for custom object panes and table tools
- `attachErrorMenu()` exposes `onTraceError` so hosts can wire the error menu
  into custom trace-arrow surfaces; the default mount traces same-sheet
  precedents for the error cell
- low-level PivotCache / PivotTable mutation wrappers are exposed on
  `WorkbookHandle`, plus `createPivotTableFromRange()` and the default
  `pivotTableDialog` extension for selection-based PivotTable creation.
  Field-level wrappers cover sort, subtotal position/functions, manual
  item visibility, date grouping, and field number formats.
- React and Vue wrappers with matching props/events/composables and
  non-remounting runtime updates for workbook/theme/locale/strings/features

Known compatibility gaps are intentionally surfaced rather than hidden:

- Ribbon UI is not bundled. Host apps compose their own chrome via
  `SpreadsheetInstance`, commands, and extension handles.
- Chart/drawing/image authoring is not implemented yet. Existing OOXML parts
  are preserved by the engine and surfaced as passthrough summaries.
- PivotTables are projected when loaded from xlsx, and selection-based
  PivotTable creation is bundled when the engine exposes mutation APIs.
  The UI currently covers row/column/value field selection, Sum/Count,
  row/column sorting, subtotal placement, value number format, and
  row/column grand-total toggles; grouping, slicers, and layout styling are
  still host-extensible surfaces.
- Quick Analysis has a spreadsheet-style floating panel (`Ctrl+Q`) and applies
  conditional formats, totals, session Format-as-Table overlays, session
  column/line charts, and sparklines. Persisted chart authoring remains
  disabled until a chart engine is writable, and PivotTable creation opens the
  bundled dialog when available.
- Sheet Views are session-owned in the UI today. Freeze/zoom/hidden row-column
  settings can still persist individually when the engine exposes those view
  APIs, but named view records do not yet round-trip as workbook metadata.
- spreadsheet Tables loaded from xlsx render as read-only table overlays and
  session Format-as-Table overlays are available; full ListObject authoring
  and structured-reference writeback are still pending.
- Threaded comments, co-authoring, live collaboration, macros, Power Query,
  full chart engine, and full ribbon command parity are outside the current
  v0.1 surface.

Good next increments for tighter compatibility are:

1. spreadsheet Tables authoring backed by formulon table APIs once create/update
   bindings are available.
2. Richer chart/drawing previews before full authoring, so preserved workbook
   content is not only listed but visually inspectable inside the UI.
3. Richer PivotTable field settings, grouping, refresh, and layout styling
   on top of the bundled creation flow.

## i18n

```ts
import { Spreadsheet } from '@libraz/formulon-cell';
import ja from '@libraz/formulon-cell/i18n/ja';
import en from '@libraz/formulon-cell/i18n/en';

const sheet = await Spreadsheet.mount(host, { locale: 'en' });

// Swap locale at runtime — every label updates in place.
sheet.i18n.setLocale('ja');

// Override a few strings without forking the dictionary.
sheet.i18n.extend('ja', { contextMenu: { copy: 'コピーする' } });

// Register a brand new locale.
import fr from './fr.js';
sheet.i18n.register('fr', fr);
sheet.i18n.setLocale('fr');
```

## React / Vue

```sh
npm install @libraz/formulon-cell-react react react-dom
# or
npm install @libraz/formulon-cell-vue vue
```

See [`@libraz/formulon-cell-react`](../formulon-cell-react/README.md) and
[`@libraz/formulon-cell-vue`](../formulon-cell-vue/README.md).

## License

[Apache-2.0](./LICENSE)
