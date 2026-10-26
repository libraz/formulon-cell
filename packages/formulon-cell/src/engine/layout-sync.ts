import type { LayoutSnapshot } from '../commands/slice-history.js';
import { clampZoom } from '../store/render-scale.js';
import type { LayoutSlice, SpreadsheetStore } from '../store/store.js';
import { SheetVisibility } from './types.js';
import type { WorkbookHandle } from './workbook-handle.js';

type AxisDefaults = Pick<LayoutSlice, 'defaultColWidth' | 'defaultRowHeight'>;

/** Canonical size defaults when the engine reports no sheet geometry: Mac
 *  75 pt × 20 pt, default profile 48 pt × 15 pt. */
const MAC_AXIS_DEFAULTS: AxisDefaults = { defaultColWidth: 100, defaultRowHeight: 80 / 3 };
const DEFAULT_AXIS_DEFAULTS: AxisDefaults = { defaultColWidth: 64, defaultRowHeight: 20 };

const engineAxisDefaults = (wb: WorkbookHandle, sheet: number): AxisDefaults | null =>
  typeof wb.getSheetDefaultLayoutPx === 'function' ? wb.getSheetDefaultLayoutPx(sheet) : null;

const platformAxisDefaults = (wb: WorkbookHandle): AxisDefaults =>
  wb.geometryPlatform === 'mac' ? MAC_AXIS_DEFAULTS : DEFAULT_AXIS_DEFAULTS;

type ColumnLayouts = ReturnType<WorkbookHandle['getColumnLayouts']>;
type RowLayouts = ReturnType<WorkbookHandle['getRowLayouts']>;

const addColumnWidths = (widths: Map<number, number>, cols: ColumnLayouts): Map<number, number> => {
  for (const c of cols) {
    if (!(c.width > 0 && c.hasWidth !== false)) continue;
    for (let col = c.first; col <= c.last; col += 1) widths.set(col, c.width);
  }
  return widths;
};

const addRowHeights = (heights: Map<number, number>, rows: RowLayouts): Map<number, number> => {
  for (const r of rows) {
    if (r.height > 0 && r.hasHeight !== false) heights.set(r.row, r.height);
  }
  return heights;
};

/**
 * Replace the active sheet's layout from engine-side overrides for `sheet`.
 * Called after loading a workbook and when switching sheets
 * so column widths / row heights stored in an .xlsx survive the round-trip.
 *
 * No-op when the engine doesn't expose `colRowSize` capability — the stub
 * fallback returns empty arrays from `getColumnLayouts` / `getRowLayouts`.
 *
 * Hidden flags and outline levels are also hydrated here since the same
 * engine vectors carry them; it would be wasteful to re-fetch separately.
 */
export function hydrateLayoutFromEngine(
  wb: WorkbookHandle,
  store: SpreadsheetStore,
  sheet: number,
): void {
  const cols = wb.getColumnLayouts(sheet);
  const rows = wb.getRowLayouts(sheet);
  const engineDefaults = engineAxisDefaults(wb, sheet);
  const defaults = engineDefaults ?? platformAxisDefaults(wb);
  const view = wb.getSheetView(sheet);
  const hasAxisLayouts =
    wb.capabilities.colRowSize || wb.capabilities.hiddenRowsCols || wb.capabilities.outlines;
  // hiddenSheets is workbook-scoped — walk every sheet so the set is correct
  // regardless of which sheet `hydrateLayoutFromEngine` is called for.
  const hiddenSheets = new Set<number>();
  const veryHiddenSheets = new Set<number>();
  if (wb.capabilities.sheetView) {
    const n = wb.sheetCount;
    for (let i = 0; i < n; i += 1) {
      const v = wb.getSheetView(i);
      if (!v) continue;
      // `visibility` is authoritative; an adapter that carries only the
      // two-state flag still resolves to hidden-or-visible.
      const visibility =
        v.visibility ?? (v.tabHidden ? SheetVisibility.Hidden : SheetVisibility.Visible);
      if (visibility === SheetVisibility.Visible) continue;
      hiddenSheets.add(i);
      // A very-hidden sheet stays out of the tab bar *and* out of the unhide
      // list, so the two states have to be tracked apart.
      if (visibility === SheetVisibility.VeryHidden) veryHiddenSheets.add(i);
    }
  }
  if (
    !hasAxisLayouts &&
    cols.length === 0 &&
    rows.length === 0 &&
    engineDefaults === null &&
    view === null &&
    hiddenSheets.size === 0
  ) {
    store.setState((s) => ({ ...s, layout: { ...s.layout, ...defaults } }));
    return;
  }

  store.setState((s) => {
    const colWidths = addColumnWidths(new Map(hasAxisLayouts ? [] : s.layout.colWidths), cols);
    const hiddenCols = new Set(hasAxisLayouts ? [] : s.layout.hiddenCols);
    const outlineCols = new Map(hasAxisLayouts ? [] : s.layout.outlineCols);
    for (const c of cols) {
      for (let col = c.first; col <= c.last; col += 1) {
        if (c.hidden) hiddenCols.add(col);
        if (c.outlineLevel > 0) outlineCols.set(col, c.outlineLevel);
      }
    }

    const rowHeights = addRowHeights(new Map(hasAxisLayouts ? [] : s.layout.rowHeights), rows);
    const hiddenRows = new Set(hasAxisLayouts ? [] : s.layout.hiddenRows);
    const outlineRows = new Map(hasAxisLayouts ? [] : s.layout.outlineRows);
    for (const r of rows) {
      if (r.hidden) hiddenRows.add(r.row);
      if (r.outlineLevel > 0) outlineRows.set(r.row, r.outlineLevel);
    }

    const layout = {
      ...s.layout,
      ...defaults,
      colWidths,
      rowHeights,
      hiddenCols,
      hiddenRows,
      outlineCols,
      outlineRows,
      hiddenSheets,
      veryHiddenSheets,
    };
    if (view) {
      layout.freezeRows = view.freezeRows;
      layout.freezeCols = view.freezeCols;
    }

    // Engine carries zoom as a percentage (10..400, default 100); the store
    // models it as a multiplier (1.0 = 100%) over the same range.
    // A view readback is authoritative even when the native value is the
    // default 100%. Hydration can follow a different sheet or a workbook
    // replacement, so retaining the previous store zoom at 100% leaks the
    // old sheet's visual scale into the new one.
    const viewport = view
      ? {
          ...s.viewport,
          zoom: Number.isFinite(view.zoomScale) ? clampZoom(view.zoomScale / 100) : 1,
        }
      : s.viewport;

    return {
      ...s,
      layout,
      viewport,
      ui: view
        ? {
            ...s.ui,
            // OOXML defaults these flags to visible. The nullish fallback
            // also keeps older/custom WorkbookHandle adapters compatible.
            showGridLines: view.showGridLines !== false,
            showHeaders: view.showRowColHeaders !== false,
            showZeros: view.showZeros !== false,
            // Unlike the three above, this one defaults to off.
            rightToLeft: view.rightToLeft === true,
            // Adapters without the field keep the sheet's remembered mode.
            workbookView: view.viewMode ?? s.ui.workbookView,
          }
        : s.ui,
    };
  });
}

/** Make `active` the sheet a saved file reopens on. */
export function syncActiveSheet(wb: WorkbookHandle, active: number): void {
  if (wb.capabilities.activeSheet && wb.activeSheet() !== active) wb.setActiveSheet(active);
}

/**
 * Re-read `sheet`'s default sizes and column/row size overrides into the store,
 * e.g. after `wb.geometryPlatform` changed the column model. Zoom, view flags,
 * freeze panes, hidden flags and outlines are left as they are.
 */
export function rehydrateAxisLayout(
  wb: WorkbookHandle,
  store: SpreadsheetStore,
  sheet: number,
): void {
  const defaults = engineAxisDefaults(wb, sheet) ?? platformAxisDefaults(wb);
  const sizes = wb.capabilities.colRowSize
    ? {
        colWidths: addColumnWidths(new Map(), wb.getColumnLayouts(sheet)),
        rowHeights: addRowHeights(new Map(), wb.getRowLayouts(sheet)),
      }
    : {};
  store.setState((s) => ({ ...s, layout: { ...s.layout, ...defaults, ...sizes } }));
}

/**
 * Push every layout-slice difference between `before` and `after` to the
 * engine for `sheet`. Drives both the forward apply and the undo/redo replay
 * of any layout mutation (col/row sizes, hidden flags, outline levels,
 * frozen panes). Each sub-sync is gated on its own capability flag so
 * partial engines (and the stub) only take the calls they support.
 *
 * Removed column and row entries use formulon's clear operations when
 * available, and retain the default-write fallback for older adapters.
 */
export function syncLayoutToEngine(
  wb: WorkbookHandle,
  layout: LayoutSlice,
  sheet: number,
  before: LayoutSnapshot,
  after: LayoutSnapshot,
): void {
  if (wb.capabilities.colRowSize) {
    const colKeys = new Set<number>();
    for (const k of before.colWidths.keys()) colKeys.add(k);
    for (const k of after.colWidths.keys()) colKeys.add(k);
    for (const col of colKeys) {
      const b = before.colWidths.get(col);
      const a = after.colWidths.get(col);
      if (b === a) continue;
      if (
        a === undefined &&
        wb.capabilities.columnWidthClear === true &&
        typeof wb.clearColumnWidth === 'function'
      ) {
        wb.clearColumnWidth(sheet, col, col);
      } else {
        wb.setColumnWidth(sheet, col, col, a ?? layout.defaultColWidth);
      }
    }

    const rowKeys = new Set<number>();
    for (const k of before.rowHeights.keys()) rowKeys.add(k);
    for (const k of after.rowHeights.keys()) rowKeys.add(k);
    for (const row of rowKeys) {
      const b = before.rowHeights.get(row);
      const a = after.rowHeights.get(row);
      if (b === a) continue;
      if (
        a === undefined &&
        wb.capabilities.rowHeightClear === true &&
        typeof wb.clearRowHeight === 'function'
      ) {
        wb.clearRowHeight(sheet, row);
      } else {
        wb.setRowHeight(sheet, row, a ?? layout.defaultRowHeight);
      }
    }
  }

  if (wb.capabilities.hiddenRowsCols) {
    const hcKeys = new Set<number>();
    for (const k of before.hiddenCols) hcKeys.add(k);
    for (const k of after.hiddenCols) hcKeys.add(k);
    for (const col of hcKeys) {
      const b = before.hiddenCols.has(col);
      const a = after.hiddenCols.has(col);
      if (b === a) continue;
      wb.setColumnHidden(sheet, col, col, a);
    }

    const hrKeys = new Set<number>();
    for (const k of before.hiddenRows) hrKeys.add(k);
    for (const k of after.hiddenRows) hrKeys.add(k);
    for (const row of hrKeys) {
      const b = before.hiddenRows.has(row);
      const a = after.hiddenRows.has(row);
      if (b === a) continue;
      wb.setRowHidden(sheet, row, a);
    }
  }

  if (wb.capabilities.outlines) {
    const ocKeys = new Set<number>();
    for (const k of before.outlineCols.keys()) ocKeys.add(k);
    for (const k of after.outlineCols.keys()) ocKeys.add(k);
    for (const col of ocKeys) {
      const b = before.outlineCols.get(col) ?? 0;
      const a = after.outlineCols.get(col) ?? 0;
      if (b === a) continue;
      wb.setColumnOutline(sheet, col, col, a);
    }

    const orKeys = new Set<number>();
    for (const k of before.outlineRows.keys()) orKeys.add(k);
    for (const k of after.outlineRows.keys()) orKeys.add(k);
    for (const row of orKeys) {
      const b = before.outlineRows.get(row) ?? 0;
      const a = after.outlineRows.get(row) ?? 0;
      if (b === a) continue;
      wb.setRowOutline(sheet, row, a);
    }
  }

  if (wb.capabilities.freeze) {
    if (before.freezeRows !== after.freezeRows || before.freezeCols !== after.freezeCols) {
      wb.setSheetFreeze(sheet, after.freezeRows, after.freezeCols);
    }
  }

  // hiddenSheets is workbook-scoped — diff doesn't depend on `sheet` param.
  if (wb.capabilities.sheetTabHidden) {
    const idxs = new Set<number>();
    for (const i of before.hiddenSheets) idxs.add(i);
    for (const i of after.hiddenSheets) idxs.add(i);
    for (const i of before.veryHiddenSheets) idxs.add(i);
    for (const i of after.veryHiddenSheets) idxs.add(i);
    for (const i of idxs) {
      const b = sheetVisibilityOf(before, i);
      const a = sheetVisibilityOf(after, i);
      if (b === a) continue;
      // `setSheetTabHidden` can neither state nor clear `veryHidden`: it
      // refuses to demote a very-hidden sheet and cannot promote a hidden one.
      // Route through the three-state setter whenever either side is that
      // state. On an engine that has only the flag, hiding is the closest
      // reachable state, and the refused demotion leaves the sheet as it was.
      const needsThreeState = a === SheetVisibility.VeryHidden || b === SheetVisibility.VeryHidden;
      const canStateVisibility = typeof wb.setSheetVisibility === 'function';
      if (needsThreeState && canStateVisibility && wb.setSheetVisibility(i, a)) continue;
      wb.setSheetTabHidden(i, a !== SheetVisibility.Visible);
    }
  }
}

const sheetVisibilityOf = (snapshot: LayoutSnapshot, sheet: number): SheetVisibility =>
  snapshot.veryHiddenSheets.has(sheet)
    ? SheetVisibility.VeryHidden
    : snapshot.hiddenSheets.has(sheet)
      ? SheetVisibility.Hidden
      : SheetVisibility.Visible;

/** Back-compat alias — the col/row size pointer-resize path still uses the
 *  narrow name. New code should use `syncLayoutToEngine`. */
export const syncLayoutSizesToEngine = syncLayoutToEngine;
