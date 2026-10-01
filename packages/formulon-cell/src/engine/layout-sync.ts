import type { LayoutSnapshot } from '../commands/history.js';
import type { LayoutSlice, SpreadsheetStore } from '../store/store.js';
import { SheetVisibility } from './types.js';
import type { WorkbookHandle } from './workbook-handle.js';

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
    view === null &&
    hiddenSheets.size === 0
  ) {
    return;
  }

  store.setState((s) => {
    const colWidths = new Map(hasAxisLayouts ? [] : s.layout.colWidths);
    const hiddenCols = new Set(hasAxisLayouts ? [] : s.layout.hiddenCols);
    const outlineCols = new Map(hasAxisLayouts ? [] : s.layout.outlineCols);
    for (const c of cols) {
      for (let col = c.first; col <= c.last; col += 1) {
        if (c.width > 0) colWidths.set(col, c.width);
        if (c.hidden) hiddenCols.add(col);
        if (c.outlineLevel > 0) outlineCols.set(col, c.outlineLevel);
      }
    }

    const rowHeights = new Map(hasAxisLayouts ? [] : s.layout.rowHeights);
    const hiddenRows = new Set(hasAxisLayouts ? [] : s.layout.hiddenRows);
    const outlineRows = new Map(hasAxisLayouts ? [] : s.layout.outlineRows);
    for (const r of rows) {
      if (r.height > 0) rowHeights.set(r.row, r.height);
      if (r.hidden) hiddenRows.add(r.row);
      if (r.outlineLevel > 0) outlineRows.set(r.row, r.outlineLevel);
    }

    const layout = {
      ...s.layout,
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
    // models it as a multiplier (1.0 = 100%) clamped to [0.5, 4].
    const viewport =
      view && view.zoomScale !== 100
        ? { ...s.viewport, zoom: Math.max(0.5, Math.min(4, view.zoomScale / 100)) }
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
          }
        : s.ui,
    };
  });
}

/**
 * Push every layout-slice difference between `before` and `after` to the
 * engine for `sheet`. Drives both the forward apply and the undo/redo replay
 * of any layout mutation (col/row sizes, hidden flags, outline levels,
 * frozen panes). Each sub-sync is gated on its own capability flag so
 * partial engines (and the stub) only take the calls they support.
 *
 * Removed size entries (present in `before`, missing from `after`) are
 * written as the layout default — the engine has no "clear override" call,
 * but writing the default is visually equivalent.
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
      wb.setColumnWidth(sheet, col, col, a ?? layout.defaultColWidth);
    }

    const rowKeys = new Set<number>();
    for (const k of before.rowHeights.keys()) rowKeys.add(k);
    for (const k of after.rowHeights.keys()) rowKeys.add(k);
    for (const row of rowKeys) {
      const b = before.rowHeights.get(row);
      const a = after.rowHeights.get(row);
      if (b === a) continue;
      wb.setRowHeight(sheet, row, a ?? layout.defaultRowHeight);
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
