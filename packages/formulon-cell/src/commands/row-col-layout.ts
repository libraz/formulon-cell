import { MAX_COL, MAX_ROW } from '../engine/address.js';
import { syncLayoutToEngine } from '../engine/layout-sync.js';
import type { Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import {
  frozenColEnd,
  frozenColStart,
  frozenRowEnd,
  frozenRowStart,
} from '../store/freeze-band.js';
import { isRowHidden } from '../store/row-visibility.js';
import { type LayoutSlice, mutators, type SpreadsheetStore } from '../store/store.js';
import {
  type AutofitOptions,
  computeAutofitColWidth,
  computeAutofitRowHeight,
  createAutofitMeasureContext,
} from './autofit-measurement.js';
import type { History } from './history.js';
import { blockedByProtection } from './protection.js';
import { captureLayoutSnapshot, recordLayoutChangeWithEngine } from './slice-history.js';

const MAX_MATERIALIZED_LAYOUT_ROWS = 100_000;
const MAX_AUTOFIT_AXIS_SPAN_PAIRS = 100_000;

const spanSize = (start: number, end: number): number => (end >= start ? end - start + 1 : 0);

/** Mark rows [r0, r1] hidden. Wrapped in a single layout history entry.
 *  When `wb` is supplied the engine receives `setRowHidden` for each row so
 *  the change round-trips through .xlsx. */
export function hideRows(
  store: SpreadsheetStore,
  history: History | null,
  r0: number,
  r1: number,
  wb?: WorkbookHandle,
): void {
  const sheet = store.getState().data.sheetIndex;
  if (blockedByProtection(store, sheet, 'hideRows')) return;
  if (spanSize(r0, r1) > MAX_MATERIALIZED_LAYOUT_ROWS) return;
  recordLayoutChangeWithEngine(history, store, wb ?? null, () => {
    store.setState((s) => {
      const next = new Set(s.layout.hiddenRows);
      for (let r = r0; r <= r1; r += 1) next.add(r);
      return { ...s, layout: { ...s.layout, hiddenRows: next } };
    });
  });
}

export function showRows(
  store: SpreadsheetStore,
  history: History | null,
  r0: number,
  r1: number,
  wb?: WorkbookHandle,
): void {
  const sheet = store.getState().data.sheetIndex;
  if (blockedByProtection(store, sheet, 'showRows')) return;
  recordLayoutChangeWithEngine(history, store, wb ?? null, () => {
    store.setState((s) => {
      const next = new Set(s.layout.hiddenRows);
      for (const row of s.layout.hiddenRows) {
        if (row >= r0 && row <= r1) next.delete(row);
      }
      return { ...s, layout: { ...s.layout, ...showDefaultHiddenRows(s.layout, next, r0, r1) } };
    });
  });
}

/** The layout patch that shows `[r0, r1]`, given the explicitly hidden rows
 *  left once they are unhidden. On a sheet that hides rows by default,
 *  unhiding everything to the last row turns the default off, keeping hidden
 *  the default-hidden rows above `r0`; a span too large to list leaves the
 *  default-hidden rows as they are. */
function showDefaultHiddenRows(
  layout: LayoutSlice,
  hiddenRows: Set<number>,
  r0: number,
  r1: number,
): Pick<LayoutSlice, 'hiddenRows' | 'defaultRowsHidden' | 'shownRows'> {
  if (layout.defaultRowsHidden !== true) return { hiddenRows };
  const shownRows = new Set(layout.shownRows);
  if (r1 >= MAX_ROW && r0 <= MAX_MATERIALIZED_LAYOUT_ROWS) {
    for (let r = 0; r < r0; r += 1) if (!shownRows.has(r)) hiddenRows.add(r);
    return { hiddenRows, defaultRowsHidden: false, shownRows: new Set() };
  }
  if (spanSize(r0, r1) > MAX_MATERIALIZED_LAYOUT_ROWS) return { hiddenRows };
  for (let r = r0; r <= r1; r += 1) shownRows.add(r);
  return { hiddenRows, shownRows };
}

export function showRowsAroundSelection(
  store: SpreadsheetStore,
  history: History | null,
  r0: number,
  r1: number,
  wb?: WorkbookHandle,
): void {
  const layout = store.getState().layout;
  let start = r0;
  let end = r1;
  while (start > 0 && isRowHidden(layout, start - 1)) start -= 1;
  while (end < MAX_ROW && isRowHidden(layout, end + 1)) end += 1;
  showRows(store, history, start, end, wb);
}

export function hideCols(
  store: SpreadsheetStore,
  history: History | null,
  c0: number,
  c1: number,
  wb?: WorkbookHandle,
): void {
  const sheet = store.getState().data.sheetIndex;
  if (blockedByProtection(store, sheet, 'hideCols')) return;
  recordLayoutChangeWithEngine(history, store, wb ?? null, () => {
    store.setState((s) => {
      const next = new Set(s.layout.hiddenCols);
      for (let c = c0; c <= c1; c += 1) next.add(c);
      return { ...s, layout: { ...s.layout, hiddenCols: next } };
    });
  });
}

export function showCols(
  store: SpreadsheetStore,
  history: History | null,
  c0: number,
  c1: number,
  wb?: WorkbookHandle,
): void {
  const sheet = store.getState().data.sheetIndex;
  if (blockedByProtection(store, sheet, 'showCols')) return;
  recordLayoutChangeWithEngine(history, store, wb ?? null, () => {
    store.setState((s) => {
      const next = new Set(s.layout.hiddenCols);
      for (let c = c0; c <= c1; c += 1) next.delete(c);
      return { ...s, layout: { ...s.layout, hiddenCols: next } };
    });
  });
}

export function showColsAroundSelection(
  store: SpreadsheetStore,
  history: History | null,
  c0: number,
  c1: number,
  wb?: WorkbookHandle,
): void {
  const hidden = store.getState().layout.hiddenCols;
  let start = c0;
  let end = c1;
  while (start > 0 && hidden.has(start - 1)) start -= 1;
  while (end < MAX_COL && hidden.has(end + 1)) end += 1;
  showCols(store, history, start, end, wb);
}

export function setRowsHeight(
  store: SpreadsheetStore,
  history: History | null,
  r0: number,
  r1: number,
  px: number,
  wb?: WorkbookHandle,
): void {
  if (!Number.isFinite(px)) return;
  if (spanSize(r0, r1) > MAX_MATERIALIZED_LAYOUT_ROWS) return;
  recordLayoutChangeWithEngine(history, store, wb ?? null, () => {
    for (let row = r0; row <= r1; row += 1) mutators.setRowHeight(store, row, px);
  });
}

export function setColsWidth(
  store: SpreadsheetStore,
  history: History | null,
  c0: number,
  c1: number,
  px: number,
  wb?: WorkbookHandle,
): void {
  if (!Number.isFinite(px)) return;
  recordLayoutChangeWithEngine(history, store, wb ?? null, () => {
    for (let col = c0; col <= c1; col += 1) mutators.setColWidth(store, col, px);
  });
}

export function autofitRowsHeight(
  store: SpreadsheetStore,
  history: History | null,
  r0: number,
  r1: number,
  wb?: WorkbookHandle,
  opts?: AutofitOptions,
): void {
  if (spanSize(r0, r1) > MAX_MATERIALIZED_LAYOUT_ROWS) return;
  const measurementOptions = withWorkbookDisplayFont(opts, wb);
  recordLayoutChangeWithEngine(history, store, wb ?? null, () => {
    const ctx = createAutofitMeasureContext();
    for (let row = r0; row <= r1; row += 1) {
      mutators.setRowHeight(
        store,
        row,
        computeAutofitRowHeight(store.getState(), row, ctx, measurementOptions),
      );
    }
  });
}

export function autofitColsWidth(
  store: SpreadsheetStore,
  history: History | null,
  c0: number,
  c1: number,
  wb?: WorkbookHandle,
  opts?: AutofitOptions,
): void {
  const measurementOptions = withWorkbookDisplayFont(opts, wb);
  recordLayoutChangeWithEngine(history, store, wb ?? null, () => {
    const ctx = createAutofitMeasureContext();
    for (let col = c0; col <= c1; col += 1) {
      mutators.setColWidth(
        store,
        col,
        computeAutofitColWidth(store.getState(), col, ctx, measurementOptions),
      );
    }
  });
}

/**
 * AutoFit every row represented by a multi-range selection.
 *
 * A row can occur in more than one selected range.  Each occurrence keeps its
 * range's column span, and the largest measured height wins.  This is the
 * behaviour users expect from Ctrl/Cmd-disjoint row selections: cells in a
 * selected band affect that band only, while an overlapping row is still fit
 * for all of its selected cells.  All measurements happen before the layout
 * history mutation so a rejected materialization budget cannot leave a
 * partial result.
 */
export function autofitRowsHeightForRanges(
  store: SpreadsheetStore,
  history: History | null,
  ranges: readonly Range[],
  wb?: WorkbookHandle,
  opts?: AutofitOptions,
): void {
  const sheet = store.getState().data.sheetIndex;
  const normalized = normalizeAutofitRanges(ranges, sheet);
  if (!normalized || normalized.length === 0) return;
  // Count the raw axis/span planning work before allocating per-axis span
  // arrays. Overlap may collapse later, but accepting it here would still
  // allow a hostile selection to materialize unbounded duplicate entries.
  if (!withinAutofitPlanningBudget(normalized, 'row')) return;

  const spansByRow = new Map<number, Array<{ from: number; to: number }>>();
  for (const range of normalized) {
    for (let row = range.r0; row <= range.r1; row += 1) {
      const spans = spansByRow.get(row);
      if (spans) {
        spans.push({ from: range.c0, to: range.c1 });
      } else {
        spansByRow.set(row, [{ from: range.c0, to: range.c1 }]);
      }
      if (spansByRow.size > MAX_MATERIALIZED_LAYOUT_ROWS) return;
    }
  }
  const mergedSpansByRow = mergeAutofitSpans(spansByRow);
  if (countAutofitSpanPairs(mergedSpansByRow) > MAX_AUTOFIT_AXIS_SPAN_PAIRS) return;

  const state = store.getState();
  const measurementOptions = withWorkbookDisplayFont(opts, wb);
  const ctx = createAutofitMeasureContext();
  const heights = new Map<number, number>();
  for (const [row, spans] of mergedSpansByRow) {
    let max = state.layout.defaultRowHeight;
    for (const span of spans) {
      max = Math.max(
        max,
        computeAutofitRowHeight(state, row, ctx, {
          ...measurementOptions,
          span,
        }),
      );
    }
    heights.set(row, max);
  }

  recordLayoutChangeWithEngine(history, store, wb ?? null, () => {
    for (const [row, height] of heights) mutators.setRowHeight(store, row, height);
  });
}

/**
 * AutoFit every column represented by a multi-range selection.
 *
 * The cross-axis row span is retained for each selected range.  If the same
 * column is selected more than once, its width is the maximum over those
 * independent measurements.  The entire change is recorded as one layout
 * history entry so one Undo restores every selected column together.
 */
export function autofitColsWidthForRanges(
  store: SpreadsheetStore,
  history: History | null,
  ranges: readonly Range[],
  wb?: WorkbookHandle,
  opts?: AutofitOptions,
): void {
  const sheet = store.getState().data.sheetIndex;
  const normalized = normalizeAutofitRanges(ranges, sheet);
  if (!normalized || normalized.length === 0) return;
  if (!withinAutofitPlanningBudget(normalized, 'col')) return;

  const spansByCol = new Map<number, Array<{ from: number; to: number }>>();
  for (const range of normalized) {
    for (let col = range.c0; col <= range.c1; col += 1) {
      const spans = spansByCol.get(col);
      if (spans) {
        spans.push({ from: range.r0, to: range.r1 });
      } else {
        spansByCol.set(col, [{ from: range.r0, to: range.r1 }]);
      }
    }
  }
  const mergedSpansByCol = mergeAutofitSpans(spansByCol);
  if (countAutofitSpanPairs(mergedSpansByCol) > MAX_AUTOFIT_AXIS_SPAN_PAIRS) return;

  const state = store.getState();
  const measurementOptions = withWorkbookDisplayFont(opts, wb);
  const ctx = createAutofitMeasureContext();
  const widths = new Map<number, number>();
  for (const [col, spans] of mergedSpansByCol) {
    let max = 0;
    for (const span of spans) {
      max = Math.max(
        max,
        computeAutofitColWidth(state, col, ctx, {
          ...measurementOptions,
          span,
        }),
      );
    }
    widths.set(col, max);
  }

  recordLayoutChangeWithEngine(history, store, wb ?? null, () => {
    for (const [col, width] of widths) mutators.setColWidth(store, col, width);
  });
}

function normalizeAutofitRanges(ranges: readonly Range[], sheet: number): Range[] | null {
  const normalized: Range[] = [];
  const seen = new Set<string>();
  for (const range of ranges) {
    // A selection containing a stale/foreign range is invalid as a whole.
    // Silently fitting only the valid subset would make one Ribbon command
    // produce a partial layout update after a sheet switch.
    if (range.sheet !== sheet) return null;
    const coordinates = [range.r0, range.r1, range.c0, range.c1];
    if (!coordinates.every(Number.isInteger)) return null;
    const r0 = Math.min(range.r0, range.r1);
    const r1 = Math.max(range.r0, range.r1);
    const c0 = Math.min(range.c0, range.c1);
    const c1 = Math.max(range.c0, range.c1);
    if (r0 < 0 || r1 > MAX_ROW || c0 < 0 || c1 > MAX_COL) return null;
    const key = `${sheet}:${r0}:${c0}:${r1}:${c1}`;
    if (seen.has(key)) continue;
    seen.add(key);
    normalized.push({ sheet, r0, c0, r1, c1 });
  }
  return normalized;
}

/**
 * Conservative preflight budget: overlapping ranges count independently.
 * This bounds planning allocations before their spans can be merged.
 */
function withinAutofitPlanningBudget(ranges: readonly Range[], axis: 'row' | 'col'): boolean {
  let planned = 0;
  for (const range of ranges) {
    const axisSpan = axis === 'row' ? spanSize(range.r0, range.r1) : spanSize(range.c0, range.c1);
    if (axisSpan > MAX_AUTOFIT_AXIS_SPAN_PAIRS - planned) return false;
    planned += axisSpan;
  }
  return true;
}

function mergeAutofitSpans(
  spansByAxis: Map<number, Array<{ from: number; to: number }>>,
): Map<number, Array<{ from: number; to: number }>> {
  const merged = new Map<number, Array<{ from: number; to: number }>>();
  for (const [axis, spans] of spansByAxis) {
    const ordered = [...spans].sort((a, b) => a.from - b.from || a.to - b.to);
    const result: Array<{ from: number; to: number }> = [];
    for (const span of ordered) {
      const previous = result.at(-1);
      if (previous && span.from <= previous.to + 1) {
        previous.to = Math.max(previous.to, span.to);
      } else {
        result.push({ ...span });
      }
    }
    merged.set(axis, result);
  }
  return merged;
}

const countAutofitSpanPairs = (
  spansByAxis: Map<number, Array<{ from: number; to: number }>>,
): number => {
  let count = 0;
  for (const spans of spansByAxis.values()) count += spans.length;
  return count;
};

function withWorkbookDisplayFont(
  options: AutofitOptions | undefined,
  wb: WorkbookHandle | undefined,
): AutofitOptions | undefined {
  if (!wb) return options;
  return {
    ...(options ?? {}),
    workbook: wb,
    ...(options?.workbookDisplayFont !== undefined
      ? {}
      : { workbookDisplayFont: wb.workbookDisplayFont }),
  };
}

/** Resolve which row/col indices to show again from the current selection.
 *  Spreadsheets return visible rows that flank a hidden band; we emulate by
 *  reporting every hidden row inside the selection. */
export function hiddenInSelection(
  layout: LayoutSlice,
  axis: 'row' | 'col',
  a: number,
  b: number,
): number[] {
  const lo = Math.min(a, b);
  const hi = Math.max(a, b);
  const out: number[] = [];
  if (axis === 'row' && layout.defaultRowsHidden === true) {
    for (let r = lo; r <= hi; r += 1) if (isRowHidden(layout, r)) out.push(r);
    return out;
  }
  const set = axis === 'row' ? layout.hiddenRows : layout.hiddenCols;
  for (const index of set) {
    if (index >= lo && index <= hi) out.push(index);
  }
  return out.sort((left, right) => left - right);
}

/** First and last hidden row or column of the selection, without listing
 *  the ones between; null when none is hidden. */
export function hiddenExtentInSelection(
  layout: LayoutSlice,
  axis: 'row' | 'col',
  a: number,
  b: number,
): { first: number; last: number } | null {
  const lo = Math.min(a, b);
  const hi = Math.max(a, b);
  if (axis === 'row' && layout.defaultRowsHidden === true) {
    // Visible rows are the listed ones, so each walk stops within them.
    let first = lo;
    while (first <= hi && !isRowHidden(layout, first)) first += 1;
    if (first > hi) return null;
    let last = hi;
    while (!isRowHidden(layout, last)) last -= 1;
    return { first, last };
  }
  const set = axis === 'row' ? layout.hiddenRows : layout.hiddenCols;
  let first = Number.POSITIVE_INFINITY;
  let last = Number.NEGATIVE_INFINITY;
  for (const index of set) {
    if (index < lo || index > hi) continue;
    first = Math.min(first, index);
    last = Math.max(last, index);
  }
  return first <= last ? { first, last } : null;
}

/** Pin `rows` rows / `cols` cols whose first row/column is `origin` (default
 *  A1). Freezing is a view setting, not an undo step. When `wb` is supplied
 *  the counts are pushed to the engine so they round-trip through .xlsx.
 *  `split` makes the band split panes instead of a freeze. */
export function setFreezePanes(
  store: SpreadsheetStore,
  rows: number,
  cols: number,
  wb?: WorkbookHandle,
  origin?: { row: number; col: number },
  split = false,
): void {
  const sheet = store.getState().data.sheetIndex;
  const before = captureLayoutSnapshot(store.getState());
  mutators.setFreezePanes(store, rows, cols, origin, split);
  if (!wb) return;
  const after = captureLayoutSnapshot(store.getState());
  syncLayoutToEngine(wb, store.getState().layout, sheet, before, after);
}

export type FreezeMode = 'none' | 'topRow' | 'firstColumn' | 'panes';

/** Where Freeze Panes and Split divide the window: above and left of the
 *  active cell, or through the middle of the window when the active cell is
 *  its top-left. Moves the window to the band's origin. */
export function panesAtActiveCell(store: SpreadsheetStore): {
  rows: number;
  cols: number;
  origin: { row: number; col: number };
} {
  const { layout, viewport, selection } = store.getState();
  let originRow = frozenRowEnd(layout) > 0 ? frozenRowStart(layout) : viewport.rowStart;
  let originCol = frozenColEnd(layout) > 0 ? frozenColStart(layout) : viewport.colStart;
  const a = selection.active;
  // An active cell that is scrolled out of view becomes the top-left first.
  if (a.row < originRow || a.col < originCol) {
    originRow = a.row;
    originCol = a.col;
  }
  const middle = a.row === originRow && a.col === originCol;
  const rows = middle ? Math.max(1, Math.floor(viewport.rowCount / 2)) : a.row - originRow;
  const cols = middle ? Math.max(1, Math.floor(viewport.colCount / 2)) : a.col - originCol;
  // The band starts where the window starts, so the window moves with it.
  store.setState((s) => ({
    ...s,
    viewport: { ...s.viewport, rowStart: originRow, colStart: originCol },
  }));
  return { rows, cols, origin: { row: originRow, col: originCol } };
}

/** Freeze (or unfreeze) the way the desktop spreadsheet does: the band starts
 *  at the top-left visible cell, and "Freeze Panes" splits at the active
 *  cell, or at the middle of the window when the active cell is the top-left.
 *  Freeze Panes on split panes freezes them where they are split. */
export function applyFreezeMode(
  store: SpreadsheetStore,
  mode: FreezeMode,
  wb?: WorkbookHandle,
): void {
  if (mode === 'none') {
    setFreezePanes(store, 0, 0, wb);
    return;
  }
  const { layout, viewport } = store.getState();
  if (mode === 'panes') {
    if (layout.splitPanes === true) {
      const origin = { row: frozenRowStart(layout), col: frozenColStart(layout) };
      setFreezePanes(store, layout.freezeRows, layout.freezeCols, wb, origin);
      return;
    }
    const { rows, cols, origin } = panesAtActiveCell(store);
    setFreezePanes(store, rows, cols, wb, origin);
    return;
  }
  const originRow = frozenRowEnd(layout) > 0 ? frozenRowStart(layout) : viewport.rowStart;
  const originCol = frozenColEnd(layout) > 0 ? frozenColStart(layout) : viewport.colStart;
  const rows = mode === 'firstColumn' ? 0 : 1;
  const cols = mode === 'firstColumn' ? 1 : 0;
  setFreezePanes(store, rows, cols, wb, { row: originRow, col: originCol });
}
