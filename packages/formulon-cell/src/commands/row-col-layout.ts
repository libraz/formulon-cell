import { MAX_COL, MAX_ROW } from '../engine/address.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { type LayoutSlice, mutators, type SpreadsheetStore } from '../store/store.js';
import {
  type AutofitOptions,
  computeAutofitColWidth,
  computeAutofitRowHeight,
  createAutofitMeasureContext,
} from './autofit-measurement.js';
import type { History } from './history.js';
import { blockedByProtection } from './protection.js';
import { recordLayoutChangeWithEngine } from './slice-history.js';

const MAX_MATERIALIZED_LAYOUT_ROWS = 100_000;

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
      return { ...s, layout: { ...s.layout, hiddenRows: next } };
    });
  });
}

export function showRowsAroundSelection(
  store: SpreadsheetStore,
  history: History | null,
  r0: number,
  r1: number,
  wb?: WorkbookHandle,
): void {
  const hidden = store.getState().layout.hiddenRows;
  let start = r0;
  let end = r1;
  while (start > 0 && hidden.has(start - 1)) start -= 1;
  while (end < MAX_ROW && hidden.has(end + 1)) end += 1;
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
  recordLayoutChangeWithEngine(history, store, wb ?? null, () => {
    const ctx = createAutofitMeasureContext();
    for (let row = r0; row <= r1; row += 1) {
      mutators.setRowHeight(store, row, computeAutofitRowHeight(store.getState(), row, ctx, opts));
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
  recordLayoutChangeWithEngine(history, store, wb ?? null, () => {
    const ctx = createAutofitMeasureContext();
    for (let col = c0; col <= c1; col += 1) {
      mutators.setColWidth(store, col, computeAutofitColWidth(store.getState(), col, ctx, opts));
    }
  });
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
  const set = axis === 'row' ? layout.hiddenRows : layout.hiddenCols;
  const out: number[] = [];
  for (const index of set) {
    if (index >= lo && index <= hi) out.push(index);
  }
  return out.sort((left, right) => left - right);
}

/** Pin `rows` rows / `cols` cols. One Cmd+Z reverts the freeze change.
 *  Pass `null` for `history` to skip recording. When `wb` is supplied and
 *  `capabilities.freeze` is on, the change is also pushed to the engine so
 *  it round-trips through .xlsx save/load. Store + engine writes share one
 *  history entry so undo/redo moves both sides in lockstep. */
export function setFreezePanes(
  store: SpreadsheetStore,
  history: History | null,
  rows: number,
  cols: number,
  wb?: WorkbookHandle,
): void {
  recordLayoutChangeWithEngine(history, store, wb ?? null, () => {
    mutators.setFreezePanes(store, rows, cols);
  });
}
