import type { Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { outlineGutterWidth, outlineSummaryFor } from '../store/outline-summary.js';
import { isRowHidden } from '../store/row-visibility.js';
import { type LayoutSlice, mutators, type SpreadsheetStore } from '../store/store.js';
import type { OutlineSummary } from '../store/types.js';
import type { History } from './history.js';
import { blockedByProtection } from './protection.js';
import { recordLayoutChangeWithEngine } from './slice-history.js';

/** Spreadsheets cap outline depth at 7 — beyond that the gutter would be unreadable. */
export const MAX_OUTLINE_LEVEL = 7;

export { OUTLINE_GUTTER_PER_LEVEL } from '../store/outline-summary.js';

const MAX_MATERIALIZED_OUTLINE_ROWS = 100_000;

const spanSize = (start: number, end: number): number => (end >= start ? end - start + 1 : 0);

/** The row or column carrying a group's `collapsed` flag: the summary just
 *  past the detail on the summary side, or none at the sheet edge. */
const summaryIndex = (before: boolean, i0: number, i1: number): number | null =>
  before ? (i0 > 0 ? i0 - 1 : null) : i1 + 1;

const withFlag = (
  set: ReadonlySet<number> | undefined,
  index: number | null,
  on: boolean,
): Set<number> => {
  const next = new Set(set);
  if (index === null) return next;
  if (on) next.add(index);
  else next.delete(index);
  return next;
};

const setRowOutline = (store: SpreadsheetStore, next: Map<number, number>): void => {
  store.setState((s) => ({
    ...s,
    layout: {
      ...s.layout,
      outlineRows: next,
      outlineRowGutter: outlineGutterWidth(next),
    },
  }));
};

const setColOutline = (store: SpreadsheetStore, next: Map<number, number>): void => {
  store.setState((s) => ({
    ...s,
    layout: {
      ...s.layout,
      outlineCols: next,
      outlineColGutter: outlineGutterWidth(next),
    },
  }));
};

/** Place the active sheet's outline summary rows below / above and summary
 *  columns right / left of their detail. */
export function setOutlineSummaryPosition(
  store: SpreadsheetStore,
  patch: Partial<OutlineSummary>,
): void {
  const sheet = store.getState().data.sheetIndex;
  if (blockedByProtection(store, sheet, 'setOutlineSummaryPosition')) return;
  mutators.setOutlineSummary(store, sheet, patch);
}

/** Increase outline level by 1 for rows in `[r0, r1]`. Caps at level 7. */
export function groupRows(
  store: SpreadsheetStore,
  history: History | null,
  r0: number,
  r1: number,
  wb?: WorkbookHandle,
): void {
  if (r0 > r1) return;
  if (blockedByProtection(store, store.getState().data.sheetIndex, 'groupRows')) return;
  if (spanSize(r0, r1) > MAX_MATERIALIZED_OUTLINE_ROWS) return;
  recordLayoutChangeWithEngine(history, store, wb ?? null, () => {
    const cur = store.getState().layout.outlineRows;
    const next = new Map(cur);
    for (let r = r0; r <= r1; r += 1) {
      const lvl = next.get(r) ?? 0;
      if (lvl < MAX_OUTLINE_LEVEL) next.set(r, lvl + 1);
    }
    setRowOutline(store, next);
  });
}

/** Decrease outline level by 1 for rows in `[r0, r1]`. Removes entries that
 *  fall to 0. */
export function ungroupRows(
  store: SpreadsheetStore,
  history: History | null,
  r0: number,
  r1: number,
  wb?: WorkbookHandle,
): void {
  if (r0 > r1) return;
  if (blockedByProtection(store, store.getState().data.sheetIndex, 'ungroupRows')) return;
  recordLayoutChangeWithEngine(history, store, wb ?? null, () => {
    const cur = store.getState().layout.outlineRows;
    const next = new Map(cur);
    for (const [r, lvl] of cur) {
      if (r < r0 || r > r1) continue;
      if (lvl <= 1) next.delete(r);
      else next.set(r, lvl - 1);
    }
    setRowOutline(store, next);
  });
}

export function groupCols(
  store: SpreadsheetStore,
  history: History | null,
  c0: number,
  c1: number,
  wb?: WorkbookHandle,
): void {
  if (c0 > c1) return;
  if (blockedByProtection(store, store.getState().data.sheetIndex, 'groupCols')) return;
  recordLayoutChangeWithEngine(history, store, wb ?? null, () => {
    const cur = store.getState().layout.outlineCols;
    const next = new Map(cur);
    for (let c = c0; c <= c1; c += 1) {
      const lvl = next.get(c) ?? 0;
      if (lvl < MAX_OUTLINE_LEVEL) next.set(c, lvl + 1);
    }
    setColOutline(store, next);
  });
}

export function ungroupCols(
  store: SpreadsheetStore,
  history: History | null,
  c0: number,
  c1: number,
  wb?: WorkbookHandle,
): void {
  if (c0 > c1) return;
  if (blockedByProtection(store, store.getState().data.sheetIndex, 'ungroupCols')) return;
  recordLayoutChangeWithEngine(history, store, wb ?? null, () => {
    const cur = store.getState().layout.outlineCols;
    const next = new Map(cur);
    for (let c = c0; c <= c1; c += 1) {
      const lvl = next.get(c) ?? 0;
      if (lvl <= 1) next.delete(c);
      else next.set(c, lvl - 1);
    }
    setColOutline(store, next);
  });
}

/** Walk outwards from `row` to find the contiguous run of rows whose outline
 *  level is ≥ `level`. Used to translate a click on a +/- toggle into the
 *  full set of rows the toggle controls. Returns null if `row` itself isn't
 *  at level ≥ `level`. */
export function rowGroupRangeAt(
  layout: LayoutSlice,
  row: number,
  level: number,
): { r0: number; r1: number } | null {
  const lvl = layout.outlineRows.get(row) ?? 0;
  if (lvl < level) return null;
  let r0 = row;
  let r1 = row;
  while ((layout.outlineRows.get(r0 - 1) ?? 0) >= level) r0 -= 1;
  while ((layout.outlineRows.get(r1 + 1) ?? 0) >= level) r1 += 1;
  return { r0, r1 };
}

export function colGroupRangeAt(
  layout: LayoutSlice,
  col: number,
  level: number,
): { c0: number; c1: number } | null {
  const lvl = layout.outlineCols.get(col) ?? 0;
  if (lvl < level) return null;
  let c0 = col;
  let c1 = col;
  while ((layout.outlineCols.get(c0 - 1) ?? 0) >= level) c0 -= 1;
  while ((layout.outlineCols.get(c1 + 1) ?? 0) >= level) c1 += 1;
  return { c0, c1 };
}

/** Hide every row in `[r0, r1]`. Wrapped in a layout history entry. */
export function collapseRowGroup(
  store: SpreadsheetStore,
  history: History | null,
  r0: number,
  r1: number,
  wb?: WorkbookHandle,
): void {
  if (blockedByProtection(store, store.getState().data.sheetIndex, 'collapseRowGroup')) return;
  if (spanSize(r0, r1) > MAX_MATERIALIZED_OUTLINE_ROWS) return;
  recordLayoutChangeWithEngine(history, store, wb ?? null, () => {
    store.setState((s) => {
      const next = new Set(s.layout.hiddenRows);
      for (let r = r0; r <= r1; r += 1) next.add(r);
      const summary = summaryIndex(!outlineSummaryFor(s.layout, s.data.sheetIndex).below, r0, r1);
      const collapsedRows = withFlag(s.layout.collapsedRows, summary, true);
      return { ...s, layout: { ...s.layout, hiddenRows: next, collapsedRows } };
    });
  });
}

export function expandRowGroup(
  store: SpreadsheetStore,
  history: History | null,
  r0: number,
  r1: number,
  wb?: WorkbookHandle,
): void {
  if (blockedByProtection(store, store.getState().data.sheetIndex, 'expandRowGroup')) return;
  recordLayoutChangeWithEngine(history, store, wb ?? null, () => {
    store.setState((s) => {
      const next = new Set(s.layout.hiddenRows);
      for (const row of s.layout.hiddenRows) {
        if (row >= r0 && row <= r1) next.delete(row);
      }
      const summary = summaryIndex(!outlineSummaryFor(s.layout, s.data.sheetIndex).below, r0, r1);
      const collapsedRows = withFlag(s.layout.collapsedRows, summary, false);
      return { ...s, layout: { ...s.layout, hiddenRows: next, collapsedRows } };
    });
  });
}

export function collapseColGroup(
  store: SpreadsheetStore,
  history: History | null,
  c0: number,
  c1: number,
  wb?: WorkbookHandle,
): void {
  if (blockedByProtection(store, store.getState().data.sheetIndex, 'collapseColGroup')) return;
  recordLayoutChangeWithEngine(history, store, wb ?? null, () => {
    store.setState((s) => {
      const next = new Set(s.layout.hiddenCols);
      for (let c = c0; c <= c1; c += 1) next.add(c);
      const summary = summaryIndex(!outlineSummaryFor(s.layout, s.data.sheetIndex).right, c0, c1);
      const collapsedCols = withFlag(s.layout.collapsedCols, summary, true);
      return { ...s, layout: { ...s.layout, hiddenCols: next, collapsedCols } };
    });
  });
}

export function expandColGroup(
  store: SpreadsheetStore,
  history: History | null,
  c0: number,
  c1: number,
  wb?: WorkbookHandle,
): void {
  if (blockedByProtection(store, store.getState().data.sheetIndex, 'expandColGroup')) return;
  recordLayoutChangeWithEngine(history, store, wb ?? null, () => {
    store.setState((s) => {
      const next = new Set(s.layout.hiddenCols);
      for (let c = c0; c <= c1; c += 1) next.delete(c);
      const summary = summaryIndex(!outlineSummaryFor(s.layout, s.data.sheetIndex).right, c0, c1);
      const collapsedCols = withFlag(s.layout.collapsedCols, summary, false);
      return { ...s, layout: { ...s.layout, hiddenCols: next, collapsedCols } };
    });
  });
}

/** True when at least one row in `[r0, r1]` is hidden — used to decide which
 *  glyph (+ vs −) to paint on the toggle. */
export function isRowGroupCollapsed(layout: LayoutSlice, r0: number, r1: number): boolean {
  if (layout.defaultRowsHidden === true) {
    for (let row = r0; row <= r1; row += 1) if (isRowHidden(layout, row)) return true;
    return false;
  }
  for (const row of layout.hiddenRows) {
    if (row >= r0 && row <= r1) return true;
  }
  return false;
}

export function isColGroupCollapsed(layout: LayoutSlice, c0: number, c1: number): boolean {
  for (const col of layout.hiddenCols) {
    if (col >= c0 && col <= c1) return true;
  }
  return false;
}

/** The group Hide Detail / Show Detail acts on: the innermost one holding the
 *  whole selection, or the one a selected summary row/column summarises.
 *  Null when the selection is outside every group. */
function detailGroupAt(
  levels: ReadonlyMap<number, number>,
  summaryBefore: boolean,
  i0: number,
  i1: number,
  runAt: (index: number, level: number) => { i0: number; i1: number } | null,
): { i0: number; i1: number } | null {
  let level = Number.POSITIVE_INFINITY;
  for (let i = i0; i <= Math.min(i1, i0 + MAX_MATERIALIZED_OUTLINE_ROWS); i += 1) {
    level = Math.min(level, levels.get(i) ?? 0);
  }
  if (i0 === i1) {
    const detail = summaryBefore ? i0 + 1 : i0 - 1;
    if ((levels.get(detail) ?? 0) > level) return runAt(detail, level + 1);
  }
  return level > 0 ? runAt(i0, level) : null;
}

/** Rows for a selection spanning rows, columns for one spanning columns; a
 *  single cell takes its row group, else its column group. */
const outlineDetailTarget = (
  store: SpreadsheetStore,
  range: Range,
): { axis: 'row' | 'col'; i0: number; i1: number } | null => {
  const { layout, data } = store.getState();
  const summary = outlineSummaryFor(layout, data.sheetIndex);
  const rows = () =>
    detailGroupAt(layout.outlineRows, !summary.below, range.r0, range.r1, (i, lvl) => {
      const r = rowGroupRangeAt(layout, i, lvl);
      return r && { i0: r.r0, i1: r.r1 };
    });
  const cols = () =>
    detailGroupAt(layout.outlineCols, !summary.right, range.c0, range.c1, (i, lvl) => {
      const c = colGroupRangeAt(layout, i, lvl);
      return c && { i0: c.c0, i1: c.c1 };
    });
  const row = range.r1 > range.r0 || range.c0 === range.c1 ? rows() : null;
  if (row) return { axis: 'row', ...row };
  const col = range.r1 > range.r0 ? null : cols();
  return col && { axis: 'col', ...col };
};

/** Data > Hide Detail: collapse the group around the selection. */
export function hideOutlineDetail(
  store: SpreadsheetStore,
  history: History | null,
  range: Range,
  wb?: WorkbookHandle,
): void {
  const target = outlineDetailTarget(store, range);
  if (!target) return;
  if (target.axis === 'row') collapseRowGroup(store, history, target.i0, target.i1, wb);
  else collapseColGroup(store, history, target.i0, target.i1, wb);
}

/** Data > Show Detail: expand the group around the selection. */
export function showOutlineDetail(
  store: SpreadsheetStore,
  history: History | null,
  range: Range,
  wb?: WorkbookHandle,
): void {
  const target = outlineDetailTarget(store, range);
  if (!target) return;
  if (target.axis === 'row') expandRowGroup(store, history, target.i0, target.i1, wb);
  else expandColGroup(store, history, target.i0, target.i1, wb);
}
