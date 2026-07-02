import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { LayoutSlice, SpreadsheetStore } from '../store/store.js';
import { type History, recordLayoutChangeWithEngine } from './history.js';
import { isSheetProtected } from './protection.js';

/** Spreadsheets cap outline depth at 7 — beyond that the gutter would be unreadable. */
export const MAX_OUTLINE_LEVEL = 7;

/** Width of one bracket slot in CSS pixels. The gutter widens by this amount
 *  per outline level so each level has its own visual track. */
export const OUTLINE_GUTTER_PER_LEVEL = 14;
const MAX_MATERIALIZED_OUTLINE_ROWS = 100_000;

const spanSize = (start: number, end: number): number => (end >= start ? end - start + 1 : 0);

const recomputeRowGutter = (levels: Map<number, number>): number => {
  let max = 0;
  for (const v of levels.values()) if (v > max) max = v;
  return max * OUTLINE_GUTTER_PER_LEVEL;
};

const recomputeColGutter = (levels: Map<number, number>): number => {
  let max = 0;
  for (const v of levels.values()) if (v > max) max = v;
  return max * OUTLINE_GUTTER_PER_LEVEL;
};

const setRowOutline = (store: SpreadsheetStore, next: Map<number, number>): void => {
  store.setState((s) => ({
    ...s,
    layout: {
      ...s.layout,
      outlineRows: next,
      outlineRowGutter: recomputeRowGutter(next),
    },
  }));
};

const setColOutline = (store: SpreadsheetStore, next: Map<number, number>): void => {
  store.setState((s) => ({
    ...s,
    layout: {
      ...s.layout,
      outlineCols: next,
      outlineColGutter: recomputeColGutter(next),
    },
  }));
};

const blockedByProtection = (store: SpreadsheetStore, op: string): boolean => {
  const sheet = store.getState().data.sheetIndex;
  if (!isSheetProtected(store.getState(), sheet)) return false;
  // eslint-disable-next-line no-console
  console.warn(`formulon-cell: ${op} blocked — sheet ${sheet} is protected`);
  return true;
};

/** Increase outline level by 1 for rows in `[r0, r1]`. Caps at level 7. */
export function groupRows(
  store: SpreadsheetStore,
  history: History | null,
  r0: number,
  r1: number,
  wb?: WorkbookHandle,
): void {
  if (r0 > r1) return;
  if (blockedByProtection(store, 'groupRows')) return;
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
  if (blockedByProtection(store, 'ungroupRows')) return;
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
  if (blockedByProtection(store, 'groupCols')) return;
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
  if (blockedByProtection(store, 'ungroupCols')) return;
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
  if (blockedByProtection(store, 'collapseRowGroup')) return;
  if (spanSize(r0, r1) > MAX_MATERIALIZED_OUTLINE_ROWS) return;
  recordLayoutChangeWithEngine(history, store, wb ?? null, () => {
    store.setState((s) => {
      const next = new Set(s.layout.hiddenRows);
      for (let r = r0; r <= r1; r += 1) next.add(r);
      return { ...s, layout: { ...s.layout, hiddenRows: next } };
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
  if (blockedByProtection(store, 'expandRowGroup')) return;
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

export function collapseColGroup(
  store: SpreadsheetStore,
  history: History | null,
  c0: number,
  c1: number,
  wb?: WorkbookHandle,
): void {
  if (blockedByProtection(store, 'collapseColGroup')) return;
  recordLayoutChangeWithEngine(history, store, wb ?? null, () => {
    store.setState((s) => {
      const next = new Set(s.layout.hiddenCols);
      for (let c = c0; c <= c1; c += 1) next.add(c);
      return { ...s, layout: { ...s.layout, hiddenCols: next } };
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
  if (blockedByProtection(store, 'expandColGroup')) return;
  recordLayoutChangeWithEngine(history, store, wb ?? null, () => {
    store.setState((s) => {
      const next = new Set(s.layout.hiddenCols);
      for (let c = c0; c <= c1; c += 1) next.delete(c);
      return { ...s, layout: { ...s.layout, hiddenCols: next } };
    });
  });
}

/** True when at least one row in `[r0, r1]` is hidden — used to decide which
 *  glyph (+ vs −) to paint on the toggle. */
export function isRowGroupCollapsed(layout: LayoutSlice, r0: number, r1: number): boolean {
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
