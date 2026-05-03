import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { type LayoutSlice, type SpreadsheetStore, mutators } from '../store/store.js';
import { type History, recordLayoutChangeWithEngine } from './history.js';

/** Rename the sheet at `idx`. Returns true on success. The store has no
 *  per-sheet name slot — the playground reads names from `wb.sheetName` —
 *  so this is a pass-through to the engine. */
export function renameSheet(wb: WorkbookHandle, idx: number, name: string): boolean {
  return wb.renameSheet(idx, name);
}

/** Remove the sheet at `idx`. When the active sheet is affected, the store's
 *  `sheetIndex` is corrected to point at a still-valid neighbor:
 *  - removing an index *before* the active: active shifts down by 1
 *  - removing the active itself: select max(idx - 1, 0)
 *  - removing an index *after* the active: no change
 *  The selection is reset to A1 on the new active sheet. Returns false when
 *  the removal is rejected (e.g. trying to remove the last sheet) or when
 *  the engine lacks `sheetMutate`. */
export function removeSheet(store: SpreadsheetStore, wb: WorkbookHandle, idx: number): boolean {
  if (wb.sheetCount <= 1) return false;
  const ok = wb.removeSheet(idx);
  if (!ok) return false;

  const cur = store.getState().data.sheetIndex;
  let next = cur;
  if (idx < cur) next = cur - 1;
  else if (idx === cur) next = Math.max(idx - 1, 0);
  // idx > cur: no change.
  if (next !== cur) {
    mutators.setSheetIndex(store, next);
  }
  return true;
}

/** Move the sheet from `from` to `to` (post-removal index). Translates the
 *  active sheet index along with the move when needed. Returns false when
 *  the engine lacks `sheetMutate` or the move is rejected. */
export function moveSheet(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  from: number,
  to: number,
): boolean {
  if (from === to) return true;
  const ok = wb.moveSheet(from, to);
  if (!ok) return false;

  const cur = store.getState().data.sheetIndex;
  let next = cur;
  if (cur === from) next = to;
  else if (from < cur && to >= cur) next = cur - 1;
  else if (from > cur && to <= cur) next = cur + 1;
  if (next !== cur) {
    mutators.setSheetIndex(store, next);
  }
  return true;
}

const firstVisibleSheet = (n: number, hidden: ReadonlySet<number>, skip: number): number => {
  for (let i = 0; i < n; i += 1) {
    if (i === skip) continue;
    if (!hidden.has(i)) return i;
  }
  return 0;
};

/** Toggle the tab-hidden flag on `idx`. Refuses to hide the last visible
 *  sheet (Excel parity — leaving a workbook with no visible sheets is
 *  invalid). When the active sheet becomes hidden, the active index advances
 *  to the next visible sheet. The mutation goes through
 *  `recordLayoutChangeWithEngine`, so it round-trips through engine save and
 *  is undoable. */
export function setSheetHidden(
  store: SpreadsheetStore,
  wb: WorkbookHandle | null,
  history: History | null,
  idx: number,
  hidden: boolean,
): boolean {
  const state = store.getState();
  const n = wb ? wb.sheetCount : 1;
  const cur = new Set(state.layout.hiddenSheets);
  if (hidden) {
    if (cur.has(idx)) return false;
    // Refuse if this would hide the last visible sheet.
    let visibleCount = 0;
    for (let i = 0; i < n; i += 1) {
      if (i === idx) continue;
      if (!cur.has(i)) visibleCount += 1;
    }
    if (visibleCount === 0) return false;
  } else if (!cur.has(idx)) return false;

  recordLayoutChangeWithEngine(history, store, wb, () => {
    const layout = store.getState().layout;
    const next = new Set(layout.hiddenSheets);
    if (hidden) next.add(idx);
    else next.delete(idx);
    store.setState((s) => ({
      ...s,
      layout: { ...s.layout, hiddenSheets: next } as LayoutSlice,
    }));
  });

  // After hiding the active sheet, hop to the next visible one.
  const after = store.getState();
  if (hidden && after.data.sheetIndex === idx) {
    mutators.setSheetIndex(store, firstVisibleSheet(n, after.layout.hiddenSheets, idx));
  }
  return true;
}
