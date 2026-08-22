import { SheetVisibility } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { type LayoutSlice, mutators, type SpreadsheetStore } from '../store/store.js';
import { type History, recordLayoutChangeWithEngine } from './history.js';
import { isWorkbookStructureProtected } from './protection.js';

const shiftedAfterRemove = (idx: number, removed: number): number | null => {
  if (idx === removed) return null;
  return idx > removed ? idx - 1 : idx;
};

const shiftedAfterMove = (idx: number, from: number, to: number): number => {
  if (idx === from) return to;
  if (from < to && idx > from && idx <= to) return idx - 1;
  if (from > to && idx >= to && idx < from) return idx + 1;
  return idx;
};

const remapWorkbookSheetLayout = (
  store: SpreadsheetStore,
  mapIndex: (idx: number) => number | null,
): void => {
  store.setState((s) => {
    const hiddenSheets = new Set<number>();
    for (const idx of s.layout.hiddenSheets) {
      const next = mapIndex(idx);
      if (next !== null) hiddenSheets.add(next);
    }
    const sheetTabColors = new Map<number, string>();
    for (const [idx, color] of s.layout.sheetTabColors) {
      const next = mapIndex(idx);
      if (next !== null) sheetTabColors.set(next, color);
    }
    const protectedSheets = new Map<number, { password?: string }>();
    for (const [idx, protection] of s.protection.protectedSheets) {
      const next = mapIndex(idx);
      if (next !== null) protectedSheets.set(next, protection);
    }
    const allowedEditRanges = s.protection.allowedEditRanges.flatMap((entry) => {
      const next = mapIndex(entry.range.sheet);
      return next === null ? [] : [{ ...entry, range: { ...entry.range, sheet: next } }];
    });
    return {
      ...s,
      layout: { ...s.layout, hiddenSheets, sheetTabColors },
      protection: { ...s.protection, protectedSheets, allowedEditRanges },
    };
  });
};

const warnWorkbookStructureProtected = (op: string): void => {
  // eslint-disable-next-line no-console
  console.warn(`formulon-cell: ${op} blocked — workbook structure is protected`);
};

function structureAllowed(store: SpreadsheetStore | null | undefined, op: string): boolean {
  if (!store || !isWorkbookStructureProtected(store.getState())) return true;
  warnWorkbookStructureProtected(op);
  return false;
}

export function addSheet(
  store: SpreadsheetStore | null | undefined,
  wb: WorkbookHandle,
  history?: History | null,
): number {
  if (!structureAllowed(store, 'add sheet')) return -1;
  const beforeActive = store?.getState().data.sheetIndex ?? 0;
  const added = wb.addSheet();
  if (added < 0) return added;
  const name = 'sheetName' in wb ? wb.sheetName(added) : undefined;
  if (store && history && !history.isReplaying()) {
    history.push({
      undo: () => {
        const active = store.getState().data.sheetIndex;
        if (active === added || active >= wb.sheetCount) {
          mutators.setSheetIndex(store, Math.min(beforeActive, Math.max(0, wb.sheetCount - 2)));
        }
        if (wb.removeSheet(added)) {
          remapWorkbookSheetLayout(store, (sheet) => shiftedAfterRemove(sheet, added));
        }
      },
      redo: () => {
        const next = wb.addSheet(name);
        if (next >= 0) mutators.setSheetIndex(store, next);
      },
    });
  }
  return added;
}

/** Rename the sheet at `idx`. Returns true on success. The store has no
 *  per-sheet name slot — the playground reads names from `wb.sheetName` —
 *  so this is a pass-through to the engine. */
export function renameSheet(
  wb: WorkbookHandle,
  idx: number,
  name: string,
  store?: SpreadsheetStore | null,
  history?: History | null,
): boolean {
  if (!structureAllowed(store, 'rename sheet')) return false;
  const beforeName = 'sheetName' in wb ? wb.sheetName(idx) : null;
  const ok = wb.renameSheet(idx, name);
  if (!ok) return false;
  if (beforeName && history && !history.isReplaying()) {
    history.push({
      undo: () => {
        wb.renameSheet(idx, beforeName);
      },
      redo: () => {
        wb.renameSheet(idx, name);
      },
    });
  }
  return true;
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
  if (!structureAllowed(store, 'remove sheet')) return false;
  if (wb.sheetCount <= 1) return false;
  const ok = wb.removeSheet(idx);
  if (!ok) return false;
  remapWorkbookSheetLayout(store, (sheet) => shiftedAfterRemove(sheet, idx));

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
  history?: History | null,
): boolean {
  if (!structureAllowed(store, 'move sheet')) return false;
  if (from === to) return true;
  const beforeActive = store.getState().data.sheetIndex;
  const ok = wb.moveSheet(from, to);
  if (!ok) return false;
  remapWorkbookSheetLayout(store, (sheet) => shiftedAfterMove(sheet, from, to));

  const cur = store.getState().data.sheetIndex;
  let next = cur;
  if (cur === from) next = to;
  else if (from < cur && to >= cur) next = cur - 1;
  else if (from > cur && to <= cur) next = cur + 1;
  if (next !== cur) {
    mutators.setSheetIndex(store, next);
  }
  const afterActive = store.getState().data.sheetIndex;
  if (history && !history.isReplaying()) {
    history.push({
      undo: () => {
        if (!wb.moveSheet(to, from)) return;
        remapWorkbookSheetLayout(store, (sheet) => shiftedAfterMove(sheet, to, from));
        mutators.setSheetIndex(store, beforeActive);
      },
      redo: () => {
        if (!wb.moveSheet(from, to)) return;
        remapWorkbookSheetLayout(store, (sheet) => shiftedAfterMove(sheet, from, to));
        mutators.setSheetIndex(store, afterActive);
      },
    });
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

/**
 * Move `idx`'s tab to one of the three visibility states. Refuses to hide the
 * last visible sheet (spreadsheet parity — a workbook with no visible sheets is
 * invalid) and refuses a state the tab is already in. When the active sheet
 * becomes hidden, the active index advances to the next visible sheet. The
 * mutation goes through `recordLayoutChangeWithEngine`, so it round-trips
 * through engine save and is undoable.
 *
 * `veryHidden` differs from `hidden` only in reach: the tab is left out of the
 * Unhide list as well as the tab bar, which is how a workbook keeps a settings
 * or lookup sheet away from a user. Both states keep the sheet's data.
 */
export function setSheetTabVisibility(
  store: SpreadsheetStore,
  wb: WorkbookHandle | null,
  history: History | null,
  idx: number,
  visibility: SheetVisibility,
): boolean {
  const toVisible = visibility === SheetVisibility.Visible;
  if (!structureAllowed(store, toVisible ? 'unhide sheet' : 'hide sheet')) return false;
  const state = store.getState();
  const n = wb ? wb.sheetCount : 1;
  const hiddenNow = new Set(state.layout.hiddenSheets);
  const veryHiddenNow = new Set(state.layout.veryHiddenSheets);
  const current = veryHiddenNow.has(idx)
    ? SheetVisibility.VeryHidden
    : hiddenNow.has(idx)
      ? SheetVisibility.Hidden
      : SheetVisibility.Visible;
  if (current === visibility) return false;
  if (!toVisible && current === SheetVisibility.Visible) {
    // Refuse if this would hide the last visible sheet.
    let visibleCount = 0;
    for (let i = 0; i < n; i += 1) {
      if (i === idx) continue;
      if (!hiddenNow.has(i)) visibleCount += 1;
    }
    if (visibleCount === 0) return false;
  }

  recordLayoutChangeWithEngine(history, store, wb, () => {
    const layout = store.getState().layout;
    const nextHidden = new Set(layout.hiddenSheets);
    const nextVeryHidden = new Set(layout.veryHiddenSheets);
    if (toVisible) {
      nextHidden.delete(idx);
      nextVeryHidden.delete(idx);
    } else {
      // Every very-hidden sheet is also hidden, so the two sets stay nested.
      nextHidden.add(idx);
      if (visibility === SheetVisibility.VeryHidden) nextVeryHidden.add(idx);
      else nextVeryHidden.delete(idx);
    }
    store.setState((s) => ({
      ...s,
      layout: {
        ...s.layout,
        hiddenSheets: nextHidden,
        veryHiddenSheets: nextVeryHidden,
      } as LayoutSlice,
    }));
  });

  // After hiding the active sheet, hop to the next visible one.
  const after = store.getState();
  if (!toVisible && after.data.sheetIndex === idx) {
    mutators.setSheetIndex(store, firstVisibleSheet(n, after.layout.hiddenSheets, idx));
  }
  return true;
}

/** Two-state view of `setSheetTabVisibility`. `false` reveals a tab from either
 *  hidden state; `true` on an already very-hidden tab is refused, since it
 *  states nothing that tab does not already satisfy. */
export function setSheetHidden(
  store: SpreadsheetStore,
  wb: WorkbookHandle | null,
  history: History | null,
  idx: number,
  hidden: boolean,
): boolean {
  // Hiding an already very-hidden tab would demote it, which is a state change
  // this two-state form never asked for. The engine's own flag setter refuses
  // the same way.
  if (hidden && store.getState().layout.hiddenSheets.has(idx)) return false;
  return setSheetTabVisibility(
    store,
    wb,
    history,
    idx,
    hidden ? SheetVisibility.Hidden : SheetVisibility.Visible,
  );
}
