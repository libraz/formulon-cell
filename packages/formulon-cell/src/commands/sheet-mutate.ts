import { hydrateSheetProtectionFromEngine } from '../engine/protection-sync.js';
import { hydrateSheetPrForSheet } from '../engine/sheet-pr-sync.js';
import { hydrateTableOverlaysFromEngine } from '../engine/table-sync.js';
import { SheetVisibility } from '../engine/types.js';
import { nearestVisibleSheet } from '../engine/view-state-sync.js';
import {
  hasWholeSheetFontStoreBinding,
  remapStoreSheetStateForStructureEvent,
} from '../engine/whole-sheet-font-sync.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { withSparklineMembers } from '../store/sparkline-slice.js';
import { type LayoutSlice, mutators, type SpreadsheetStore } from '../store/store.js';
import type { OutlineSummary } from '../store/types.js';
import { isMacChrome } from './chrome-platform.js';
import type { History } from './history.js';
import { isWorkbookStructureProtected } from './protection.js';
import { recordLayoutChangeWithEngine, recordSparklineChange } from './slice-history.js';
import { findSlicerTable } from './slicers.js';
import { renameSheetInSparklineSources } from './sparkline.js';

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
  wb?: WorkbookHandle,
): void => {
  const boundStore = wb ? hasWholeSheetFontStoreBinding(wb, store) : false;
  // A mounted WorkbookHandle has already performed this remap from its
  // structure event. Command-only callers still need the address-keyed store
  // state moved alongside the native sheets, but must not apply the mapping a
  // second time after the event path has handled it.
  if (!boundStore) remapStoreSheetStateForStructureEvent(store, mapIndex);
  store.setState((s) => {
    const hiddenSheets = new Set<number>();
    for (const idx of s.layout.hiddenSheets) {
      const next = mapIndex(idx);
      if (next !== null) hiddenSheets.add(next);
    }
    const veryHiddenSheets = new Set<number>();
    for (const idx of s.layout.veryHiddenSheets) {
      const next = mapIndex(idx);
      if (next !== null) veryHiddenSheets.add(next);
    }
    const sheetTabColors = new Map<number, string>();
    for (const [idx, color] of s.layout.sheetTabColors) {
      const next = mapIndex(idx);
      if (next !== null) sheetTabColors.set(next, color);
    }
    const outlineSummary = new Map<number, OutlineSummary>();
    for (const [idx, summary] of s.layout.outlineSummary ?? []) {
      const next = mapIndex(idx);
      if (next !== null) outlineSummary.set(next, summary);
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
      layout: {
        ...s.layout,
        hiddenSheets,
        veryHiddenSheets,
        sheetTabColors,
        ...(s.layout.outlineSummary ? { outlineSummary } : {}),
      },
      protection: { ...s.protection, protectedSheets, allowedEditRanges },
    };
  });
};

/** Recalculate through the workbook's policy, then reload the active sheet's
 *  cells: a structural sheet change can alter values on sheets that were not
 *  written to. */
const recalcAndRefresh = (store: SpreadsheetStore | null | undefined, wb: WorkbookHandle): void => {
  wb.recalcAuto();
  if (store) mutators.refreshActiveSheetCells(store, wb);
};

/** Per mounted store, how a newly active sheet is loaded from the engine. */
const sheetActivationHydrators = new WeakMap<SpreadsheetStore, () => void>();

/** Register how the mounted instance loads its active sheet (cells, layout,
 *  merges, validation and the rest), so a structure change activates a sheet
 *  the way switching tabs does. */
export function registerSheetActivationHydrator(
  store: SpreadsheetStore,
  hydrate: () => void,
): () => void {
  sheetActivationHydrators.set(store, hydrate);
  return () => {
    if (sheetActivationHydrators.get(store) === hydrate) sheetActivationHydrators.delete(store);
  };
}

/** Make `idx` the active sheet and load it. Without a mounted instance only
 *  its cells are loaded. */
const activateSheet = (
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  idx: number,
  options?: Parameters<typeof mutators.setSheetIndex>[2],
): void => {
  mutators.setSheetIndex(store, idx, options);
  const hydrate = sheetActivationHydrators.get(store);
  if (hydrate) hydrate();
  else mutators.replaceCells(store, wb.cells(idx));
};

/** Table overlays name their sheet; re-read them after the sheets were reordered. */
const refreshTableOverlays = (store: SpreadsheetStore, wb: WorkbookHandle): void => {
  hydrateTableOverlaysFromEngine(wb, store);
};

/** Slicers name a table, which goes away with its sheet. */
const pruneOrphanSlicers = (store: SpreadsheetStore, wb: WorkbookHandle): void => {
  for (const spec of store.getState().slicers.slicers) {
    if (!findSlicerTable(wb, spec.tableName)) mutators.removeSlicer(store, spec.id);
  }
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

/** Highest `SheetN` number handed out per workbook. Excel keeps this counter
 *  rising for the session: a deleted sheet's number is not reused. The first
 *  add takes the lowest free number. */
const sheetNumbers = new WeakMap<WorkbookHandle, number>();

/** The next default sheet number: the counter's next value that no sheet
 *  already uses as `SheetN`, ignoring case. */
function nextSheetNumber(wb: WorkbookHandle): number {
  const taken = new Set<string>();
  for (let i = 0; i < wb.sheetCount; i += 1) taken.add(wb.sheetName(i).toLocaleLowerCase());
  let n = (sheetNumbers.get(wb) ?? 0) + 1;
  while (taken.has(`sheet${n}`)) n += 1;
  return n;
}

export function addSheet(
  store: SpreadsheetStore | null | undefined,
  wb: WorkbookHandle,
  history?: History | null,
): number {
  if (!structureAllowed(store, 'add sheet')) return -1;
  const beforeActive = store?.getState().data.sheetIndex ?? 0;
  const number = nextSheetNumber(wb);
  const added = wb.addSheet(`Sheet${number}`);
  if (added < 0) return added;
  sheetNumbers.set(wb, number);
  const name = 'sheetName' in wb ? wb.sheetName(added) : undefined;
  if (store && history && !history.isReplaying()) {
    history.push({
      undo: () => {
        const active = store.getState().data.sheetIndex;
        if (active === added || active >= wb.sheetCount) {
          activateSheet(store, wb, Math.min(beforeActive, Math.max(0, wb.sheetCount - 2)));
        }
        if (wb.removeSheet(added)) {
          remapWorkbookSheetLayout(store, (sheet) => shiftedAfterRemove(sheet, added), wb);
        }
      },
      redo: () => {
        const next = wb.addSheet(name);
        if (next >= 0) activateSheet(store, wb, next);
      },
    });
  }
  return added;
}

/** Add a sheet directly after the active one and return its index (-1 on
 *  refusal). Undo removes the sheet in one step. */
export function addSheetAfterActive(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  history?: History | null,
): number {
  const active = store.getState().data.sheetIndex;
  history?.begin();
  try {
    const added = addSheet(store, wb, history);
    const target = active + 1;
    if (added < 0 || target >= added) return added;
    return moveSheet(store, wb, added, target, history) ? target : added;
  } finally {
    history?.end();
  }
}

/** Longest sheet name the desktop spreadsheet accepts. */
const MAX_SHEET_NAME_LENGTH = 31;

/** Whether `name` is a legal name for the sheet at `idx`: 1-31 characters,
 *  none of \\ / ? * [ ] :, no leading or trailing apostrophe, and unique among
 *  the other sheets ignoring case. Off the Mac (`mac` false), `History` is reserved in any case,
 *  as the Windows desktop app reserves it; the Mac app accepts it. */
export function isValidSheetName(
  wb: WorkbookHandle,
  idx: number,
  name: string,
  mac = false,
): boolean {
  if (name.length === 0 || name.length > MAX_SHEET_NAME_LENGTH) return false;
  if (!mac && name.toLowerCase() === 'history') return false;
  if (/[\\/?*[\]:]/.test(name) || name.startsWith("'") || name.endsWith("'")) return false;
  const existing = wb.sheetIndexByName(name);
  return existing === null || existing === idx;
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
  if (!isValidSheetName(wb, idx, name, store ? isMacChrome(store) : false)) return false;
  const beforeName = 'sheetName' in wb ? wb.sheetName(idx) : null;
  const ok = wb.renameSheet(idx, name);
  if (!ok) return false;
  recalcAndRefresh(store, wb);
  history?.begin();
  try {
    if (store && beforeName) {
      // Sparkline sources name their sheet, so they follow the rename.
      recordSparklineChange(history ?? null, store, () => {
        store.setState((s) => {
          const sparklines = renameSheetInSparklineSources(
            s.sparkline.sparklines,
            beforeName,
            name,
          );
          return sparklines
            ? { ...s, sparkline: withSparklineMembers(s.sparkline, sparklines) }
            : s;
        });
      });
    }
    if (beforeName && history && !history.isReplaying()) {
      history.push({
        undo: () => {
          wb.renameSheet(idx, beforeName);
          recalcAndRefresh(store, wb);
        },
        redo: () => {
          wb.renameSheet(idx, name);
          recalcAndRefresh(store, wb);
        },
      });
    }
  } finally {
    history?.end();
  }
  return true;
}

/** Remove the sheet at `idx`. When the active sheet is affected, the store's
 *  `sheetIndex` is corrected to point at a still-valid neighbor:
 *  - removing an index *before* the active: active shifts down by 1
 *  - removing the active itself: the next visible sheet to the right, else the
 *    nearest visible one on the left
 *  - removing an index *after* the active: no change
 *  The selection is reset to A1 on the new active sheet. Returns false when
 *  the removal is rejected (e.g. trying to remove the last sheet) or when
 *  the engine lacks `sheetMutate`. */
export function removeSheet(store: SpreadsheetStore, wb: WorkbookHandle, idx: number): boolean {
  if (!structureAllowed(store, 'remove sheet')) return false;
  if (wb.sheetCount <= 1) return false;
  const ok = wb.removeSheet(idx);
  if (!ok) return false;
  remapWorkbookSheetLayout(store, (sheet) => shiftedAfterRemove(sheet, idx), wb);
  pruneOrphanSlicers(store, wb);
  refreshTableOverlays(store, wb);

  const cur = store.getState().data.sheetIndex;
  if (idx < cur) {
    mutators.renumberActiveSheet(store, cur - 1);
  } else if (idx === cur) {
    const next = nearestVisibleSheet(wb.sheetCount, store.getState().layout.hiddenSheets, idx);
    activateSheet(store, wb, next < 0 ? cur : next, { discardCurrent: true });
  }
  recalcAndRefresh(store, wb);
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
  remapWorkbookSheetLayout(store, (sheet) => shiftedAfterMove(sheet, from, to), wb);
  refreshTableOverlays(store, wb);

  const cur = store.getState().data.sheetIndex;
  mutators.renumberActiveSheet(store, shiftedAfterMove(cur, from, to));
  const afterActive = store.getState().data.sheetIndex;
  if (history && !history.isReplaying()) {
    history.push({
      undo: () => {
        if (!wb.moveSheet(to, from)) return;
        remapWorkbookSheetLayout(store, (sheet) => shiftedAfterMove(sheet, to, from), wb);
        refreshTableOverlays(store, wb);
        mutators.renumberActiveSheet(store, beforeActive);
      },
      redo: () => {
        if (!wb.moveSheet(from, to)) return;
        remapWorkbookSheetLayout(store, (sheet) => shiftedAfterMove(sheet, from, to), wb);
        refreshTableOverlays(store, wb);
        mutators.renumberActiveSheet(store, afterActive);
      },
    });
  }
  return true;
}

/** Read the copy's sheet-level state from the engine, which copied it with the
 *  sheet: protection and its allowed edit ranges (under fresh ids), tab colour
 *  and outline summary. The store mirrors the engine; nothing is duplicated. */
const hydrateCopiedSheet = (store: SpreadsheetStore, wb: WorkbookHandle, copy: number): void => {
  const hydrate = (): void => {
    hydrateSheetProtectionFromEngine(wb, store, copy);
    hydrateSheetPrForSheet(wb, store, copy);
  };
  wb.withEngineSyncMuted(hydrate);
};

/** Copy the sheet at `source` so the copy lands before the sheet at `before`
 *  (the sheet count appends) and make it the active sheet. The engine names
 *  the copy (`Sheet1 (2)`). Returns the copy's index, or -1 when refused. One
 *  undo step removes the copy and restores the active sheet; redo re-creates
 *  it under the same name. */
export function copySheet(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  source: number,
  before: number,
  history?: History | null,
): number {
  if (!structureAllowed(store, 'copy sheet')) return -1;
  const beforeActive = store.getState().data.sheetIndex;
  const insert = (name?: string): number => {
    const copy = wb.copySheet(source, before, name);
    if (copy < 0) return copy;
    remapWorkbookSheetLayout(store, (sheet) => (sheet >= copy ? sheet + 1 : sheet), wb);
    hydrateCopiedSheet(store, wb, copy);
    refreshTableOverlays(store, wb);
    wb.recalcAuto();
    activateSheet(store, wb, copy);
    return copy;
  };
  const copy = insert();
  if (copy < 0) return copy;
  const name = wb.sheetName(copy);
  if (history && !history.isReplaying()) {
    history.push({
      undo: () => {
        if (!wb.removeSheet(copy)) return;
        remapWorkbookSheetLayout(store, (sheet) => shiftedAfterRemove(sheet, copy), wb);
        refreshTableOverlays(store, wb);
        wb.recalcAuto();
        activateSheet(store, wb, beforeActive, { discardCurrent: true });
      },
      redo: () => {
        insert(name);
      },
    });
  }
  return copy;
}

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
    const next = nearestVisibleSheet(n, after.layout.hiddenSheets, idx);
    if (next >= 0) mutators.setSheetIndex(store, next);
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
