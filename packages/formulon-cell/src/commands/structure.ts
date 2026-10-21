import { AnchorEditAs } from '@libraz/formulon';
import { addrKey, MAX_COL, MAX_ROW, parseAddrKey } from '../engine/address.js';
import {
  type ActiveAxisFormatSnapshot,
  captureActiveAxisFormatOverrideForStructure,
  remapActiveAxisFormatOverridesForStructure,
  restoreActiveAxisFormatOverrideForStructure,
  restoreNativeAxisSnapshot,
} from '../engine/axis-format-sync.js';
import {
  type AxisStructureShift,
  mapAxisAddress,
  remapAxisFormatOverride,
  remapSheetFontOverride,
} from '../engine/axis-structure-remap.js';
import {
  captureSyncedFormatKeysForStructure,
  remapSyncedFormatKeysForStructure,
  restoreSyncedFormatKeysForStructure,
} from '../engine/cell-format-sync.js';
import { hydrateChartsFromEngine } from '../engine/charts-sync.js';
import {
  type FormatHistoryTrackingSnapshot,
  formatHistoryBridgeFor,
  type NativeFormatSnapshot,
} from '../engine/format-history-bridge.js';
import type {
  Addr,
  CellValue,
  Range,
  SheetAxisFormatOverride,
  SheetFontOverride,
} from '../engine/types.js';
import { writeCell } from '../engine/value.js';
import {
  type ActiveWholeSheetFontSnapshot,
  captureActiveWholeSheetFontOverrideForStructure,
  remapActiveWholeSheetFontOverridesForStructure,
  restoreActiveWholeSheetFontOverrideForStructure,
} from '../engine/whole-sheet-font-sync.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { EngineDrawingObject } from '../engine/workbook-handle-annotations.js';
import { addMergeToMaps } from '../store/merge-maps.js';
import {
  type ConditionalRule,
  type LayoutSlice,
  mutators,
  type SpreadsheetStore,
  type State,
} from '../store/store.js';
import {
  inheritFormatsByCol,
  inheritFormatsByRow,
  normalizeAxisEdit,
  shiftFilterCriteria,
  shiftFormatsByCol,
  shiftFormatsByRow,
  shiftIndexedMap,
  shiftIndexedMapWithInheritance,
  shiftIndexedSet,
  shiftRangeAxis,
} from './axis-shift.js';
import { recordFilterChange } from './filter.js';
import { collectAllFormulas, type FormulaRecord } from './formula-records.js';
import { adjustFormulaForRowColEdit } from './formula-refs.js';
import type { History } from './history.js';
import { blockedByProtection } from './protection.js';
import {
  captureLayoutSnapshot,
  recordConditionalRulesChange,
  recordFormatChange,
  recordLayoutChange,
  recordMergesChange,
  recordMergesChangeWithEngine,
} from './slice-history.js';

interface CellRecord {
  addr: Addr;
  value: CellValue;
  formula: string | null;
}

function collectAllCells(wb: WorkbookHandle, sheet: number): CellRecord[] {
  const out: CellRecord[] = [];
  for (const c of wb.cells(sheet)) {
    out.push({ addr: c.addr, value: c.value, formula: c.formula });
  }
  return out;
}

/** Reject an insertion that would move persisted content past the worksheet
 *  boundary. The native engine and the JS fallback must make the same choice;
 *  otherwise the fallback silently drops the tail cell while the native path
 *  reports a different workbook shape. Sparse formatting/layout and merges
 *  are included because shifting either one out of bounds would also lose
 *  workbook state. */
function insertionWouldOverflow(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  sheet: number,
  axis: 'row' | 'col',
  split: number,
  count: number,
): boolean {
  const max = axis === 'row' ? MAX_ROW : MAX_COL;
  const lastMovable = max - count;
  const indexOf = (addr: Addr): number => (axis === 'row' ? addr.row : addr.col);

  for (const cell of wb.cells(sheet)) {
    if (indexOf(cell.addr) <= lastMovable) continue;
    if (cell.formula !== null || cell.value.kind !== 'blank') return true;
  }

  const state = store.getState();
  for (const key of state.format.formats.keys()) {
    const addr = parseAddrKey(key);
    if (addr && addr.sheet === sheet && indexOf(addr) > lastMovable) return true;
  }

  const indexedMaps =
    axis === 'row'
      ? [state.layout.rowHeights, state.layout.hiddenRows, state.layout.outlineRows]
      : [state.layout.colWidths, state.layout.hiddenCols, state.layout.outlineCols];
  for (const indexed of indexedMaps) {
    for (const index of indexed.keys()) {
      if (index > lastMovable) return true;
    }
  }

  for (const merge of state.merges.byAnchor.values()) {
    if (merge.sheet !== sheet) continue;
    const end = axis === 'row' ? merge.r1 : merge.c1;
    if (end > lastMovable && end >= split) return true;
  }
  return false;
}

/** Apply a row/col shift to all cells on `sheet`. Cells in the band are
 *  relocated; formulas everywhere are rewritten so refs follow the move.
 *  When `delta < 0`, cells in the deletion band are dropped and refs into
 *  the band are replaced with `#REF!`. */
function applyAxisShiftToCells(
  wb: WorkbookHandle,
  sheet: number,
  axis: 'row' | 'col',
  split: number,
  delta: number,
): void {
  if (delta === 0) return;
  const all = collectAllCells(wb, sheet);
  wb.withBatchedRecalc(() => writeAxisShiftedCells(wb, all, axis, split, delta));
}

function writeAxisShiftedCells(
  wb: WorkbookHandle,
  all: readonly CellRecord[],
  axis: 'row' | 'col',
  split: number,
  delta: number,
): void {
  const max = axis === 'row' ? MAX_ROW : MAX_COL;
  // Blank every cell that needs to move (or be deleted) before re-writing —
  // some target slots may overlap source slots when delta < count.
  for (const c of all) {
    const k = axis === 'row' ? c.addr.row : c.addr.col;
    if (k >= split) wb.setBlank(c.addr);
  }

  for (const c of all) {
    const k = axis === 'row' ? c.addr.row : c.addr.col;
    const inMovedBand = k >= split;
    const inDeletedBand = delta < 0 && inMovedBand && k < split - delta;
    if (inDeletedBand) continue; // dropped

    const newAddr: Addr = inMovedBand
      ? axis === 'row'
        ? { ...c.addr, row: k + delta }
        : { ...c.addr, col: k + delta }
      : c.addr;

    const newFormula = c.formula ? adjustFormulaForRowColEdit(c.formula, axis, split, delta) : null;

    if (inMovedBand) {
      // Cells in the moved band were blanked above; rewrite at new addr.
      const target = k + delta;
      if (target < 0 || target > max) continue;
      writeCell(wb, newAddr, c.value, newFormula);
    } else if (newFormula !== c.formula && newFormula !== null) {
      // Stationary cell whose formula references the band — overwrite in place.
      // (writeCell would also work but `setFormula` is direct.)
      wb.setFormula(c.addr, newFormula);
    }
  }
}

/** Move the frozen band with the cells: rows or columns inserted above/left of
 *  it push it along, ones inserted inside grow it, deleted ones shrink it. */
function shiftFreezeBand(
  store: SpreadsheetStore,
  history: History | null,
  axis: 'row' | 'col',
  kind: 'insert' | 'delete',
  split: number,
  n: number,
): void {
  const layout = store.getState().layout;
  const row = axis === 'row';
  const count = row ? layout.freezeRows : layout.freezeCols;
  if (count === 0) return;
  const origin = (row ? layout.freezeOriginRow : layout.freezeOriginCol) ?? 0;
  const end = origin + count;
  const edge = (e: number): number =>
    kind === 'insert' ? (split < e ? e + n : e) : e <= split ? e : e >= split + n ? e - n : split;
  const nextOrigin = edge(origin);
  const nextCount = edge(end) - nextOrigin;
  const apply = (o: number, c: number): void =>
    applyLayoutPatch(
      store,
      row
        ? { freezeRows: c, freezeOriginRow: c > 0 ? o : 0 }
        : { freezeCols: c, freezeOriginCol: c > 0 ? o : 0 },
    );
  apply(nextOrigin, nextCount);
  if (!history || history.isReplaying()) return;
  history.push({ undo: () => apply(origin, count), redo: () => apply(nextOrigin, nextCount) });
}

function applyLayoutPatch(store: SpreadsheetStore, patch: Partial<LayoutSlice>): void {
  store.setState((s) => ({ ...s, layout: { ...s.layout, ...patch } }));
}

function mergeRangesForSheet(state: State, sheet: number): Range[] {
  return [...state.merges.byAnchor.values()]
    .filter((range) => range.sheet === sheet)
    .map((range) => ({ ...range }));
}

function sameMergeRanges(a: readonly Range[], b: readonly Range[]): boolean {
  if (a.length !== b.length) return false;
  return a.every((left, index) => {
    const right = b[index];
    return (
      right !== undefined &&
      left.sheet === right.sheet &&
      left.r0 === right.r0 &&
      left.c0 === right.c0 &&
      left.r1 === right.r1 &&
      left.c1 === right.c1
    );
  });
}

function syncMergesToEngine(wb: WorkbookHandle, sheet: number, ranges: readonly Range[]): void {
  if (!wb.capabilities.merges) return;
  wb.engineClearMerges(sheet);
  for (const range of ranges) wb.engineAddMerge(sheet, range);
}

/** Re-point merges, conditional-format ranges, and the autofilter region after
 *  a row/col insert or delete so they track the cells they annotate. Each
 *  concern records its own history entry inside the surrounding transaction. */
function shiftAnchoredRanges(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  history: History | null,
  sheet: number,
  axis: 'row' | 'col',
  split: number,
  delta: number,
  nativeAxisOp: boolean,
): void {
  // Merges: rebuild both lookup maps from the shifted anchors. A merge fully
  // inside a deleted band, or collapsed to a single cell, is dropped.
  const mutateMerges = (): void => {
    store.setState((s) => {
      const byAnchor = new Map<string, Range>();
      const byCell = new Map<string, string>();
      for (const merge of s.merges.byAnchor.values()) {
        const shifted = merge.sheet === sheet ? shiftRangeAxis(merge, axis, split, delta) : merge;
        if (!shifted) continue;
        if (shifted.r0 === shifted.r1 && shifted.c0 === shifted.c1) continue;
        addMergeToMaps(byAnchor, byCell, shifted);
      }
      return { ...s, merges: { byAnchor, byCell } };
    });
  };
  // A native axis edit already moves merges inside the engine. During undo the
  // merge history entry runs before the inverse axis operation, so syncing the
  // engine here would make the native operation move the restored ranges a
  // second time. The native axis history entry restores the exact engine merge
  // snapshot after its inverse; the JS fallback still mirrors the store here.
  if (nativeAxisOp) {
    recordMergesChange(history, store, mutateMerges);
    const afterMerges = mergeRangesForSheet(store.getState(), sheet);
    // Native engines generally move merges correctly, but a merge reduced to
    // one cell is still reported by some versions. Excel removes that merge,
    // so reconcile the native result with the store's range normalization.
    if (wb.capabilities.merges && !sameMergeRanges(wb.getMerges(sheet), afterMerges)) {
      syncMergesToEngine(wb, sheet, afterMerges);
      if (history && !history.isReplaying()) {
        history.push({
          undo: () => {},
          redo: () => syncMergesToEngine(wb, sheet, afterMerges),
        });
      }
    }
  } else {
    recordMergesChangeWithEngine(history, store, wb, sheet, mutateMerges);
  }

  // Conditional formats: shift each rule's range; drop rules whose range is
  // fully consumed by a deletion.
  recordConditionalRulesChange(history, store, () => {
    store.setState((s) => {
      const rules: ConditionalRule[] = [];
      for (const rule of s.conditional.rules) {
        if (rule.range.sheet !== sheet) {
          rules.push(rule);
          continue;
        }
        const shifted = shiftRangeAxis(rule.range, axis, split, delta);
        if (!shifted) continue;
        rules.push({ ...rule, range: shifted });
      }
      return { ...s, conditional: { rules } };
    });
  });

  // Copy marquee. Transient UI state that history never captures, so it is
  // shifted outside the recorders — otherwise the outline would keep pointing
  // at the source's pre-shift indices. A cut payload keeps its original source
  // coordinates until paste; a successful structure edit invalidates that
  // payload, so Excel cancels the cut marquee instead of shifting it.
  const copyUi = store.getState().ui;
  if (copyUi.copyMode === 'cut') {
    if (copyUi.copyRanges) mutators.setCopyRanges(store, null);
    else mutators.setCopyRange(store, null);
  } else {
    store.setState((s) => {
      const shiftMarquee = (range: Range): Range | null =>
        range.sheet === sheet ? shiftRangeAxis(range, axis, split, delta) : range;
      const copyRange = s.ui.copyRange ? shiftMarquee(s.ui.copyRange) : null;
      const copyRanges = s.ui.copyRanges
        ? s.ui.copyRanges.map(shiftMarquee).filter((r): r is Range => r !== null)
        : null;
      if (copyRange === s.ui.copyRange && copyRanges === s.ui.copyRanges) return s;
      return {
        ...s,
        ui: { ...s.ui, copyRange, copyRanges: copyRanges?.length ? copyRanges : null },
      };
    });
  }

  // Autofilter region + per-column criteria.
  recordFilterChange(history, store, () => {
    store.setState((s) => {
      const fr = s.ui.filterRange;
      if (!fr || fr.sheet !== sheet) return s;
      const shifted = shiftRangeAxis(fr, axis, split, delta);
      return {
        ...s,
        ui: {
          ...s.ui,
          filterRange: shifted,
          filterCriteria: shifted
            ? shiftFilterCriteria(s.ui.filterCriteria, sheet, axis, split, delta)
            : [],
        },
      };
    });
  });
}

const cloneAxisOverride = (
  override: SheetAxisFormatOverride | undefined,
): SheetAxisFormatOverride | undefined =>
  override
    ? {
        baseline: {
          columns: override.baseline.columns.map((entry) => ({ ...entry })),
          rows: override.baseline.rows.map((entry) => ({ ...entry })),
        },
        desired: {
          columns: override.desired.columns.map((entry) => ({ ...entry })),
          rows: override.desired.rows.map((entry) => ({ ...entry })),
        },
        ...(override.directXfs
          ? {
              directXfs: override.directXfs.map((entry) => ({
                key: entry.key,
                before: { ...entry.before },
                afterXfIndex: entry.afterXfIndex,
              })),
            }
          : {}),
        ...(override.fontPatch ? { fontPatch: { ...override.fontPatch } } : {}),
      }
    : undefined;

const cloneSheetFontOverride = (
  override: SheetFontOverride | undefined,
): SheetFontOverride | undefined =>
  override
    ? {
        ...override,
        ...(override.patch ? { patch: { ...override.patch } } : {}),
        ...(override.baseline
          ? {
              baseline: {
                columns: override.baseline.columns.map((entry) => ({ ...entry })),
                rows: override.baseline.rows.map((entry) => ({ ...entry })),
              },
            }
          : {}),
      }
    : undefined;

interface StructureCacheSnapshot {
  readonly activeAxis: ActiveAxisFormatSnapshot | null;
  readonly activeWholeSheetFont: ActiveWholeSheetFontSnapshot | null;
  readonly syncedFormatKeys: ReadonlySet<string> | null;
}

/** The native structure child is the first entry in a composite history
 * transaction. Later entries can already have restored their own before
 * snapshots when the native replay fails, so keep an isolated full store
 * checkpoint for putting those children back at the committed after-state. */
type StructureStoreSnapshot = State;

export interface StructureReplayCheckpoint {
  readonly store: StructureStoreSnapshot;
  readonly caches: StructureCacheSnapshot;
}

export interface StructureReplayCheckpointRef {
  readonly capture?: () => StructureReplayCheckpoint;
  before?: StructureReplayCheckpoint;
  after?: StructureReplayCheckpoint;
}

const captureStructureStoreSnapshot = (store: SpreadsheetStore): StructureStoreSnapshot =>
  structuredClone(store.getState()) as StructureStoreSnapshot;

const restoreStructureStoreSnapshot = (
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  sheet: number,
  snapshot: StructureStoreSnapshot,
  caches: StructureCacheSnapshot | undefined,
): void => {
  withEngineSyncMuted(wb, () => {
    store.setState(() => structuredClone(snapshot) as State);
    if (!caches) return;
    restoreActiveAxisFormatOverrideForStructure(wb, sheet, caches.activeAxis);
    restoreActiveWholeSheetFontOverrideForStructure(wb, sheet, caches.activeWholeSheetFont);
    restoreSyncedFormatKeysForStructure(wb, sheet, caches.syncedFormatKeys);
  });
};

export const captureStructureReplayCheckpoint = (
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  sheet: number,
): StructureReplayCheckpoint => ({
  store: captureStructureStoreSnapshot(store),
  caches: {
    activeAxis: captureActiveAxisFormatOverrideForStructure(wb, sheet),
    activeWholeSheetFont: captureActiveWholeSheetFontOverrideForStructure(wb, sheet),
    syncedFormatKeys: captureSyncedFormatKeysForStructure(wb, sheet),
  },
});

/** Create a replay checkpoint handle for a composite command that invokes a
 * structural edit and appends more history children afterwards. The caller
 * owns the final marker; the structure child uses the same handle so native
 * failure recovery sees the committed outer transaction state. */
export const createStructureReplayCheckpointRef = (
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  sheet: number,
): StructureReplayCheckpointRef => ({
  capture: () => captureStructureReplayCheckpoint(store, wb, sheet),
});

const structureStyleRange = (sheet: number, axis: 'row' | 'col', split: number): Range => ({
  sheet,
  r0: axis === 'row' ? split : 0,
  c0: axis === 'col' ? split : 0,
  r1: MAX_ROW,
  c1: MAX_COL,
});

const withEngineSyncMuted = <T>(wb: WorkbookHandle, fn: () => T): T => {
  const candidate = wb as WorkbookHandle & {
    withEngineSyncMuted?: <R>(callback: () => R) => R;
  };
  return typeof candidate.withEngineSyncMuted === 'function'
    ? candidate.withEngineSyncMuted(fn)
    : fn();
};

const axisSnapshotFor = (
  wb: WorkbookHandle,
  sheet: number,
): ReturnType<WorkbookHandle['getSheetFontAxisSnapshot']> => {
  const candidate = wb as WorkbookHandle & {
    getSheetFontAxisSnapshot?: WorkbookHandle['getSheetFontAxisSnapshot'];
  };
  return typeof candidate.getSheetFontAxisSnapshot === 'function'
    ? candidate.getSheetFontAxisSnapshot(sheet)
    : null;
};

/** Publish a structural axis remap while the store mirror is muted. The
 * native operation has already shifted its records; this keeps all compact
 * JS/native caches in the same coordinate space before an ordinary store
 * subscriber can flush stale keys. */
function applyStructureFormatRemap(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  sheet: number,
  shift: AxisStructureShift,
  axisOverride: SheetAxisFormatOverride | undefined,
  sheetFontOverride: SheetFontOverride | undefined,
  cacheSnapshot?: StructureCacheSnapshot,
): void {
  withEngineSyncMuted(wb, () => {
    store.setState((state) => {
      const currentAxes = state.format.sheetAxisOverrides;
      const nextAxes = currentAxes ? new Map(currentAxes) : new Map();
      const nextAxisOverride = cloneAxisOverride(axisOverride);
      if (nextAxisOverride) nextAxes.set(sheet, nextAxisOverride);
      else nextAxes.delete(sheet);
      const currentFonts = state.format.sheetFontOverrides;
      const nextFonts = currentFonts ? new Map(currentFonts) : new Map();
      const nextSheetFontOverride = cloneSheetFontOverride(sheetFontOverride);
      if (nextSheetFontOverride) nextFonts.set(sheet, nextSheetFontOverride);
      else nextFonts.delete(sheet);
      return {
        ...state,
        format: {
          ...state.format,
          sheetAxisOverrides: nextAxes,
          sheetFontOverrides: nextFonts,
        },
      };
    });
    if (cacheSnapshot) {
      restoreActiveAxisFormatOverrideForStructure(wb, sheet, cacheSnapshot.activeAxis);
      restoreActiveWholeSheetFontOverrideForStructure(
        wb,
        sheet,
        cacheSnapshot.activeWholeSheetFont,
      );
      restoreSyncedFormatKeysForStructure(wb, sheet, cacheSnapshot.syncedFormatKeys);
    } else {
      remapActiveAxisFormatOverridesForStructure(wb, sheet, shift);
      remapActiveWholeSheetFontOverridesForStructure(wb, sheet, shift);
      remapSyncedFormatKeysForStructure(wb, sheet, shift);
    }
  });
}

/**
 * Apply a row/col shift using the engine's native `insertRows/deleteRows/
 * insertCols/deleteCols` ops. Pushes one history entry that inverts via the
 * opposite engine op (with a captured cell-band restore for delete). Cells
 * outside the shifted band are *not* touched in the store cache here; the
 * caller is expected to refresh via `replaceCells` after this returns.
 */
/** Whether deleting `count` lines at `split` removes the drawing object: the
 * engine drops a move-and-size object whose both corners lie in the band. */
const drawingRemovedByDelete = (
  object: EngineDrawingObject,
  axis: 'row' | 'col',
  split: number,
  count: number,
): boolean => {
  const inBand = (line: number): boolean => line >= split && line < split + count;
  return (
    object.editAs === AnchorEditAs.TwoCell &&
    inBand(axis === 'row' ? object.fromRow : object.fromCol) &&
    inBand(axis === 'row' ? object.toRow : object.toCol)
  );
};

function applyAxisShiftViaEngine(
  wb: WorkbookHandle,
  store: SpreadsheetStore,
  history: History | null,
  sheet: number,
  axis: 'row' | 'col',
  split: number,
  delta: number,
  restoreAfterFailedReplay?: (phase: 'redo' | 'undo') => void,
  replayCheckpoint?: StructureReplayCheckpointRef,
): boolean {
  if (delta === 0) return true;
  const positive = delta > 0;
  const count = Math.abs(delta);
  const record = history !== null && !history.isReplaying();
  const shift: AxisStructureShift = { axis, split, delta };
  const bridge = record ? formatHistoryBridgeFor(store) : undefined;
  const styleRange = structureStyleRange(sheet, axis, split);
  const rangeStyleCapable = record && wb.capabilities.rangeStyleSnapshot === true;
  const beforeRangeStyle = rangeStyleCapable ? wb.snapshotRangeStyle(styleRange) : null;
  // A capable native module must produce the opaque checkpoint before any
  // structural mutation. Falling back to the manual sparse checkpoint here
  // would silently lose blank-cell XFs and axis metadata on Undo.
  if (rangeStyleCapable && beforeRangeStyle === null) return false;
  const useRangeStyleCheckpoint = rangeStyleCapable;

  // A structure edit moves sparse direct XFs even when the store has no
  // projected CellFormat for them. Capture the source keys before native
  // mutation and their mapped destinations after it. Deleted-band keys stay in
  // the before checkpoint so Undo can recreate them after the native inverse.
  const sourceKeys = new Set<string>();
  if (record) {
    for (const key of store.getState().format.formats.keys()) {
      const addr = parseAddrKey(key);
      if (!addr || addr.sheet !== sheet) continue;
      const index = axis === 'row' ? addr.row : addr.col;
      if (index >= split) sourceKeys.add(key);
    }
    for (const entry of store.getState().format.sheetAxisOverrides?.get(sheet)?.directXfs ?? []) {
      sourceKeys.add(entry.key);
    }
  }
  const beforeNative: NativeFormatSnapshot =
    record && bridge && !useRangeStyleCheckpoint ? bridge.capture(sourceKeys) : new Map();
  const beforeTracking: FormatHistoryTrackingSnapshot =
    record && bridge ? bridge.captureTracking(sourceKeys) : new Map();
  const beforeAxes = record && !useRangeStyleCheckpoint ? axisSnapshotFor(wb, sheet) : null;
  let afterNative: NativeFormatSnapshot = new Map();
  let afterTracking: FormatHistoryTrackingSnapshot = new Map();
  let afterAxes = beforeAxes;
  let afterRangeStyle: Uint8Array | null = null;

  // For delete (delta < 0), capture cells in the band so undo can rewrite
  // them. Insert is fully invertible by the opposite engine op alone.
  const captured: CellRecord[] = [];
  const originalFormulas: FormulaRecord[] = [];
  if (!positive && record) {
    for (const c of wb.cells(sheet)) {
      const k = axis === 'row' ? c.addr.row : c.addr.col;
      if (k >= split && k < split + count) {
        captured.push({ addr: c.addr, value: c.value, formula: c.formula });
      }
    }
    originalFormulas.push(...collectAllFormulas(wb));
  }
  // The engine deletes the pictures and charts lying inside the band; undo puts
  // them back from these captures, under their own ids.
  const deletedDrawings: Uint8Array[] = [];
  if (!positive && record && wb.capabilities.drawings) {
    for (const object of wb.listDrawingObjects(sheet)) {
      if (!drawingRemovedByDelete(object, axis, split, count)) continue;
      const snapshot = wb.snapshotDrawing(sheet, object.objectId);
      if (snapshot === null) return false;
      deletedDrawings.push(snapshot);
    }
  }

  const canRestoreMerges = record && wb.capabilities.merges;
  const beforeMerges = canRestoreMerges ? wb.getMerges(sheet) : [];
  let afterMerges: Range[] = [];
  let initialApply = true;

  const runNativeOp = (insert: boolean): boolean => {
    let ok = false;
    if (insert) {
      ok =
        axis === 'row'
          ? wb.engineInsertRows(sheet, split, count)
          : wb.engineInsertCols(sheet, split, count);
    } else {
      ok =
        axis === 'row'
          ? wb.engineDeleteRows(sheet, split, count)
          : wb.engineDeleteCols(sheet, split, count);
    }
    if (!ok) return false;
    wb.recalcAuto();
    hydrateChartsFromEngine(wb, store, sheet);
    return true;
  };

  const restoreMerges = (snapshot: readonly Range[]): void => {
    if (!canRestoreMerges) return;
    if (!wb.engineClearMerges(sheet)) throw new Error('structure: failed to restore merges');
    for (const merge of snapshot) {
      if (!wb.engineAddMerge(sheet, merge)) throw new Error('structure: failed to restore merge');
    }
  };

  const afterKeys = (): Set<string> => {
    const keys = new Set<string>();
    for (const key of sourceKeys) {
      const addr = parseAddrKey(key);
      if (!addr) continue;
      const mapped = mapAxisAddress(addr, shift);
      if (mapped) keys.add(addrKey(mapped));
    }
    return keys;
  };

  const restoreCheckpoint = (
    axes: ReturnType<WorkbookHandle['getSheetFontAxisSnapshot']>,
    native: NativeFormatSnapshot,
    tracking: FormatHistoryTrackingSnapshot,
    rangeStyle: Uint8Array | null = null,
  ): void => {
    const errors: unknown[] = [];
    const currentAxes = axes !== null ? axisSnapshotFor(wb, sheet) : null;
    let currentRangeStyle: Uint8Array | null = null;
    if (rangeStyle !== null) {
      if (wb.capabilities.rangeStyleSnapshot !== true) {
        errors.push(new Error('structure: native range style snapshot capability unavailable'));
      } else {
        try {
          currentRangeStyle = wb.snapshotRangeStyle(styleRange);
          if (currentRangeStyle === null)
            throw new Error('structure: native range style rollback snapshot failed');
        } catch (error) {
          errors.push(error);
        }
      }
    }
    let currentNative: NativeFormatSnapshot | null = null;
    let currentTracking: FormatHistoryTrackingSnapshot | null = null;
    if (bridge) {
      try {
        currentNative = bridge.capture(native.keys());
        currentTracking = bridge.captureTracking(native.keys());
      } catch (error) {
        // Continue with the requested restore. If it fails too, the capture
        // error is retained alongside the setter error for the caller.
        errors.push(error);
      }
    }
    let axisFailed = false;
    let directFailed = false;
    let rangeFailed = false;
    const restore = (): void => {
      // A null range checkpoint means this is the legacy/manual path. Keep
      // the axis/direct-XF restore active for older native bundles instead of
      // treating the absence of a range snapshot as a successful restore.
      let rangeRestored = false;
      if (rangeStyle !== null) {
        try {
          if (
            wb.capabilities.rangeStyleSnapshot !== true ||
            !wb.restoreRangeStyle(sheet, rangeStyle)
          )
            throw new Error('structure: failed to restore native range styles');
          rangeRestored = true;
        } catch (error) {
          rangeFailed = true;
          errors.push(error);
        }
      }
      if (!rangeRestored) {
        // Axis and direct-XF recovery are independent. A failed axis setter
        // must not short-circuit a direct checkpoint restore (and vice versa).
        if (axes !== null) {
          try {
            if (!restoreNativeAxisSnapshot(wb, sheet, axes))
              throw new Error('structure: failed to restore native axis styles');
          } catch (error) {
            axisFailed = true;
            errors.push(error);
          }
        }
        if (bridge) {
          try {
            bridge.restore(native);
          } catch (error) {
            directFailed = true;
            errors.push(error);
          }
          if (!axisFailed && !directFailed) {
            try {
              bridge.restoreTracking(tracking);
            } catch (error) {
              errors.push(error);
            }
          }
        }
      } else if (bridge) {
        try {
          bridge.restoreTracking(tracking);
        } catch (error) {
          errors.push(error);
        }
      }
    };
    if (bridge) bridge.withSyncMuted(restore);
    else withEngineSyncMuted(wb, restore);
    if (axisFailed || directFailed) {
      const rollback = (): void => {
        if ((rangeFailed || errors.length > 0) && currentRangeStyle !== null) {
          try {
            if (!wb.restoreRangeStyle(sheet, currentRangeStyle))
              throw new Error('structure: failed to roll back native range styles');
          } catch (error) {
            errors.push(error);
          }
        }
        if (axisFailed && currentAxes !== null) {
          try {
            if (!restoreNativeAxisSnapshot(wb, sheet, currentAxes))
              throw new Error('structure: failed to roll back native axis styles');
          } catch (error) {
            errors.push(error);
          }
        }
        if (directFailed && bridge && currentNative !== null) {
          try {
            bridge.restore(currentNative);
            if (currentTracking !== null) bridge.restoreTracking(currentTracking);
          } catch (error) {
            errors.push(error);
          }
        }
      };
      if (bridge) bridge.withSyncMuted(rollback);
      else withEngineSyncMuted(wb, rollback);
    } else if (errors.length > 0 && currentRangeStyle !== null) {
      const rollbackRange = (): void => {
        try {
          if (!wb.restoreRangeStyle(sheet, currentRangeStyle as Uint8Array))
            throw new Error('structure: failed to roll back native range styles');
        } catch (error) {
          errors.push(error);
        }
        if (bridge && currentTracking !== null) {
          try {
            bridge.restoreTracking(currentTracking);
          } catch (error) {
            errors.push(error);
          }
        }
      };
      if (bridge) bridge.withSyncMuted(rollbackRange);
      else withEngineSyncMuted(wb, rollbackRange);
    }
    if (errors.length === 1) throw errors[0];
    if (errors.length > 1) {
      throw new AggregateError(errors, 'structure: native format checkpoint restore failed', {
        cause: errors[0],
      });
    }
  };

  const restoreDeletedBand = (): void => {
    if (positive) return;
    wb.withBatchedRecalc(() => {
      for (const c of captured) writeCell(wb, c.addr, c.value, c.formula);
      for (const entry of originalFormulas) wb.setFormula(entry.addr, entry.formula);
    });
    if (deletedDrawings.length === 0) return;
    for (const snapshot of deletedDrawings) {
      if (wb.restoreDrawing(sheet, snapshot) === null) {
        throw new Error('structure: failed to restore a deleted drawing');
      }
    }
    hydrateChartsFromEngine(wb, store, sheet);
  };

  const rollbackRedo = (): void => {
    if (!runNativeOp(!positive)) throw new Error('structure: native redo rollback failed');
    if (canRestoreMerges) restoreMerges(beforeMerges);
    restoreDeletedBand();
    restoreCheckpoint(beforeAxes, beforeNative, beforeTracking, beforeRangeStyle);
    wb.recalcAuto();
  };

  const rollbackUndo = (): void => {
    if (!runNativeOp(positive)) throw new Error('structure: native undo rollback failed');
    if (canRestoreMerges) restoreMerges(afterMerges);
    restoreCheckpoint(afterAxes, afterNative, afterTracking, afterRangeStyle);
    wb.recalcAuto();
  };

  const apply = (): void => {
    let applied = false;
    try {
      if (!initialApply && replayCheckpoint) {
        replayCheckpoint.before =
          replayCheckpoint.capture?.() ?? captureStructureReplayCheckpoint(store, wb, sheet);
      }
      if (!runNativeOp(positive)) throw new Error('structure: native axis edit failed');
      applied = true;
      if (canRestoreMerges) {
        if (initialApply) afterMerges = wb.getMerges(sheet);
        else restoreMerges(afterMerges);
      }
      if (!initialApply) {
        restoreCheckpoint(afterAxes, afterNative, afterTracking, afterRangeStyle);
        wb.recalcAuto();
      }
      initialApply = false;
    } catch (error) {
      if (!applied) {
        const recoveryErrors: unknown[] = [];
        try {
          restoreAfterFailedReplay?.('redo');
        } catch (rollbackError) {
          recoveryErrors.push(rollbackError);
        }
        if (recoveryErrors.length > 0)
          throw new AggregateError([error, ...recoveryErrors], 'structure: redo recovery failed', {
            cause: error,
          });
        throw error;
      }
      const recoveryErrors: unknown[] = [];
      try {
        rollbackRedo();
      } catch (rollbackError) {
        recoveryErrors.push(rollbackError);
      }
      try {
        restoreAfterFailedReplay?.('redo');
      } catch (restoreError) {
        recoveryErrors.push(restoreError);
      }
      if (recoveryErrors.length > 0)
        throw new AggregateError([error, ...recoveryErrors], 'structure: redo rollback failed', {
          cause: error,
        });
      throw error;
    }
  };
  const invert = (): void => {
    let applied = false;
    try {
      if (replayCheckpoint?.after === undefined) {
        if (replayCheckpoint) {
          replayCheckpoint.after =
            replayCheckpoint.capture?.() ?? captureStructureReplayCheckpoint(store, wb, sheet);
        }
      }
      if (!runNativeOp(!positive)) throw new Error('structure: native axis undo failed');
      applied = true;
      if (canRestoreMerges) restoreMerges(beforeMerges);
      restoreDeletedBand();
      wb.recalcAuto();
      restoreCheckpoint(beforeAxes, beforeNative, beforeTracking, beforeRangeStyle);
      wb.recalcAuto();
    } catch (error) {
      if (!applied) {
        const recoveryErrors: unknown[] = [];
        try {
          restoreAfterFailedReplay?.('undo');
        } catch (rollbackError) {
          recoveryErrors.push(rollbackError);
        }
        if (recoveryErrors.length > 0)
          throw new AggregateError([error, ...recoveryErrors], 'structure: undo recovery failed', {
            cause: error,
          });
        throw error;
      }
      const recoveryErrors: unknown[] = [];
      try {
        rollbackUndo();
      } catch (rollbackError) {
        recoveryErrors.push(rollbackError);
      }
      try {
        restoreAfterFailedReplay?.('undo');
      } catch (restoreError) {
        recoveryErrors.push(restoreError);
      }
      if (recoveryErrors.length > 0)
        throw new AggregateError([error, ...recoveryErrors], 'structure: undo rollback failed', {
          cause: error,
        });
      throw error;
    }
  };

  let initialApplied = false;
  try {
    if (!runNativeOp(positive)) return false;
    initialApplied = true;
    if (canRestoreMerges) afterMerges = wb.getMerges(sheet);
    if (record) {
      afterAxes = !useRangeStyleCheckpoint ? axisSnapshotFor(wb, sheet) : null;
      if (useRangeStyleCheckpoint) {
        afterRangeStyle = wb.snapshotRangeStyle(styleRange);
        if (afterRangeStyle === null)
          throw new Error('structure: native range style after-snapshot failed');
      }
      if (bridge) {
        const keys = afterKeys();
        if (!useRangeStyleCheckpoint) afterNative = bridge.capture(keys);
        afterTracking = bridge.captureTracking(keys);
      }
    }
  } catch (error) {
    if (!initialApplied) throw error;
    const rollbackErrors: unknown[] = [];
    try {
      if (!runNativeOp(!positive)) throw new Error('structure: native axis rollback failed');
      if (canRestoreMerges) restoreMerges(beforeMerges);
      restoreDeletedBand();
      restoreCheckpoint(beforeAxes, beforeNative, beforeTracking, beforeRangeStyle);
      wb.recalcAuto();
    } catch (rollbackError) {
      rollbackErrors.push(rollbackError);
    }
    if (rollbackErrors.length > 0) {
      throw new AggregateError(
        [error, ...rollbackErrors],
        'structure: native axis edit rollback failed',
        {
          cause: error,
        },
      );
    }
    throw error;
  }
  initialApply = false;
  if (record) {
    history.push({ undo: invert, redo: apply });
  }
  return true;
}

/** Shared body of the four insert/delete verbs. Cells, formats, sizes, hidden
 *  and outline sets, freeze panes and anchored ranges all shift along `axis`
 *  inside one history transaction. */
function editAxis(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  history: History | null,
  axis: 'row' | 'col',
  kind: 'insert' | 'delete',
  at: number,
  count: number,
  outerReplayCheckpoint?: StructureReplayCheckpointRef,
): boolean {
  const max = axis === 'row' ? MAX_ROW : MAX_COL;
  const edit = normalizeAxisEdit(at, count, max, kind);
  if (!edit) return false;
  const split = edit.at;
  const n = edit.count;
  const delta = kind === 'insert' ? n : -n;
  const sheet = store.getState().data.sheetIndex;
  if (blockedByProtection(store, sheet, `${kind}${axis === 'row' ? 'Rows' : 'Cols'}`)) {
    return false;
  }
  if (kind === 'insert' && insertionWouldOverflow(store, wb, sheet, axis, split, n)) return false;
  const nativeAxisOp = wb.capabilities.insertDeleteRowsCols;
  const structureShift: AxisStructureShift = { axis, split, delta };
  const inverseStructureShift: AxisStructureShift = { axis, split, delta: -delta };
  const beforeAxisOverride = cloneAxisOverride(
    store.getState().format.sheetAxisOverrides?.get(sheet),
  );
  const beforeSheetFontOverride = cloneSheetFontOverride(
    store.getState().format.sheetFontOverrides?.get(sheet),
  );
  const afterAxisOverride = beforeAxisOverride
    ? remapAxisFormatOverride(beforeAxisOverride, structureShift)
    : undefined;
  const afterSheetFontOverride = beforeSheetFontOverride
    ? remapSheetFontOverride(beforeSheetFontOverride, structureShift)
    : undefined;
  const captureStructureCaches = history !== null && !history.isReplaying();
  const beforeStructureCaches: StructureCacheSnapshot | undefined = captureStructureCaches
    ? {
        activeAxis: captureActiveAxisFormatOverrideForStructure(wb, sheet),
        activeWholeSheetFont: captureActiveWholeSheetFontOverrideForStructure(wb, sheet),
        syncedFormatKeys: captureSyncedFormatKeysForStructure(wb, sheet),
      }
    : undefined;

  // Native structure history is the first child in a transaction so redo
  // applies the engine mutation before the muted JS cache remap. Undo runs
  // those later children first; if the native restore then rejects, put the
  // committed after-cache back before propagating the failure.
  let restoreAfterFailedReplay: ((phase: 'redo' | 'undo') => void) | undefined;
  const replayCheckpoint = outerReplayCheckpoint ?? {};

  if (history) history.begin();
  try {
    // 1. shift cells & rewrite formula refs.
    if (nativeAxisOp) {
      if (
        !applyAxisShiftViaEngine(
          wb,
          store,
          history,
          sheet,
          axis,
          split,
          delta,
          (phase) => restoreAfterFailedReplay?.(phase),
          replayCheckpoint,
        )
      )
        return false;
    } else {
      applyAxisShiftToCells(wb, sheet, axis, split, delta);
    }

    // 2. Shift compact axis state and its workbook-local caches before any
    // ordinary store subscriber can flush the old coordinate keys. This is a
    // separate history child because native structure history is first and
    // format-map history is recorded below it.
    applyStructureFormatRemap(
      store,
      wb,
      sheet,
      structureShift,
      afterAxisOverride,
      afterSheetFontOverride,
    );
    const afterStructureCaches: StructureCacheSnapshot | undefined = captureStructureCaches
      ? {
          activeAxis: captureActiveAxisFormatOverrideForStructure(wb, sheet),
          activeWholeSheetFont: captureActiveWholeSheetFontOverrideForStructure(wb, sheet),
          syncedFormatKeys: captureSyncedFormatKeysForStructure(wb, sheet),
        }
      : undefined;
    restoreAfterFailedReplay = (phase) => {
      const checkpoint = phase === 'redo' ? replayCheckpoint.before : replayCheckpoint.after;
      if (checkpoint) {
        restoreStructureStoreSnapshot(store, wb, sheet, checkpoint.store, checkpoint.caches);
        return;
      }
      applyStructureFormatRemap(
        store,
        wb,
        sheet,
        structureShift,
        afterAxisOverride,
        afterSheetFontOverride,
        afterStructureCaches,
      );
    };
    if (history) {
      history.push({
        undo: () =>
          applyStructureFormatRemap(
            store,
            wb,
            sheet,
            inverseStructureShift,
            beforeAxisOverride,
            beforeSheetFontOverride,
            beforeStructureCaches,
          ),
        redo: () =>
          applyStructureFormatRemap(
            store,
            wb,
            sheet,
            structureShift,
            afterAxisOverride,
            afterSheetFontOverride,
            afterStructureCaches,
          ),
      });
    }

    // 3. shift sparse cell formats; an insertion inherits the format of the band above/left.
    recordFormatChange(history, store, () => {
      store.setState((s) => {
        const shift = axis === 'row' ? shiftFormatsByRow : shiftFormatsByCol;
        const inherit = axis === 'row' ? inheritFormatsByRow : inheritFormatsByCol;
        const shifted = shift(s.format.formats, sheet, split, delta);
        const formats =
          kind === 'insert' ? inherit(s.format.formats, shifted, sheet, split, n) : shifted;
        return { ...s, format: { ...s.format, formats } };
      });
    });

    // 4. shift layout (sizes, hidden set, outline levels) and the frozen band.
    recordLayoutChange(history, store, () => {
      const before = captureLayoutSnapshot(store.getState());
      const row = axis === 'row';
      const sizes = row ? before.rowHeights : before.colWidths;
      const hidden = row ? before.hiddenRows : before.hiddenCols;
      const outline = row ? before.outlineRows : before.outlineCols;
      const nextSizes =
        kind === 'insert'
          ? shiftIndexedMapWithInheritance(sizes, split, delta, n, max)
          : shiftIndexedMap(sizes, split, delta, max);
      const nextHidden = shiftIndexedSet(hidden, split, delta, max);
      const nextOutline = shiftIndexedMap(outline, split, delta, max);
      applyLayoutPatch(
        store,
        row
          ? {
              rowHeights: nextSizes,
              hiddenRows: nextHidden,
              outlineRows: nextOutline,
            }
          : {
              colWidths: nextSizes,
              hiddenCols: nextHidden,
              outlineCols: nextOutline,
            },
      );
    });
    shiftFreezeBand(store, history, axis, kind, split, n);

    // 5. re-point merges, conditional formats, and the autofilter region.
    shiftAnchoredRanges(store, wb, history, sheet, axis, split, delta, nativeAxisOp);
    if (history && captureStructureCaches && outerReplayCheckpoint === undefined) {
      // Undo visits children in reverse order. This final marker therefore
      // arms the committed after-state immediately before the other children
      // begin undoing; redo captures its before-state at the native child so
      // selections or unrelated edits made after this command are not rolled
      // back by a failed replay.
      history.push({
        undo: () => {
          replayCheckpoint.after = captureStructureReplayCheckpoint(store, wb, sheet);
        },
        redo: () => {
          replayCheckpoint.before = undefined;
          replayCheckpoint.after = undefined;
        },
      });
    }
    return true;
  } finally {
    if (history) history.end();
  }
}

/** Insert `count` blank rows at `atRow` on the active sheet. Cells, formats,
 *  row heights, freeze pane, and hidden-row set all shift down. Wrapped in a
 *  single history transaction. When the engine exposes
 *  `insertDeleteRowsCols`, the cell-shift step delegates to the native op
 *  for cross-sheet ref / merge / array-formula correctness; otherwise falls
 *  back to a JS-side cell rewrite. */
export function insertRows(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  history: History | null,
  atRow: number,
  count = 1,
  replayCheckpoint?: StructureReplayCheckpointRef,
): boolean {
  return editAxis(store, wb, history, 'row', 'insert', atRow, count, replayCheckpoint);
}

export function deleteRows(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  history: History | null,
  atRow: number,
  count = 1,
  replayCheckpoint?: StructureReplayCheckpointRef,
): boolean {
  return editAxis(store, wb, history, 'row', 'delete', atRow, count, replayCheckpoint);
}

export function insertCols(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  history: History | null,
  atCol: number,
  count = 1,
  replayCheckpoint?: StructureReplayCheckpointRef,
): boolean {
  return editAxis(store, wb, history, 'col', 'insert', atCol, count, replayCheckpoint);
}

export function deleteCols(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  history: History | null,
  atCol: number,
  count = 1,
  replayCheckpoint?: StructureReplayCheckpointRef,
): boolean {
  return editAxis(store, wb, history, 'col', 'delete', atCol, count, replayCheckpoint);
}
