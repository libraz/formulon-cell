import { addrKey, MAX_COL, MAX_ROW, parseAddrKey } from '../engine/address.js';
import { captureActiveAxisFormatOverrideForStructure } from '../engine/axis-format-sync.js';
import {
  type AxisStructureShift,
  remapAxisFormatOverride,
  remapSheetFontOverride,
} from '../engine/axis-structure-remap.js';
import { captureSyncedFormatKeysForStructure } from '../engine/synced-format-keys.js';
import type { Addr, Range } from '../engine/types.js';
import { captureActiveWholeSheetFontOverrideForStructure } from '../engine/whole-sheet-font-sync.js';
import type { CellPatch, WorkbookHandle } from '../engine/workbook-handle.js';
import { rangesIntersect } from '../store/selection-geometry.js';
import { type LayoutSlice, mutators, type SpreadsheetStore, type State } from '../store/store.js';
import {
  inheritFormatsByCol,
  inheritFormatsByRow,
  normalizeAxisEdit,
  shiftFormatsByCol,
  shiftFormatsByRow,
  shiftHyperlinkSpans,
  shiftIndexedMap,
  shiftIndexedMapWithInheritance,
  shiftIndexedSet,
  shiftRangeValidations,
} from './axis-shift.js';
import type { CellRecord } from './cell-shift.js';
import { isMacChrome } from './chrome-platform.js';
import { applyPatchesAtomic } from './clipboard/apply-patches-atomic.js';
import { alertDataTableEdit, axisEditSplitsDataTable } from './data-table-guard.js';
import { filterVisibleRowRuns } from './filter-visible.js';
import { adjustFormulaForRowColEdit } from './formula-refs.js';
import { type History, transactIn } from './history.js';
import { alertPivotEdit, axisEditCrossesPivot } from './pivot-guard.js';
import { blockedByProtection, isBandUnlocked, sheetPermits } from './protection.js';
import { captureLayoutSnapshot, recordFormatChange, recordLayoutChange } from './slice-history.js';
import { shiftAnchoredRanges } from './structure-anchored-ranges.js';
import {
  applyStructureFormatRemap,
  captureStructureReplayCheckpoint,
  cloneAxisOverride,
  cloneSheetFontOverride,
  restoreStructureStoreSnapshot,
  type StructureCacheSnapshot,
  type StructureReplayCheckpointRef,
} from './structure-checkpoint.js';
import { applyAxisShiftViaEngine } from './structure-native-op.js';
import {
  keptTableDataRow,
  rowDeleteHitsTableHeader,
  shiftTablesForCols,
  shiftTablesForRows,
} from './table-rows.js';

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
      ? [
          state.layout.rowHeights,
          state.layout.hiddenRows,
          state.layout.outlineRows,
          state.layout.collapsedRows ?? new Set<number>(),
          state.layout.shownRows ?? new Set<number>(),
        ]
      : [
          state.layout.colWidths,
          state.layout.hiddenCols,
          state.layout.outlineCols,
          state.layout.collapsedCols ?? new Set<number>(),
        ];
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
 *  the band are replaced with `#REF!`. One engine transaction: a refused
 *  write leaves every cell as it was and records nothing. */
function applyAxisShiftToCells(
  wb: WorkbookHandle,
  history: History | null,
  sheet: number,
  axis: 'row' | 'col',
  split: number,
  delta: number,
): void {
  if (delta === 0) return;
  const all = collectAllCells(wb, sheet);
  applyPatchesAtomic(wb, axisShiftedCellPatches(all, axis, split, delta), history);
}

function axisShiftedCellPatches(
  all: readonly CellRecord[],
  axis: 'row' | 'col',
  split: number,
  delta: number,
): CellPatch[] {
  const max = axis === 'row' ? MAX_ROW : MAX_COL;
  const writes = new Map<string, CellPatch>();
  for (const c of all) {
    const k = axis === 'row' ? c.addr.row : c.addr.col;
    const inMovedBand = k >= split;
    const inDeletedBand = delta < 0 && inMovedBand && k < split - delta;
    if (inDeletedBand) continue; // dropped

    const newFormula = c.formula ? adjustFormulaForRowColEdit(c.formula, axis, split, delta) : null;

    if (inMovedBand) {
      const target = k + delta;
      if (target < 0 || target > max) continue;
      const newAddr: Addr =
        axis === 'row' ? { ...c.addr, row: target } : { ...c.addr, col: target };
      writes.set(addrKey(newAddr), { addr: newAddr, value: c.value, formula: newFormula });
    } else if (newFormula !== c.formula && newFormula !== null) {
      // Stationary cell whose formula references the band: rewrite in place.
      writes.set(addrKey(c.addr), { addr: c.addr, value: c.value, formula: newFormula });
    }
  }
  // Band cells no moved cell lands on are vacated.
  const blanks: CellPatch[] = [];
  for (const c of all) {
    const k = axis === 'row' ? c.addr.row : c.addr.col;
    if (k >= split && !writes.has(addrKey(c.addr))) {
      blanks.push({ addr: c.addr, value: { kind: 'blank' }, formula: null });
    }
  }
  return [...writes.values(), ...blanks];
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
  const sheet = store.getState().data.sheetIndex;
  const nativeAxisOp = wb.capabilities.insertDeleteRowsCols;
  // Deleting every data row of a table keeps the first one, blank: the engine is given the whole
  // delete and removes the rows below it, which are what everything else shifts by.
  const nativeDelete =
    nativeAxisOp && axis === 'row' && kind === 'delete'
      ? keptTableDataRow(store, sheet, edit.at, edit.count)
      : null;
  const split = nativeDelete ? edit.at + 1 : edit.at;
  const n = nativeDelete ? edit.count - 1 : edit.count;
  const delta = kind === 'insert' ? n : -n;
  const permission =
    kind === 'insert'
      ? axis === 'row'
        ? 'insertRows'
        : 'insertColumns'
      : axis === 'row'
        ? 'deleteRows'
        : 'deleteColumns';
  if (blockedByProtection(store, sheet, `${kind}${axis === 'row' ? 'Rows' : 'Cols'}`, permission)) {
    return false;
  }
  // Windows rule: a protected sheet deletes only rows (columns) holding no locked cell; the Mac
  // desktop deletes them by the option alone.
  if (
    kind === 'delete' &&
    !isMacChrome(store) &&
    !sheetPermits(store.getState(), sheet) &&
    !isBandUnlocked(store.getState(), sheet, axis, edit.at, edit.count)
  )
    return false;
  if (kind === 'insert' && insertionWouldOverflow(store, wb, sheet, axis, split, n)) return false;
  if (axisEditSplitsDataTable(wb, sheet, axis, kind, edit.at, edit.count)) {
    alertDataTableEdit(wb);
    return false;
  }
  if (axisEditCrossesPivot(wb, sheet, axis, kind, edit.at, edit.count)) {
    alertPivotEdit(wb);
    return false;
  }
  // Excel offers no delete over a table's header row.
  if (
    axis === 'row' &&
    kind === 'delete' &&
    rowDeleteHitsTableHeader(store, sheet, edit.at, edit.count)
  )
    return false;
  const selectionGrown = selectionGrownByMerges(store.getState(), sheet);
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

  return transactIn(history, () => {
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
          nativeDelete ? { split: edit.at, count: edit.count } : undefined,
        )
      )
        return false;
    } else {
      applyAxisShiftToCells(wb, history, sheet, axis, split, delta);
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
        const inherited =
          kind === 'insert' ? inherit(s.format.formats, shifted, sheet, split, n) : shifted;
        const formats = shiftHyperlinkSpans(s.format.formats, inherited, sheet, axis, split, delta);
        const rangeValidations = shiftRangeValidations(
          s.format.rangeValidations,
          sheet,
          axis,
          split,
          delta,
        );
        return { ...s, format: { ...s.format, formats, rangeValidations } };
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
      const collapsed = row ? before.collapsedRows : before.collapsedCols;
      const nextCollapsed = shiftIndexedSet(collapsed, split, delta, max);
      applyLayoutPatch(
        store,
        row
          ? {
              rowHeights: nextSizes,
              hiddenRows: nextHidden,
              outlineRows: nextOutline,
              collapsedRows: nextCollapsed,
              shownRows: shiftIndexedSet(before.shownRows, split, delta, max),
            }
          : {
              colWidths: nextSizes,
              hiddenCols: nextHidden,
              outlineCols: nextOutline,
              collapsedCols: nextCollapsed,
            },
      );
    });
    shiftFreezeBand(store, history, axis, kind, split, n);

    // 5. re-point merges, conditional formats, and the autofilter region.
    shiftAnchoredRanges(store, wb, history, sheet, axis, split, delta, nativeAxisOp);
    if (axis === 'row') shiftTablesForRows(store, wb, history, sheet, split, delta, nativeAxisOp);
    else if (kind === 'insert')
      shiftTablesForCols(store, wb, history, sheet, split, n, nativeAxisOp);
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
    if (selectionGrown) mutators.setRange(store, selectionGrown);
    return true;
  });
}

/** The selection widened to every merge it touches, or null when it already
 *  covers them. A whole-row selection that crosses a merge keeps the merge's
 *  full extent through an insert or delete. */
function selectionGrownByMerges(state: State, sheet: number): Range | null {
  const current = state.selection.range;
  if (current.sheet !== sheet) return null;
  let grown: Range = { ...current };
  for (let changed = true; changed; ) {
    changed = false;
    for (const merge of state.merges.byAnchor.values()) {
      if (merge.sheet !== sheet || !rangesIntersect(merge, grown)) continue;
      const next = {
        ...grown,
        r0: Math.min(grown.r0, merge.r0),
        c0: Math.min(grown.c0, merge.c0),
        r1: Math.max(grown.r1, merge.r1),
        c1: Math.max(grown.c1, merge.c1),
      };
      if (
        next.r0 !== grown.r0 ||
        next.c0 !== grown.c0 ||
        next.r1 !== grown.r1 ||
        next.c1 !== grown.c1
      ) {
        grown = next;
        changed = true;
      }
    }
  }
  const same =
    grown.r0 === current.r0 &&
    grown.c0 === current.c0 &&
    grown.r1 === current.r1 &&
    grown.c1 === current.c1;
  return same ? null : grown;
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

/**
 * The user's Delete Sheet Rows: rows an AutoFilter hides survive, only the
 * visible runs go, bottom-up so the earlier runs keep their indices.
 */
export function deleteVisibleRows(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  history: History | null,
  atRow: number,
  count = 1,
  replayCheckpoint?: StructureReplayCheckpointRef,
): boolean {
  const sheet = store.getState().data.sheetIndex;
  const runs = filterVisibleRowRuns(store.getState(), sheet, atRow, atRow + count - 1);
  const [first] = runs;
  if (runs.length === 0) return false;
  if (runs.length === 1 && first && first[0] === atRow && first[1] === atRow + count - 1) {
    return editAxis(store, wb, history, 'row', 'delete', atRow, count, replayCheckpoint);
  }
  return transactIn(history, () => {
    let ok = true;
    for (const [r0, r1] of [...runs].reverse()) {
      ok = editAxis(store, wb, history, 'row', 'delete', r0, r1 - r0 + 1) && ok;
    }
    return ok;
  });
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
