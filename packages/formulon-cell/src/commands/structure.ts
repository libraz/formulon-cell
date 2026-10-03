import { MAX_COL, MAX_ROW } from '../engine/address.js';
import type { Addr, CellValue, Range } from '../engine/types.js';
import { writeCell } from '../engine/value.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
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
  parseFormatKey,
  shiftFilterCriteria,
  shiftFormatsByCol,
  shiftFormatsByRow,
  shiftIndexedMap,
  shiftIndexedMapWithInheritance,
  shiftIndexedSet,
  shiftRangeAxis,
} from './axis-shift.js';
import { recordFilterChange } from './filter.js';
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

interface FormulaRecord {
  addr: Addr;
  formula: string;
}

function collectAllFormulas(wb: WorkbookHandle): FormulaRecord[] {
  const out: FormulaRecord[] = [];
  for (let sheet = 0; sheet < wb.sheetCount; sheet += 1) {
    for (const c of wb.cells(sheet)) {
      if (c.formula !== null) out.push({ addr: c.addr, formula: c.formula });
    }
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
    const addr = parseFormatKey(key);
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

/**
 * Apply a row/col shift using the engine's native `insertRows/deleteRows/
 * insertCols/deleteCols` ops. Pushes one history entry that inverts via the
 * opposite engine op (with a captured cell-band restore for delete). Cells
 * outside the shifted band are *not* touched in the store cache here; the
 * caller is expected to refresh via `replaceCells` after this returns.
 */
function applyAxisShiftViaEngine(
  wb: WorkbookHandle,
  history: History | null,
  sheet: number,
  axis: 'row' | 'col',
  split: number,
  delta: number,
): boolean {
  if (delta === 0) return true;
  const positive = delta > 0;
  const count = Math.abs(delta);
  const record = history !== null && !history.isReplaying();

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
    return true;
  };

  const restoreMerges = (snapshot: readonly Range[]): void => {
    if (!canRestoreMerges) return;
    if (!wb.engineClearMerges(sheet)) throw new Error('structure: failed to restore merges');
    for (const merge of snapshot) {
      if (!wb.engineAddMerge(sheet, merge)) throw new Error('structure: failed to restore merge');
    }
  };

  const apply = (): void => {
    if (!runNativeOp(positive)) throw new Error('structure: native axis edit failed');
    if (canRestoreMerges) {
      if (initialApply) afterMerges = wb.getMerges(sheet);
      else restoreMerges(afterMerges);
    }
    initialApply = false;
  };
  const invert = (): void => {
    if (!runNativeOp(!positive)) throw new Error('structure: native axis undo failed');
    if (canRestoreMerges) restoreMerges(beforeMerges);

    if (!positive) {
      // Restore the deleted band and every original formula. The native
      // inverse recreates the row/column slots, but #REF! replacements made
      // by the delete cannot be reconstructed by that inverse alone.
      wb.withBatchedRecalc(() => {
        for (const c of captured) writeCell(wb, c.addr, c.value, c.formula);
        for (const entry of originalFormulas) wb.setFormula(entry.addr, entry.formula);
      });
    }
    wb.recalcAuto();
  };

  if (!runNativeOp(positive)) return false;
  if (canRestoreMerges) afterMerges = wb.getMerges(sheet);
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

  if (history) history.begin();
  try {
    // 1. shift cells & rewrite formula refs.
    if (nativeAxisOp) {
      if (!applyAxisShiftViaEngine(wb, history, sheet, axis, split, delta)) return false;
    } else {
      applyAxisShiftToCells(wb, sheet, axis, split, delta);
    }

    // 2. shift formats; an insertion inherits the format of the band above/left.
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

    // 3. shift layout (sizes, hidden set, outline levels, freeze count).
    recordLayoutChange(history, store, () => {
      const before = captureLayoutSnapshot(store.getState());
      const row = axis === 'row';
      const sizes = row ? before.rowHeights : before.colWidths;
      const hidden = row ? before.hiddenRows : before.hiddenCols;
      const outline = row ? before.outlineRows : before.outlineCols;
      const freeze = row ? before.freezeRows : before.freezeCols;
      const nextSizes =
        kind === 'insert'
          ? shiftIndexedMapWithInheritance(sizes, split, delta, n, max)
          : shiftIndexedMap(sizes, split, delta, max);
      const nextHidden = shiftIndexedSet(hidden, split, delta, max);
      const nextOutline = shiftIndexedMap(outline, split, delta, max);
      let nextFreeze = freeze;
      if (freeze > split) {
        nextFreeze = kind === 'insert' ? freeze + n : Math.max(split, freeze - n);
      }
      applyLayoutPatch(
        store,
        row
          ? {
              rowHeights: nextSizes,
              hiddenRows: nextHidden,
              outlineRows: nextOutline,
              freezeRows: nextFreeze,
            }
          : {
              colWidths: nextSizes,
              hiddenCols: nextHidden,
              outlineCols: nextOutline,
              freezeCols: nextFreeze,
            },
      );
    });

    // 4. re-point merges, conditional formats, and the autofilter region.
    shiftAnchoredRanges(store, wb, history, sheet, axis, split, delta, nativeAxisOp);
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
): boolean {
  return editAxis(store, wb, history, 'row', 'insert', atRow, count);
}

export function deleteRows(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  history: History | null,
  atRow: number,
  count = 1,
): boolean {
  return editAxis(store, wb, history, 'row', 'delete', atRow, count);
}

export function insertCols(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  history: History | null,
  atCol: number,
  count = 1,
): boolean {
  return editAxis(store, wb, history, 'col', 'insert', atCol, count);
}

export function deleteCols(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  history: History | null,
  atCol: number,
  count = 1,
): boolean {
  return editAxis(store, wb, history, 'col', 'delete', atCol, count);
}
