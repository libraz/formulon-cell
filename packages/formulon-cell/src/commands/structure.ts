import { addrKey } from '../engine/address.js';
import type { Addr, CellValue, Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import {
  type CellFormat,
  type ConditionalRule,
  type LayoutSlice,
  mutators,
  type SpreadsheetStore,
  type State,
  type ValueFilterCriteria,
} from '../store/store.js';
import {
  type AutofitOptions,
  computeAutofitColWidth,
  computeAutofitRowHeight,
  createAutofitMeasureContext,
} from './autofit-measurement.js';
import { recordFilterChange } from './filter.js';
import { adjustFormulaForRowColEdit } from './formula-refs.js';
import {
  captureLayoutSnapshot,
  type History,
  recordConditionalRulesChange,
  recordFormatChange,
  recordLayoutChange,
  recordLayoutChangeWithEngine,
  recordMergesChange,
  recordMergesChangeWithEngine,
} from './history.js';
import { isSheetProtected } from './protection.js';

/** Spreadsheet-parity gate for row/col structure changes. When `sheet` is
 *  protected the operation is rejected (no-op + warning) regardless of
 *  per-cell locks — spreadsheets disable the insert/delete row/col commands
 *  wholesale on protected sheets. */
function blockedByProtection(store: SpreadsheetStore, sheet: number, op: string): boolean {
  if (!isSheetProtected(store.getState(), sheet)) return false;
  // eslint-disable-next-line no-console
  console.warn(`formulon-cell: ${op} blocked — sheet ${sheet} is protected`);
  return true;
}

interface CellRecord {
  addr: Addr;
  value: CellValue;
  formula: string | null;
}

const MAX_ROW = 1048575;
const MAX_COL = 16383;
const MAX_MATERIALIZED_LAYOUT_ROWS = 100_000;

interface AxisEdit {
  at: number;
  count: number;
}

function normalizeAxisEdit(
  at: number,
  count: number,
  max: number,
  kind: 'insert' | 'delete',
): AxisEdit | null {
  if (!Number.isInteger(at) || !Number.isInteger(count)) return null;
  if (!Number.isFinite(at) || !Number.isFinite(count)) return null;
  if (at < 0 || at > max || count <= 0) return null;
  // A row/column insertion must leave room for every cell that is moved
  // right/down. Deletions may consume the final row/column.
  const remaining = kind === 'insert' ? max - at : max + 1 - at;
  const normalizedCount = Math.min(count, remaining);
  return normalizedCount > 0 ? { at, count: normalizedCount } : null;
}

const spanSize = (start: number, end: number): number => (end >= start ? end - start + 1 : 0);

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

function cloneInsertedFormat(format: CellFormat): CellFormat {
  const next: CellFormat = { ...format };
  delete next.hyperlink;
  delete next.hyperlinkDisplay;
  delete next.hyperlinkTooltip;
  delete next.comment;
  delete next.commentAuthor;
  delete next.validation;
  if (format.borders) next.borders = { ...format.borders };
  if (format.numFmt) next.numFmt = { ...format.numFmt };
  if (format.phonetic) next.phonetic = format.phonetic.map((run) => ({ ...run }));
  return next;
}

function parseFormatKey(key: string): Addr | null {
  const parts = key.split(':');
  if (parts.length !== 3) return null;
  const sheet = Number(parts[0]);
  const row = Number(parts[1]);
  const col = Number(parts[2]);
  if (!Number.isInteger(sheet) || !Number.isInteger(row) || !Number.isInteger(col)) return null;
  return { sheet, row, col };
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

function writeCell(wb: WorkbookHandle, addr: Addr, value: CellValue, formula: string | null): void {
  if (formula) {
    wb.setFormula(addr, formula);
    return;
  }
  switch (value.kind) {
    case 'number':
      wb.setNumber(addr, value.value);
      return;
    case 'text':
      wb.setText(addr, value.value);
      return;
    case 'bool':
      wb.setBool(addr, value.value);
      return;
    case 'error':
      wb.setError(addr, value.code);
      return;
    default:
      wb.setBlank(addr);
  }
}

/** Shift indices in a sparse Map keyed by integer index. Indices >= split move
 *  by `delta` (delta>0 = shift right/down, delta<0 = shift left/up). For
 *  delete (delta<0), keys in [split, split+|delta|) are removed. */
function shiftIndexedMap(
  src: Map<number, number>,
  split: number,
  delta: number,
  max = Number.POSITIVE_INFINITY,
): Map<number, number> {
  const out = new Map<number, number>();
  for (const [k, v] of src) {
    if (k < split) {
      out.set(k, v);
      continue;
    }
    if (delta < 0 && k < split - delta) continue; // dropped
    const next = k + delta;
    if (next < 0 || next > max) continue;
    out.set(next, v);
  }
  return out;
}

function shiftIndexedSet(
  src: Set<number>,
  split: number,
  delta: number,
  max = Number.POSITIVE_INFINITY,
): Set<number> {
  const out = new Set<number>();
  for (const k of src) {
    if (k < split) {
      out.add(k);
      continue;
    }
    if (delta < 0 && k < split - delta) continue;
    const next = k + delta;
    if (next < 0 || next > max) continue;
    out.add(next);
  }
  return out;
}

/** Shift addrKey-keyed formats so any key with row >= splitRow moves by deltaRow.
 *  When deltaRow < 0, formats in the deleted band are dropped. */
function shiftFormatsByRow(
  src: Map<string, CellFormat>,
  sheet: number,
  splitRow: number,
  deltaRow: number,
): Map<string, CellFormat> {
  const out = new Map<string, CellFormat>();
  for (const [key, fmt] of src) {
    const parts = key.split(':');
    if (parts.length !== 3) {
      out.set(key, fmt);
      continue;
    }
    const s = Number(parts[0]);
    const r = Number(parts[1]);
    const c = Number(parts[2]);
    if (s !== sheet || r < splitRow) {
      out.set(key, fmt);
      continue;
    }
    if (deltaRow < 0 && r < splitRow - deltaRow) continue; // in deleted band
    const nextRow = r + deltaRow;
    if (nextRow < 0 || nextRow > MAX_ROW) continue;
    out.set(addrKey({ sheet: s, row: nextRow, col: c }), fmt);
  }
  return out;
}

function inheritFormatsByRow(
  src: Map<string, CellFormat>,
  shifted: Map<string, CellFormat>,
  sheet: number,
  splitRow: number,
  count: number,
): Map<string, CellFormat> {
  if (splitRow <= 0 || count <= 0) return shifted;
  const out = new Map(shifted);
  for (const [key, fmt] of src) {
    const addr = parseFormatKey(key);
    if (!addr || addr.sheet !== sheet || addr.row !== splitRow - 1) continue;
    for (let row = splitRow; row < splitRow + count && row <= MAX_ROW; row += 1) {
      out.set(addrKey({ sheet, row, col: addr.col }), cloneInsertedFormat(fmt));
    }
  }
  return out;
}

function shiftFormatsByCol(
  src: Map<string, CellFormat>,
  sheet: number,
  splitCol: number,
  deltaCol: number,
): Map<string, CellFormat> {
  const out = new Map<string, CellFormat>();
  for (const [key, fmt] of src) {
    const parts = key.split(':');
    if (parts.length !== 3) {
      out.set(key, fmt);
      continue;
    }
    const s = Number(parts[0]);
    const r = Number(parts[1]);
    const c = Number(parts[2]);
    if (s !== sheet || c < splitCol) {
      out.set(key, fmt);
      continue;
    }
    if (deltaCol < 0 && c < splitCol - deltaCol) continue;
    const nextCol = c + deltaCol;
    if (nextCol < 0 || nextCol > MAX_COL) continue;
    out.set(addrKey({ sheet: s, row: r, col: nextCol }), fmt);
  }
  return out;
}

function inheritFormatsByCol(
  src: Map<string, CellFormat>,
  shifted: Map<string, CellFormat>,
  sheet: number,
  splitCol: number,
  count: number,
): Map<string, CellFormat> {
  if (splitCol <= 0 || count <= 0) return shifted;
  const out = new Map(shifted);
  for (const [key, fmt] of src) {
    const addr = parseFormatKey(key);
    if (!addr || addr.sheet !== sheet || addr.col !== splitCol - 1) continue;
    for (let col = splitCol; col < splitCol + count && col <= MAX_COL; col += 1) {
      out.set(addrKey({ sheet, row: addr.row, col }), cloneInsertedFormat(fmt));
    }
  }
  return out;
}

function applyLayoutPatch(store: SpreadsheetStore, patch: Partial<LayoutSlice>): void {
  store.setState((s) => ({ ...s, layout: { ...s.layout, ...patch } }));
}

function shiftIndexedMapWithInheritance(
  src: Map<number, number>,
  split: number,
  delta: number,
  count: number,
  max = Number.POSITIVE_INFINITY,
): Map<number, number> {
  const out = shiftIndexedMap(src, split, delta, max);
  if (split <= 0 || count <= 0) return out;
  const inherited = src.get(split - 1);
  if (inherited === undefined) return out;
  for (let index = split; index < split + count && index <= max; index += 1) {
    out.set(index, inherited);
  }
  return out;
}

/** Map a 1-D interval [lo,hi] through a row/col insert (delta>0) or delete
 *  (delta<0) at `split`. Returns null when a deletion consumes the whole
 *  interval. Inserting inside a span widens it — matching how spreadsheets grow
 *  a merge / conditional-format / filter region when rows or cols are added
 *  within it. */
function adjustInterval(
  lo: number,
  hi: number,
  split: number,
  delta: number,
): [number, number] | null {
  if (delta > 0) {
    return [lo >= split ? lo + delta : lo, hi >= split ? hi + delta : hi];
  }
  const count = -delta;
  const bandHi = split + count - 1;
  const nlo = lo < split ? lo : lo > bandHi ? lo - count : split;
  const nhi = hi < split ? hi : hi > bandHi ? hi - count : split - 1;
  return nlo > nhi ? null : [nlo, nhi];
}

/** Shift a range's row or col span for a structure edit. Returns null when a
 *  deletion removes the whole span. */
function shiftRangeAxis(
  range: Range,
  axis: 'row' | 'col',
  split: number,
  delta: number,
): Range | null {
  const lo = axis === 'row' ? range.r0 : range.c0;
  const hi = axis === 'row' ? range.r1 : range.c1;
  const res = adjustInterval(lo, hi, split, delta);
  if (!res) return null;
  const [a, b] = res;
  const max = axis === 'row' ? MAX_ROW : MAX_COL;
  if (a < 0 || b > max) return null;
  return axis === 'row' ? { ...range, r0: a, r1: b } : { ...range, c0: a, c1: b };
}

function shiftFilterCriteria(
  criteria: readonly ValueFilterCriteria[],
  sheet: number,
  axis: 'row' | 'col',
  split: number,
  delta: number,
): ValueFilterCriteria[] {
  const out: ValueFilterCriteria[] = [];
  for (const c of criteria) {
    if (c.range.sheet !== sheet) {
      out.push(c);
      continue;
    }
    const shifted = shiftRangeAxis(c.range, axis, split, delta);
    if (!shifted) continue; // filtered column removed
    let byCol = c.byCol;
    if (axis === 'col') {
      const mapped = adjustInterval(byCol, byCol, split, delta);
      if (!mapped) continue; // this column was deleted
      byCol = mapped[0];
    }
    out.push({ ...c, range: shifted, byCol });
  }
  return out;
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
        const ak = addrKey({ sheet: shifted.sheet, row: shifted.r0, col: shifted.c0 });
        byAnchor.set(ak, shifted);
        for (let row = shifted.r0; row <= shifted.r1; row += 1) {
          for (let col = shifted.c0; col <= shifted.c1; col += 1) {
            if (row === shifted.r0 && col === shifted.c0) continue;
            byCell.set(addrKey({ sheet: shifted.sheet, row, col }), ak);
          }
        }
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
  const edit = normalizeAxisEdit(atRow, count, MAX_ROW, 'insert');
  if (!edit) return false;
  const row = edit.at;
  const n = edit.count;
  const sheet = store.getState().data.sheetIndex;
  if (blockedByProtection(store, sheet, 'insertRows')) return false;
  if (insertionWouldOverflow(store, wb, sheet, 'row', row, n)) return false;
  const nativeAxisOp = wb.capabilities.insertDeleteRowsCols;

  if (history) history.begin();
  try {
    // 1. shift cells & rewrite formula refs.
    if (nativeAxisOp) {
      if (!applyAxisShiftViaEngine(wb, history, sheet, 'row', row, n)) return false;
    } else {
      applyAxisShiftToCells(wb, sheet, 'row', row, n);
    }

    // 2. shift formats.
    recordFormatChange(history, store, () => {
      store.setState((s) => ({
        ...s,
        format: {
          ...s.format,
          formats: inheritFormatsByRow(
            s.format.formats,
            shiftFormatsByRow(s.format.formats, sheet, row, n),
            sheet,
            row,
            n,
          ),
        },
      }));
    });

    // 3. shift layout (rowHeights map, hiddenRows set, freezeRows count).
    recordLayoutChange(history, store, () => {
      const before = captureLayoutSnapshot(store.getState());
      const fr = before.freezeRows > row ? before.freezeRows + n : before.freezeRows;
      applyLayoutPatch(store, {
        rowHeights: shiftIndexedMapWithInheritance(before.rowHeights, row, n, n, MAX_ROW),
        hiddenRows: shiftIndexedSet(before.hiddenRows, row, n, MAX_ROW),
        outlineRows: shiftIndexedMap(before.outlineRows, row, n, MAX_ROW),
        freezeRows: fr,
      });
    });

    // 4. re-point merges, conditional formats, and the autofilter region.
    shiftAnchoredRanges(store, wb, history, sheet, 'row', row, n, nativeAxisOp);
    return true;
  } finally {
    if (history) history.end();
  }
}

export function deleteRows(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  history: History | null,
  atRow: number,
  count = 1,
): boolean {
  const edit = normalizeAxisEdit(atRow, count, MAX_ROW, 'delete');
  if (!edit) return false;
  const row = edit.at;
  const n = edit.count;
  const sheet = store.getState().data.sheetIndex;
  if (blockedByProtection(store, sheet, 'deleteRows')) return false;
  const nativeAxisOp = wb.capabilities.insertDeleteRowsCols;

  if (history) history.begin();
  try {
    if (nativeAxisOp) {
      if (!applyAxisShiftViaEngine(wb, history, sheet, 'row', row, -n)) return false;
    } else {
      applyAxisShiftToCells(wb, sheet, 'row', row, -n);
    }

    recordFormatChange(history, store, () => {
      store.setState((s) => ({
        ...s,
        format: { ...s.format, formats: shiftFormatsByRow(s.format.formats, sheet, row, -n) },
      }));
    });

    recordLayoutChange(history, store, () => {
      const before = captureLayoutSnapshot(store.getState());
      let fr = before.freezeRows;
      if (fr > row) fr = Math.max(row, fr - n);
      applyLayoutPatch(store, {
        rowHeights: shiftIndexedMap(before.rowHeights, row, -n, MAX_ROW),
        hiddenRows: shiftIndexedSet(before.hiddenRows, row, -n, MAX_ROW),
        outlineRows: shiftIndexedMap(before.outlineRows, row, -n, MAX_ROW),
        freezeRows: fr,
      });
    });

    shiftAnchoredRanges(store, wb, history, sheet, 'row', row, -n, nativeAxisOp);
    return true;
  } finally {
    if (history) history.end();
  }
}

export function insertCols(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  history: History | null,
  atCol: number,
  count = 1,
): boolean {
  const edit = normalizeAxisEdit(atCol, count, MAX_COL, 'insert');
  if (!edit) return false;
  const col = edit.at;
  const n = edit.count;
  const sheet = store.getState().data.sheetIndex;
  if (blockedByProtection(store, sheet, 'insertCols')) return false;
  if (insertionWouldOverflow(store, wb, sheet, 'col', col, n)) return false;
  const nativeAxisOp = wb.capabilities.insertDeleteRowsCols;

  if (history) history.begin();
  try {
    if (nativeAxisOp) {
      if (!applyAxisShiftViaEngine(wb, history, sheet, 'col', col, n)) return false;
    } else {
      applyAxisShiftToCells(wb, sheet, 'col', col, n);
    }

    recordFormatChange(history, store, () => {
      store.setState((s) => ({
        ...s,
        format: {
          ...s.format,
          formats: inheritFormatsByCol(
            s.format.formats,
            shiftFormatsByCol(s.format.formats, sheet, col, n),
            sheet,
            col,
            n,
          ),
        },
      }));
    });

    recordLayoutChange(history, store, () => {
      const before = captureLayoutSnapshot(store.getState());
      const fc = before.freezeCols > col ? before.freezeCols + n : before.freezeCols;
      applyLayoutPatch(store, {
        colWidths: shiftIndexedMapWithInheritance(before.colWidths, col, n, n, MAX_COL),
        hiddenCols: shiftIndexedSet(before.hiddenCols, col, n, MAX_COL),
        outlineCols: shiftIndexedMap(before.outlineCols, col, n, MAX_COL),
        freezeCols: fc,
      });
    });

    shiftAnchoredRanges(store, wb, history, sheet, 'col', col, n, nativeAxisOp);
    return true;
  } finally {
    if (history) history.end();
  }
}

export function deleteCols(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  history: History | null,
  atCol: number,
  count = 1,
): boolean {
  const edit = normalizeAxisEdit(atCol, count, MAX_COL, 'delete');
  if (!edit) return false;
  const col = edit.at;
  const n = edit.count;
  const sheet = store.getState().data.sheetIndex;
  if (blockedByProtection(store, sheet, 'deleteCols')) return false;
  const nativeAxisOp = wb.capabilities.insertDeleteRowsCols;

  if (history) history.begin();
  try {
    if (nativeAxisOp) {
      if (!applyAxisShiftViaEngine(wb, history, sheet, 'col', col, -n)) return false;
    } else {
      applyAxisShiftToCells(wb, sheet, 'col', col, -n);
    }

    recordFormatChange(history, store, () => {
      store.setState((s) => ({
        ...s,
        format: { ...s.format, formats: shiftFormatsByCol(s.format.formats, sheet, col, -n) },
      }));
    });

    recordLayoutChange(history, store, () => {
      const before = captureLayoutSnapshot(store.getState());
      let fc = before.freezeCols;
      if (fc > col) fc = Math.max(col, fc - n);
      applyLayoutPatch(store, {
        colWidths: shiftIndexedMap(before.colWidths, col, -n, MAX_COL),
        hiddenCols: shiftIndexedSet(before.hiddenCols, col, -n, MAX_COL),
        outlineCols: shiftIndexedMap(before.outlineCols, col, -n, MAX_COL),
        freezeCols: fc,
      });
    });

    shiftAnchoredRanges(store, wb, history, sheet, 'col', col, -n, nativeAxisOp);
    return true;
  } finally {
    if (history) history.end();
  }
}

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

/** Set the per-sheet zoom level. `zoom` is a multiplier (1.0 = 100%) and is
 *  clamped by the store mutator to [0.5, 4]. When `wb` is supplied the
 *  engine receives the equivalent percentage so the value round-trips
 *  through .xlsx. Not journaled — spreadsheets treat zoom as a view setting
 *  outside the undo stack. */
export function setSheetZoom(store: SpreadsheetStore, zoom: number, wb?: WorkbookHandle): void {
  mutators.setZoom(store, zoom);
  if (wb) {
    const sheet = store.getState().data.sheetIndex;
    const pct = Math.round(store.getState().viewport.zoom * 100);
    wb.setSheetZoom(sheet, pct);
  }
}

/** Internal exports — kept narrow so the API is just the eight verbs above. */
export const __testing = {
  shiftIndexedMap,
  shiftIndexedSet,
  shiftFormatsByRow,
  shiftFormatsByCol,
  shiftFormulaRefs: adjustFormulaForRowColEdit,
};
