import { addrKey } from '../engine/address.js';
import type { Addr, CellValue, Range } from '../engine/types.js';
import { formatCell } from '../engine/value.js';
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
import { recordFilterChange } from './filter.js';
import { formatNumber } from './format.js';
import { adjustFormulaForRowColEdit } from './formula-refs.js';
import {
  captureLayoutSnapshot,
  type History,
  recordConditionalRulesChange,
  recordFormatChange,
  recordLayoutChange,
  recordLayoutChangeWithEngine,
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

const spanSize = (start: number, end: number): number => (end >= start ? end - start + 1 : 0);

function collectAllCells(wb: WorkbookHandle, sheet: number): CellRecord[] {
  const out: CellRecord[] = [];
  for (const c of wb.cells(sheet)) {
    out.push({ addr: c.addr, value: c.value, formula: c.formula });
  }
  return out;
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
): Map<number, number> {
  const out = new Map<number, number>();
  for (const [k, v] of src) {
    if (k < split) {
      out.set(k, v);
      continue;
    }
    if (delta < 0 && k < split - delta) continue; // dropped
    out.set(k + delta, v);
  }
  return out;
}

function shiftIndexedSet(src: Set<number>, split: number, delta: number): Set<number> {
  const out = new Set<number>();
  for (const k of src) {
    if (k < split) {
      out.add(k);
      continue;
    }
    if (delta < 0 && k < split - delta) continue;
    out.add(k + delta);
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
    out.set(addrKey({ sheet: s, row: r + deltaRow, col: c }), fmt);
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
    out.set(addrKey({ sheet: s, row: r, col: c + deltaCol }), fmt);
  }
  return out;
}

function applyLayoutPatch(store: SpreadsheetStore, patch: Partial<LayoutSlice>): void {
  store.setState((s) => ({ ...s, layout: { ...s.layout, ...patch } }));
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
): void {
  // Merges: rebuild both lookup maps from the shifted anchors. A merge fully
  // inside a deleted band, or collapsed to a single cell, is dropped.
  recordMergesChangeWithEngine(history, store, wb, sheet, () => {
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
  });

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
): void {
  if (delta === 0) return;
  const positive = delta > 0;
  const count = Math.abs(delta);

  // For delete (delta < 0), capture cells in the band so undo can rewrite
  // them. Insert is fully invertible by the opposite engine op alone.
  const captured: CellRecord[] = [];
  if (!positive) {
    for (const c of wb.cells(sheet)) {
      const k = axis === 'row' ? c.addr.row : c.addr.col;
      if (k >= split && k < split + count) {
        captured.push({ addr: c.addr, value: c.value, formula: c.formula });
      }
    }
  }

  const apply = (): void => {
    if (positive) {
      if (axis === 'row') wb.engineInsertRows(sheet, split, count);
      else wb.engineInsertCols(sheet, split, count);
    } else {
      if (axis === 'row') wb.engineDeleteRows(sheet, split, count);
      else wb.engineDeleteCols(sheet, split, count);
    }
    wb.recalc();
  };
  const invert = (): void => {
    if (positive) {
      if (axis === 'row') wb.engineDeleteRows(sheet, split, count);
      else wb.engineDeleteCols(sheet, split, count);
    } else {
      if (axis === 'row') wb.engineInsertRows(sheet, split, count);
      else wb.engineInsertCols(sheet, split, count);
      // Restore captured cells. wb.setX runs through the per-cell journal,
      // but History.replaying short-circuits the push so no extra entries
      // are recorded.
      for (const c of captured) writeCell(wb, c.addr, c.value, c.formula);
    }
    wb.recalc();
  };

  apply();
  if (history && !history.isReplaying()) {
    history.push({ undo: invert, redo: apply });
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
): void {
  if (count <= 0) return;
  const sheet = store.getState().data.sheetIndex;
  if (blockedByProtection(store, sheet, 'insertRows')) return;

  if (history) history.begin();
  try {
    // 1. shift cells & rewrite formula refs.
    if (wb.capabilities.insertDeleteRowsCols) {
      applyAxisShiftViaEngine(wb, history, sheet, 'row', atRow, count);
    } else {
      applyAxisShiftToCells(wb, sheet, 'row', atRow, count);
      // Per-cell setNumber/setText skip recalc; force one pass at the end so
      // formulas in the shifted band see their (possibly already-restored)
      // operands.
      wb.recalc();
    }

    // 2. shift formats.
    recordFormatChange(history, store, () => {
      store.setState((s) => ({
        ...s,
        format: { ...s.format, formats: shiftFormatsByRow(s.format.formats, sheet, atRow, count) },
      }));
    });

    // 3. shift layout (rowHeights map, hiddenRows set, freezeRows count).
    recordLayoutChange(history, store, () => {
      const before = captureLayoutSnapshot(store.getState());
      const fr = before.freezeRows > atRow ? before.freezeRows + count : before.freezeRows;
      applyLayoutPatch(store, {
        rowHeights: shiftIndexedMap(before.rowHeights, atRow, count),
        hiddenRows: shiftIndexedSet(before.hiddenRows, atRow, count),
        outlineRows: shiftIndexedMap(before.outlineRows, atRow, count),
        freezeRows: fr,
      });
    });

    // 4. re-point merges, conditional formats, and the autofilter region.
    shiftAnchoredRanges(store, wb, history, sheet, 'row', atRow, count);
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
): void {
  if (count <= 0) return;
  const sheet = store.getState().data.sheetIndex;
  if (blockedByProtection(store, sheet, 'deleteRows')) return;
  // Cap count so we don't try to delete past MAX_ROW.
  const n = Math.min(count, MAX_ROW + 1 - atRow);
  if (n <= 0) return;

  if (history) history.begin();
  try {
    if (wb.capabilities.insertDeleteRowsCols) {
      applyAxisShiftViaEngine(wb, history, sheet, 'row', atRow, -n);
    } else {
      applyAxisShiftToCells(wb, sheet, 'row', atRow, -n);
      wb.recalc();
    }

    recordFormatChange(history, store, () => {
      store.setState((s) => ({
        ...s,
        format: { ...s.format, formats: shiftFormatsByRow(s.format.formats, sheet, atRow, -n) },
      }));
    });

    recordLayoutChange(history, store, () => {
      const before = captureLayoutSnapshot(store.getState());
      let fr = before.freezeRows;
      if (fr > atRow) fr = Math.max(atRow, fr - n);
      applyLayoutPatch(store, {
        rowHeights: shiftIndexedMap(before.rowHeights, atRow, -n),
        hiddenRows: shiftIndexedSet(before.hiddenRows, atRow, -n),
        outlineRows: shiftIndexedMap(before.outlineRows, atRow, -n),
        freezeRows: fr,
      });
    });

    shiftAnchoredRanges(store, wb, history, sheet, 'row', atRow, -n);
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
): void {
  if (count <= 0) return;
  const sheet = store.getState().data.sheetIndex;
  if (blockedByProtection(store, sheet, 'insertCols')) return;

  if (history) history.begin();
  try {
    if (wb.capabilities.insertDeleteRowsCols) {
      applyAxisShiftViaEngine(wb, history, sheet, 'col', atCol, count);
    } else {
      applyAxisShiftToCells(wb, sheet, 'col', atCol, count);
      wb.recalc();
    }

    recordFormatChange(history, store, () => {
      store.setState((s) => ({
        ...s,
        format: { ...s.format, formats: shiftFormatsByCol(s.format.formats, sheet, atCol, count) },
      }));
    });

    recordLayoutChange(history, store, () => {
      const before = captureLayoutSnapshot(store.getState());
      const fc = before.freezeCols > atCol ? before.freezeCols + count : before.freezeCols;
      applyLayoutPatch(store, {
        colWidths: shiftIndexedMap(before.colWidths, atCol, count),
        hiddenCols: shiftIndexedSet(before.hiddenCols, atCol, count),
        outlineCols: shiftIndexedMap(before.outlineCols, atCol, count),
        freezeCols: fc,
      });
    });

    shiftAnchoredRanges(store, wb, history, sheet, 'col', atCol, count);
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
): void {
  if (count <= 0) return;
  const sheet = store.getState().data.sheetIndex;
  if (blockedByProtection(store, sheet, 'deleteCols')) return;
  const n = Math.min(count, MAX_COL + 1 - atCol);
  if (n <= 0) return;

  if (history) history.begin();
  try {
    if (wb.capabilities.insertDeleteRowsCols) {
      applyAxisShiftViaEngine(wb, history, sheet, 'col', atCol, -n);
    } else {
      applyAxisShiftToCells(wb, sheet, 'col', atCol, -n);
      wb.recalc();
    }

    recordFormatChange(history, store, () => {
      store.setState((s) => ({
        ...s,
        format: { ...s.format, formats: shiftFormatsByCol(s.format.formats, sheet, atCol, -n) },
      }));
    });

    recordLayoutChange(history, store, () => {
      const before = captureLayoutSnapshot(store.getState());
      let fc = before.freezeCols;
      if (fc > atCol) fc = Math.max(atCol, fc - n);
      applyLayoutPatch(store, {
        colWidths: shiftIndexedMap(before.colWidths, atCol, -n),
        hiddenCols: shiftIndexedSet(before.hiddenCols, atCol, -n),
        outlineCols: shiftIndexedMap(before.outlineCols, atCol, -n),
        freezeCols: fc,
      });
    });

    shiftAnchoredRanges(store, wb, history, sheet, 'col', atCol, -n);
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
): void {
  if (spanSize(r0, r1) > MAX_MATERIALIZED_LAYOUT_ROWS) return;
  recordLayoutChangeWithEngine(history, store, wb ?? null, () => {
    const ctx = createAutofitMeasureContext();
    for (let row = r0; row <= r1; row += 1) {
      mutators.setRowHeight(store, row, computeAutofitRowHeight(store.getState(), row, ctx));
    }
  });
}

export function autofitColsWidth(
  store: SpreadsheetStore,
  history: History | null,
  c0: number,
  c1: number,
  wb?: WorkbookHandle,
): void {
  recordLayoutChangeWithEngine(history, store, wb ?? null, () => {
    const ctx = createAutofitMeasureContext();
    for (let col = c0; col <= c1; col += 1) {
      mutators.setColWidth(store, col, computeAutofitColWidth(store.getState(), col, ctx));
    }
  });
}

function createAutofitMeasureContext(): CanvasRenderingContext2D | null {
  const doc = globalThis.document;
  const canvas = doc?.createElement?.('canvas');
  return canvas?.getContext?.('2d') ?? null;
}

function computeAutofitColWidth(
  state: State,
  col: number,
  ctx: CanvasRenderingContext2D | null,
): number {
  const sheet = state.data.sheetIndex;
  const padding = 16;
  const minWidth = 48;
  let max = 0;

  for (const [key, cell] of state.data.cells) {
    const parsed = parseCellKey(key);
    if (!parsed || parsed.sheet !== sheet || parsed.col !== col) continue;
    const text = autofitDisplayText(state, key, cell);
    if (!text) continue;
    const fmt = state.format.formats.get(key);
    const fontSize = fmt?.fontSize ?? 13;
    if (ctx) ctx.font = autofitFont(fmt);
    const width =
      maxExplicitLineWidth(text, ctx, fontSize) +
      (isFilterHeaderCell(state, parsed.sheet, parsed.row, parsed.col) ? 28 : 0);
    if (width > max) max = width;
  }

  return Math.max(minWidth, Math.ceil(max) + padding);
}

function computeAutofitRowHeight(
  state: State,
  row: number,
  ctx: CanvasRenderingContext2D | null,
): number {
  const sheet = state.data.sheetIndex;
  let max = state.layout.defaultRowHeight;

  for (const [key, cell] of state.data.cells) {
    const parsed = parseCellKey(key);
    if (!parsed || parsed.sheet !== sheet || parsed.row !== row) continue;
    const text = autofitDisplayText(state, key, cell);
    if (!text) continue;
    const fmt = state.format.formats.get(key);
    const fontSize = fmt?.fontSize ?? 13;
    if (ctx) ctx.font = autofitFont(fmt);
    const lineHeight = Math.round(fontSize * 1.28);
    const colW = state.layout.colWidths.get(parsed.col) ?? state.layout.defaultColWidth;
    const lines = autofitLineCount(text, fmt?.wrap === true, colW, ctx, fontSize);
    const height = Math.ceil(lines * lineHeight + 8);
    if (height > max) max = height;
  }

  return max;
}

function parseCellKey(key: string): { sheet: number; row: number; col: number } | null {
  const parts = key.split(':');
  if (parts.length !== 3) return null;
  const sheet = Number(parts[0]);
  const row = Number(parts[1]);
  const col = Number(parts[2]);
  if (!Number.isInteger(sheet) || !Number.isInteger(row) || !Number.isInteger(col)) return null;
  return { sheet, row, col };
}

function autofitDisplayText(
  state: State,
  key: string,
  cell: { value: CellValue; formula: string | null },
): string {
  if (state.ui.showFormulas && cell.formula) return cell.formula;
  const fmt = state.format.formats.get(key);
  if (cell.value.kind === 'number' && fmt?.numFmt)
    return formatNumber(cell.value.value, fmt.numFmt);
  return formatCell(cell.value);
}

function isFilterHeaderCell(state: State, sheet: number, row: number, col: number): boolean {
  const fr = state.ui.filterRange;
  return !!fr && fr.sheet === sheet && fr.r0 === row && col >= fr.c0 && col <= fr.c1;
}

function autofitFont(format: CellFormat | undefined): string {
  const styleSlant = format?.italic ? 'italic ' : '';
  const weight = format?.bold ? 700 : 400;
  const size = format?.fontSize ?? 13;
  const family = format?.fontFamily ?? 'system-ui, sans-serif';
  return `${styleSlant}${weight} ${size}px ${fontFamilyCss(family)}`;
}

function fontFamilyCss(family: string): string {
  return family
    .split(',')
    .map((part) => {
      const trimmed = part.trim();
      if (/^["'].*["']$/.test(trimmed) || /^[a-z-]+$/i.test(trimmed)) return trimmed;
      return `"${trimmed.replace(/"/g, '\\"')}"`;
    })
    .join(', ');
}

function maxExplicitLineWidth(
  text: string,
  ctx: CanvasRenderingContext2D | null,
  fontSize: number,
): number {
  let max = 0;
  for (const line of text.split(/\r\n|\r|\n/)) {
    const measured = ctx ? ctx.measureText(line).width : 0;
    const width = measured > 0 ? measured : line.length * fontSize * 0.54;
    if (width > max) max = width;
  }
  return max;
}

function autofitLineCount(
  text: string,
  wrap: boolean,
  colWidth: number,
  ctx: CanvasRenderingContext2D | null,
  fontSize: number,
): number {
  const paragraphs = text.split(/\r\n|\r|\n/);
  if (!wrap) return Math.max(1, paragraphs.length);

  const available = Math.max(1, colWidth - 12);
  let lines = 0;
  for (const paragraph of paragraphs) {
    if (paragraph.length === 0) {
      lines += 1;
      continue;
    }
    lines += wrapAutofitParagraph(paragraph, available, ctx, fontSize);
  }
  return Math.max(1, lines);
}

function wrapAutofitParagraph(
  paragraph: string,
  maxWidth: number,
  ctx: CanvasRenderingContext2D | null,
  fontSize: number,
): number {
  const words = paragraph.split(/(\s+)/);
  let line = '';
  let count = 0;
  for (const word of words) {
    const candidate = line + word;
    if (measureAutofitText(candidate, ctx, fontSize) <= maxWidth || line === '') {
      line = candidate;
    } else {
      count += 1;
      line = word.trimStart();
    }
  }
  return count + (line ? 1 : 0);
}

function measureAutofitText(
  text: string,
  ctx: CanvasRenderingContext2D | null,
  fontSize: number,
): number {
  const measured = ctx ? ctx.measureText(text).width : 0;
  return measured > 0 ? measured : text.length * fontSize * 0.54;
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
