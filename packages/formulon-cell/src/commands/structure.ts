import type { Addr, CellValue } from '../engine/types.js';
import { type WorkbookHandle, addrKey } from '../engine/workbook-handle.js';
import {
  type CellFormat,
  type LayoutSlice,
  type SpreadsheetStore,
  mutators,
} from '../store/store.js';
import {
  type History,
  captureLayoutSnapshot,
  recordFormatChange,
  recordLayoutChange,
  recordLayoutChangeWithEngine,
} from './history.js';

interface CellRecord {
  addr: Addr;
  value: CellValue;
  formula: string | null;
}

const MAX_ROW = 1048575;
const MAX_COL = 16383;

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

    const newFormula = c.formula ? shiftFormulaRefs(c.formula, axis, split, delta) : null;

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

/** Convert an A1 col label to a 0-indexed column. "A" → 0, "Z" → 25, "AA" → 26. */
function colLabelToIndex(label: string): number {
  let n = 0;
  for (let i = 0; i < label.length; i += 1) {
    n = n * 26 + (label.charCodeAt(i) - 64);
  }
  return n - 1;
}

/** Convert a 0-indexed column to its A1 label. */
function colIndexToLabel(col: number): string {
  let n = col;
  let out = '';
  do {
    out = String.fromCharCode(65 + (n % 26)) + out;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return out;
}

/**
 * Rewrite cell references in `src` for a row or column shift on the active
 * sheet. Absolute references (`$A$1`, `A$1`, `$A1`) keep their pinned axis.
 * References that fall inside a deleted band become `#REF!` per Excel's
 * convention.
 *
 * String literals (`"..."`) and `#REF!` tokens are left untouched. Sheet
 * qualifiers (`Sheet2!A1`, `'X'!A1`) are preserved by tokenization — the
 * shift only rewrites the trailing A1 part since refs inside the body of
 * another sheet still need to follow the current sheet's geometry; for v0.9
 * the stub doesn't model cross-sheet refs anyway.
 */
function shiftFormulaRefs(src: string, axis: 'row' | 'col', split: number, delta: number): string {
  if (delta === 0) return src;
  let out = '';
  let i = 0;

  while (i < src.length) {
    const ch = src[i] ?? '';

    // Pass through string literals untouched. Excel uses `""` to escape a
    // double quote inside a string.
    if (ch === '"') {
      out += ch;
      i += 1;
      while (i < src.length) {
        const c = src[i] ?? '';
        out += c;
        i += 1;
        if (c === '"') {
          if (src[i] === '"') {
            out += '"';
            i += 1;
            continue;
          }
          break;
        }
      }
      continue;
    }

    // Try to match a cell reference at this position.
    // Pattern: optional `$`, letters, optional `$`, digits.
    const refMatch = matchRef(src, i);
    if (refMatch) {
      const { absCol, label, absRow, rowStr, end } = refMatch;
      // Skip if previous char makes this look like part of an identifier
      // (e.g. function name "SIN10" — won't happen for valid Excel, but be
      // defensive; underscores aren't allowed in refs).
      const prev = i > 0 ? src[i - 1] : '';
      if (prev && /[A-Za-z0-9_]/.test(prev)) {
        out += ch;
        i += 1;
        continue;
      }
      const col = colLabelToIndex(label);
      const row = Number.parseInt(rowStr, 10) - 1;

      let nextCol = col;
      let nextRow = row;
      let invalid = false;

      if (axis === 'row' && !absRow) {
        if (row >= split) {
          if (delta < 0 && row < split - delta) invalid = true;
          else nextRow = row + delta;
        }
      } else if (axis === 'col' && !absCol) {
        if (col >= split) {
          if (delta < 0 && col < split - delta) invalid = true;
          else nextCol = col + delta;
        }
      }

      if (invalid) out += '#REF!';
      else
        out += `${absCol ? '$' : ''}${colIndexToLabel(nextCol)}${absRow ? '$' : ''}${nextRow + 1}`;

      i = end;
      continue;
    }

    out += ch;
    i += 1;
  }

  return out;
}

interface RefMatch {
  absCol: boolean;
  label: string;
  absRow: boolean;
  rowStr: string;
  end: number;
}

function matchRef(src: string, start: number): RefMatch | null {
  let i = start;
  let absCol = false;
  if (src[i] === '$') {
    absCol = true;
    i += 1;
  }
  const lettersStart = i;
  while (i < src.length) {
    const c = src[i] ?? '';
    if (c >= 'A' && c <= 'Z') i += 1;
    else if (c >= 'a' && c <= 'z') i += 1;
    else break;
  }
  if (i === lettersStart) return null;
  const label = src.slice(lettersStart, i).toUpperCase();
  let absRow = false;
  if (src[i] === '$') {
    absRow = true;
    i += 1;
  }
  const digitsStart = i;
  while (i < src.length) {
    const c = src[i] ?? '';
    if (c >= '0' && c <= '9') i += 1;
    else break;
  }
  if (i === digitsStart) return null;
  // Reject if followed by `(` — that's a function call, not a reference.
  // (`A1()` is invalid Excel anyway, but be safe.)
  if (src[i] === '(') return null;
  return { absCol, label, absRow, rowStr: src.slice(digitsStart, i), end: i };
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
        format: { formats: shiftFormatsByRow(s.format.formats, sheet, atRow, count) },
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
        format: { formats: shiftFormatsByRow(s.format.formats, sheet, atRow, -n) },
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
        format: { formats: shiftFormatsByCol(s.format.formats, sheet, atCol, count) },
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
        format: { formats: shiftFormatsByCol(s.format.formats, sheet, atCol, -n) },
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
  recordLayoutChangeWithEngine(history, store, wb ?? null, () => {
    store.setState((s) => {
      const next = new Set(s.layout.hiddenRows);
      for (let r = r0; r <= r1; r += 1) next.delete(r);
      return { ...s, layout: { ...s.layout, hiddenRows: next } };
    });
  });
}

export function hideCols(
  store: SpreadsheetStore,
  history: History | null,
  c0: number,
  c1: number,
  wb?: WorkbookHandle,
): void {
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
  recordLayoutChangeWithEngine(history, store, wb ?? null, () => {
    store.setState((s) => {
      const next = new Set(s.layout.hiddenCols);
      for (let c = c0; c <= c1; c += 1) next.delete(c);
      return { ...s, layout: { ...s.layout, hiddenCols: next } };
    });
  });
}

/** Resolve which row/col indices to show again from the current selection.
 *  Excel returns visible rows that flank a hidden band; we emulate by
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
  for (let i = lo; i <= hi; i += 1) if (set.has(i)) out.push(i);
  return out;
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
 *  through .xlsx. Not journaled — Excel treats zoom as a view setting
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
  shiftFormulaRefs,
};
