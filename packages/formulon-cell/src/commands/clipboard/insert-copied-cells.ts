import { addrKey } from '../../engine/address.js';
import type { Addr, CellValue, Range } from '../../engine/types.js';
import type { WorkbookHandle } from '../../engine/workbook-handle.js';
import type { CellFormat, SpreadsheetStore, State } from '../../store/store.js';
import { coerceInputForCell, writeCoerced } from '../coerce-input.js';
import { adjustFormulaForCellBandShift } from '../formula-refs.js';
import { type History, recordFormatChange, recordMergesChangeWithEngine } from '../history.js';
import { isCellWritable, isSheetProtected } from '../protection.js';
import { shiftFormulaRefs } from '../refs.js';
import type { ClipboardSnapshot } from './snapshot.js';
import { parseTSV } from './tsv.js';

export type InsertCopiedCellsDirection = 'right' | 'down';

export interface InsertCopiedCellsResult {
  writtenRange: Range;
}

interface CellRecord {
  addr: Addr;
  value: CellValue;
  formula: string | null;
}

const MAX_ROW = 1048575;
const MAX_COL = 16383;
const MAX_INSERT_COPIED_CELLS = 100_000;

export function insertCopiedCellsFromTSV(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  history: History | null,
  text: string,
  direction: InsertCopiedCellsDirection,
  snapshot?: ClipboardSnapshot | null,
): InsertCopiedCellsResult | null {
  if (!text && !snapshot) return null;
  const rows = snapshot ? [] : parseTSV(text);
  if (!snapshot && rows.length === 0) return null;
  const height = snapshot?.rows ?? rows.length;
  const width = snapshot?.cols ?? rows.reduce((max, row) => Math.max(max, row.length), 0);
  if (width <= 0) return null;
  if (height * width > MAX_INSERT_COPIED_CELLS) return null;

  const state = store.getState();
  const origin = state.selection.active;
  const sheet = origin.sheet;
  if (isSheetProtected(state, sheet)) {
    // eslint-disable-next-line no-console
    console.warn(`formulon-cell: insert copied cells blocked — sheet ${sheet} is protected`);
    return null;
  }

  const affected: Range =
    direction === 'down'
      ? {
          sheet,
          r0: origin.row,
          c0: origin.col,
          r1: MAX_ROW,
          c1: Math.min(MAX_COL, origin.col + width - 1),
        }
      : {
          sheet,
          r0: origin.row,
          c0: origin.col,
          r1: Math.min(MAX_ROW, origin.row + height - 1),
          c1: MAX_COL,
        };
  if (!canShiftMerges(state, affected, direction)) return null;

  if (history) history.begin();
  try {
    shiftCells(store.getState(), wb, affected, direction, direction === 'down' ? height : width);
    shiftFormats(store, history, affected, direction, direction === 'down' ? height : width);
    shiftMerges(store, wb, history, affected, direction, direction === 'down' ? height : width);

    if (snapshot) {
      writeSnapshotIntoInsertedRange(store, wb, history, snapshot, origin);
    } else {
      for (let r = 0; r < rows.length; r += 1) {
        const cells = rows[r] ?? [];
        for (let c = 0; c < cells.length; c += 1) {
          const addr: Addr = { sheet, row: origin.row + r, col: origin.col + c };
          if (!isCellWritable(store.getState(), addr)) continue;
          writeCoerced(wb, addr, coerceInputForCell(store.getState(), addr, cells[c] ?? ''));
        }
      }
    }
    copySourceMerges(store, wb, history, origin, height, width);
    wb.recalc();
  } finally {
    if (history) history.end();
  }

  return {
    writtenRange: {
      sheet,
      r0: origin.row,
      c0: origin.col,
      r1: origin.row + height - 1,
      c1: origin.col + width - 1,
    },
  };
}

function writeSnapshotIntoInsertedRange(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  history: History | null,
  snapshot: ClipboardSnapshot,
  origin: Addr,
): void {
  const formatWrites: { key: string; format: CellFormat | null }[] = [];
  for (let r = 0; r < snapshot.rows; r += 1) {
    for (let c = 0; c < snapshot.cols; c += 1) {
      const src = snapshot.cells[r]?.[c];
      if (!src) continue;
      const addr: Addr = { sheet: origin.sheet, row: origin.row + r, col: origin.col + c };
      if (!isCellWritable(store.getState(), addr)) continue;
      if (src.formula) {
        const formula =
          snapshot.mode === 'cut'
            ? src.formula
            : shiftFormulaRefs(
                src.formula,
                addr.row - (snapshot.range.r0 + r),
                addr.col - (snapshot.range.c0 + c),
              );
        wb.setFormula(addr, formula);
      } else {
        writeCell(wb, addr, src.value, null);
      }
      formatWrites.push({
        key: addrKey(addr),
        format: src.format
          ? { ...src.format, borders: src.format.borders ? { ...src.format.borders } : undefined }
          : null,
      });
    }
  }
  if (formatWrites.length === 0) return;
  recordFormatChange(history, store, () => {
    store.setState((s) => {
      const formats = new Map(s.format.formats);
      for (const { key, format } of formatWrites) {
        if (format) formats.set(key, format);
        else formats.delete(key);
      }
      return { ...s, format: { ...s.format, formats } };
    });
  });
}

function shiftCells(
  _state: State,
  wb: WorkbookHandle,
  affected: Range,
  direction: InsertCopiedCellsDirection,
  delta: number,
): void {
  const all = Array.from(wb.cells(affected.sheet)).map<CellRecord>((c) => ({
    addr: c.addr,
    value: c.value,
    formula: c.formula,
  }));
  const moving = all.filter((cell) => inRange(cell.addr, affected));
  for (const cell of moving) wb.setBlank(cell.addr);

  const sorted =
    direction === 'down'
      ? moving.sort((a, b) => b.addr.row - a.addr.row)
      : moving.sort((a, b) => b.addr.col - a.addr.col);
  for (const cell of sorted) {
    const next: Addr =
      direction === 'down'
        ? { ...cell.addr, row: cell.addr.row + delta }
        : { ...cell.addr, col: cell.addr.col + delta };
    if (next.row > MAX_ROW || next.col > MAX_COL) continue;
    const formula = cell.formula
      ? adjustFormulaForCellBandShift(cell.formula, affected, direction, delta)
      : null;
    writeCell(wb, next, cell.value, formula);
  }

  for (const cell of all) {
    if (cell.formula && !inRange(cell.addr, affected)) {
      const nextFormula = adjustFormulaForCellBandShift(cell.formula, affected, direction, delta);
      if (nextFormula !== cell.formula) wb.setFormula(cell.addr, nextFormula);
    }
  }
}

function shiftFormats(
  store: SpreadsheetStore,
  history: History | null,
  affected: Range,
  direction: InsertCopiedCellsDirection,
  delta: number,
): void {
  recordFormatChange(history, store, () => {
    store.setState((s) => ({
      ...s,
      format: {
        ...s.format,
        formats: shiftFormatMap(s.format.formats, affected, direction, delta),
      },
    }));
  });
}

function shiftMerges(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  history: History | null,
  affected: Range,
  direction: InsertCopiedCellsDirection,
  delta: number,
): void {
  recordMergesChangeWithEngine(history, store, wb, affected.sheet, () => {
    store.setState((s) => {
      const byAnchor = new Map<string, Range>();
      const byCell = new Map<string, string>();
      for (const merge of s.merges.byAnchor.values()) {
        const shifted =
          merge.sheet === affected.sheet && mergeIntersectsShiftBand(merge, affected, direction)
            ? shiftRange(merge, direction, delta)
            : merge;
        addMergeToMaps(byAnchor, byCell, shifted);
      }
      return { ...s, merges: { byAnchor, byCell } };
    });
  });
}

function copySourceMerges(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  history: History | null,
  origin: Addr,
  height: number,
  width: number,
): void {
  const state = store.getState();
  const source = state.ui.copyRange;
  if (!source || source.sheet !== origin.sheet) return;
  const copied: Range = {
    sheet: source.sheet,
    r0: source.r0,
    c0: source.c0,
    r1: Math.min(source.r0 + height - 1, source.r1),
    c1: Math.min(source.c0 + width - 1, source.c1),
  };
  const sourceMerges = Array.from(state.merges.byAnchor.values()).filter(
    (m) =>
      m.sheet === copied.sheet &&
      m.r0 >= copied.r0 &&
      m.r1 <= copied.r1 &&
      m.c0 >= copied.c0 &&
      m.c1 <= copied.c1,
  );
  if (sourceMerges.length === 0) return;
  recordMergesChangeWithEngine(history, store, wb, origin.sheet, () => {
    store.setState((s) => {
      const byAnchor = new Map(s.merges.byAnchor);
      const byCell = new Map(s.merges.byCell);
      for (const merge of sourceMerges) {
        const next: Range = {
          sheet: origin.sheet,
          r0: origin.row + (merge.r0 - copied.r0),
          c0: origin.col + (merge.c0 - copied.c0),
          r1: origin.row + (merge.r1 - copied.r0),
          c1: origin.col + (merge.c1 - copied.c0),
        };
        removeIntersectingMerges(byAnchor, byCell, next);
        addMergeToMaps(byAnchor, byCell, next);
      }
      return { ...s, merges: { byAnchor, byCell } };
    });
  });
}

function canShiftMerges(
  state: State,
  affected: Range,
  direction: InsertCopiedCellsDirection,
): boolean {
  for (const merge of state.merges.byAnchor.values()) {
    if (merge.sheet !== affected.sheet) continue;
    if (!rangesIntersect(merge, affected)) continue;
    if (!mergeIntersectsShiftBand(merge, affected, direction)) continue;
    const fullyInsideBand =
      direction === 'down'
        ? merge.c0 >= affected.c0 && merge.c1 <= affected.c1 && merge.r0 >= affected.r0
        : merge.r0 >= affected.r0 && merge.r1 <= affected.r1 && merge.c0 >= affected.c0;
    if (!fullyInsideBand) {
      // eslint-disable-next-line no-console
      console.warn('formulon-cell: insert copied cells blocked — merge would be split');
      return false;
    }
  }
  return true;
}

function mergeIntersectsShiftBand(
  merge: Range,
  affected: Range,
  direction: InsertCopiedCellsDirection,
): boolean {
  return direction === 'down'
    ? merge.r1 >= affected.r0 && merge.c1 >= affected.c0 && merge.c0 <= affected.c1
    : merge.c1 >= affected.c0 && merge.r1 >= affected.r0 && merge.r0 <= affected.r1;
}

function shiftFormatMap(
  formats: Map<string, CellFormat>,
  affected: Range,
  direction: InsertCopiedCellsDirection,
  delta: number,
): Map<string, CellFormat> {
  const next = new Map<string, CellFormat>();
  for (const [key, fmt] of formats) {
    const parts = key.split(':');
    if (parts.length !== 3) {
      next.set(key, fmt);
      continue;
    }
    const addr: Addr = { sheet: Number(parts[0]), row: Number(parts[1]), col: Number(parts[2]) };
    if (!inRange(addr, affected)) {
      next.set(key, fmt);
      continue;
    }
    const shifted =
      direction === 'down'
        ? { ...addr, row: addr.row + delta }
        : { ...addr, col: addr.col + delta };
    if (shifted.row <= MAX_ROW && shifted.col <= MAX_COL) next.set(addrKey(shifted), fmt);
  }
  return next;
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

function inRange(addr: Addr, range: Range): boolean {
  return (
    addr.sheet === range.sheet &&
    addr.row >= range.r0 &&
    addr.row <= range.r1 &&
    addr.col >= range.c0 &&
    addr.col <= range.c1
  );
}

function rangesIntersect(a: Range, b: Range): boolean {
  return a.sheet === b.sheet && !(a.r1 < b.r0 || a.r0 > b.r1 || a.c1 < b.c0 || a.c0 > b.c1);
}

function shiftRange(range: Range, direction: InsertCopiedCellsDirection, delta: number): Range {
  return direction === 'down'
    ? { ...range, r0: range.r0 + delta, r1: range.r1 + delta }
    : { ...range, c0: range.c0 + delta, c1: range.c1 + delta };
}

function addMergeToMaps(byAnchor: Map<string, Range>, byCell: Map<string, string>, range: Range) {
  const ak = addrKey({ sheet: range.sheet, row: range.r0, col: range.c0 });
  byAnchor.set(ak, range);
  for (let row = range.r0; row <= range.r1; row += 1) {
    for (let col = range.c0; col <= range.c1; col += 1) {
      if (row === range.r0 && col === range.c0) continue;
      byCell.set(addrKey({ sheet: range.sheet, row, col }), ak);
    }
  }
}

function removeIntersectingMerges(
  byAnchor: Map<string, Range>,
  byCell: Map<string, string>,
  range: Range,
): void {
  for (const [anchorKey, merge] of byAnchor) {
    if (!rangesIntersect(merge, range)) continue;
    byAnchor.delete(anchorKey);
    for (let row = merge.r0; row <= merge.r1; row += 1) {
      for (let col = merge.c0; col <= merge.c1; col += 1) {
        byCell.delete(addrKey({ sheet: merge.sheet, row, col }));
      }
    }
  }
}

// Reference shifting delegates to the shared `adjustFormulaForCellBandShift`
// (commands/formula-refs.ts) so sheet qualifiers, function names, ranges, and
// grid bounds are handled the same way as the other structural edits.
