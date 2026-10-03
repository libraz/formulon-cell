import { addrKey } from '../../engine/address.js';
import type { Addr, CellValue, Range } from '../../engine/types.js';
import type { WorkbookHandle } from '../../engine/workbook-handle.js';
import type { CellFormat, SpreadsheetStore, State } from '../../store/store.js';
import { adjustFormulaForCellBandShift } from '../formula-refs.js';
import { type History, recordFormatChange, recordMergesChangeWithEngine } from '../history.js';
import { isCellWritable } from '../protection.js';
import { shiftFormulaRefs } from '../refs.js';
import type { InsertCopiedCellsDirection } from './insert-copied-cells.js';
import type { ClipboardSnapshot } from './snapshot.js';

export interface CellRecord {
  addr: Addr;
  value: CellValue;
  formula: string | null;
}

export const MAX_ROW = 1048575;
export const MAX_COL = 16383;

export function writeSnapshotIntoInsertedRange(
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

export function shiftCells(
  _state: State,
  wb: WorkbookHandle,
  affected: Range,
  direction: InsertCopiedCellsDirection,
  delta: number,
): void {
  wb.withBatchedRecalc(() => writeShiftedCells(wb, affected, direction, delta));
}

function writeShiftedCells(
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

export function shiftFormats(
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

export function shiftMerges(
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

export function copySnapshotMerges(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  history: History | null,
  origin: Addr,
  snapshot: ClipboardSnapshot | null | undefined,
): void {
  if (!snapshot || snapshot.merges?.length === 0) return;
  const sourceMerges = snapshot.merges ?? [];
  recordMergesChangeWithEngine(history, store, wb, origin.sheet, () => {
    store.setState((s) => {
      const byAnchor = new Map(s.merges.byAnchor);
      const byCell = new Map(s.merges.byCell);
      for (const merge of sourceMerges) {
        if (
          !Number.isInteger(merge.r0) ||
          !Number.isInteger(merge.c0) ||
          !Number.isInteger(merge.r1) ||
          !Number.isInteger(merge.c1) ||
          merge.r0 < 0 ||
          merge.c0 < 0 ||
          merge.r1 < merge.r0 ||
          merge.c1 < merge.c0 ||
          merge.r1 >= snapshot.rows ||
          merge.c1 >= snapshot.cols ||
          (merge.sheet !== undefined && merge.sheet !== snapshot.range.sheet)
        ) {
          continue;
        }
        const next: Range = {
          sheet: origin.sheet,
          r0: origin.row + merge.r0,
          c0: origin.col + merge.c0,
          r1: origin.row + merge.r1,
          c1: origin.col + merge.c1,
        };
        if (next.r1 > MAX_ROW || next.c1 > MAX_COL) continue;
        removeIntersectingMerges(byAnchor, byCell, next);
        addMergeToMaps(byAnchor, byCell, next);
      }
      return { ...s, merges: { byAnchor, byCell } };
    });
  });
}

export function canShiftMerges(
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

export function writeCell(
  wb: WorkbookHandle,
  addr: Addr,
  value: CellValue,
  formula: string | null,
): void {
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

function inRange(addr: Addr, range: Range): boolean {
  return (
    addr.sheet === range.sheet &&
    addr.row >= range.r0 &&
    addr.row <= range.r1 &&
    addr.col >= range.c0 &&
    addr.col <= range.c1
  );
}

export function rangesIntersect(a: Range, b: Range): boolean {
  return a.sheet === b.sheet && !(a.r1 < b.r0 || a.r0 > b.r1 || a.c1 < b.c0 || a.c0 > b.c1);
}

function shiftRange(range: Range, direction: InsertCopiedCellsDirection, delta: number): Range {
  return direction === 'down'
    ? { ...range, r0: range.r0 + delta, r1: range.r1 + delta }
    : { ...range, c0: range.c0 + delta, c1: range.c1 + delta };
}

export function addMergeToMaps(
  byAnchor: Map<string, Range>,
  byCell: Map<string, string>,
  range: Range,
) {
  const ak = addrKey({ sheet: range.sheet, row: range.r0, col: range.c0 });
  byAnchor.set(ak, range);
  for (let row = range.r0; row <= range.r1; row += 1) {
    for (let col = range.c0; col <= range.c1; col += 1) {
      if (row === range.r0 && col === range.c0) continue;
      byCell.set(addrKey({ sheet: range.sheet, row, col }), ak);
    }
  }
}

export function removeIntersectingMerges(
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
