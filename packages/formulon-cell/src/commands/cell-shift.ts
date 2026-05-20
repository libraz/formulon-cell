import { addrKey } from '../engine/address.js';
import type { Addr, CellValue, Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { CellFormat, SpreadsheetStore, State } from '../store/store.js';
import { type History, recordFormatChange, recordMergesChangeWithEngine } from './history.js';
import { isSheetProtected } from './protection.js';

export type InsertCellsDirection = 'down' | 'right';
export type DeleteCellsDirection = 'up' | 'left';

interface CellRecord {
  addr: Addr;
  value: CellValue;
  formula: string | null;
}

const MAX_ROW = 1048575;
const MAX_COL = 16383;

export function insertCells(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  history: History | null,
  range: Range,
  direction: InsertCellsDirection,
): boolean {
  const sheet = range.sheet;
  if (isSheetProtected(store.getState(), sheet)) {
    // eslint-disable-next-line no-console
    console.warn(`formulon-cell: insert cells blocked — sheet ${sheet} is protected`);
    return false;
  }
  const affected: Range =
    direction === 'down'
      ? { sheet, r0: range.r0, c0: range.c0, r1: MAX_ROW, c1: range.c1 }
      : { sheet, r0: range.r0, c0: range.c0, r1: range.r1, c1: MAX_COL };
  const delta = direction === 'down' ? range.r1 - range.r0 + 1 : range.c1 - range.c0 + 1;
  return shiftCellBand(store, wb, history, affected, direction, delta);
}

export function deleteCells(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  history: History | null,
  range: Range,
  direction: DeleteCellsDirection,
): boolean {
  const sheet = range.sheet;
  if (isSheetProtected(store.getState(), sheet)) {
    // eslint-disable-next-line no-console
    console.warn(`formulon-cell: delete cells blocked — sheet ${sheet} is protected`);
    return false;
  }
  const affected: Range =
    direction === 'up'
      ? { sheet, r0: range.r0, c0: range.c0, r1: MAX_ROW, c1: range.c1 }
      : { sheet, r0: range.r0, c0: range.c0, r1: range.r1, c1: MAX_COL };
  const delta = direction === 'up' ? -(range.r1 - range.r0 + 1) : -(range.c1 - range.c0 + 1);
  return shiftCellBand(store, wb, history, affected, direction === 'up' ? 'down' : 'right', delta);
}

function shiftCellBand(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  history: History | null,
  affected: Range,
  axis: InsertCellsDirection,
  delta: number,
): boolean {
  if (delta === 0) return true;
  if (!canShiftMerges(store.getState(), affected, axis)) return false;

  if (history) history.begin();
  try {
    const beforeCells = history && !history.isReplaying() ? collectSheetCells(wb, affected.sheet) : null;
    shiftCells(wb, affected, axis, delta);
    if (history && beforeCells) {
      const afterCells = collectSheetCells(wb, affected.sheet);
      history.push({
        undo: () => restoreSheetCells(wb, affected.sheet, beforeCells),
        redo: () => restoreSheetCells(wb, affected.sheet, afterCells),
      });
    }
    shiftFormats(store, history, affected, axis, delta);
    shiftMerges(store, wb, history, affected, axis, delta);
    wb.recalc();
  } finally {
    if (history) history.end();
  }
  return true;
}

function collectSheetCells(wb: WorkbookHandle, sheet: number): CellRecord[] {
  return Array.from(wb.cells(sheet)).map<CellRecord>((c) => ({
    addr: c.addr,
    value: c.value,
    formula: c.formula,
  }));
}

function restoreSheetCells(wb: WorkbookHandle, sheet: number, cells: readonly CellRecord[]): void {
  for (const cell of Array.from(wb.cells(sheet))) wb.setBlank(cell.addr);
  for (const cell of cells) writeCell(wb, cell.addr, cell.value, cell.formula);
  wb.recalc();
}

function shiftCells(
  wb: WorkbookHandle,
  affected: Range,
  axis: InsertCellsDirection,
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
    axis === 'down'
      ? moving.sort((a, b) => (delta > 0 ? b.addr.row - a.addr.row : a.addr.row - b.addr.row))
      : moving.sort((a, b) => (delta > 0 ? b.addr.col - a.addr.col : a.addr.col - b.addr.col));

  for (const cell of sorted) {
    const next: Addr =
      axis === 'down'
        ? { ...cell.addr, row: cell.addr.row + delta }
        : { ...cell.addr, col: cell.addr.col + delta };
    if (!inShiftTarget(next, affected, axis)) continue;
    const formula = cell.formula ? shiftFormulaRefsInBand(cell.formula, affected, axis, delta) : null;
    writeCell(wb, next, cell.value, formula);
  }

  for (const cell of all) {
    if (cell.formula && !inRange(cell.addr, affected)) {
      const nextFormula = shiftFormulaRefsInBand(cell.formula, affected, axis, delta);
      if (nextFormula !== cell.formula) wb.setFormula(cell.addr, nextFormula);
    }
  }
}

function shiftFormats(
  store: SpreadsheetStore,
  history: History | null,
  affected: Range,
  axis: InsertCellsDirection,
  delta: number,
): void {
  recordFormatChange(history, store, () => {
    store.setState((s) => ({
      ...s,
      format: { ...s.format, formats: shiftFormatMap(s.format.formats, affected, axis, delta) },
    }));
  });
}

function shiftMerges(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  history: History | null,
  affected: Range,
  axis: InsertCellsDirection,
  delta: number,
): void {
  recordMergesChangeWithEngine(history, store, wb, affected.sheet, () => {
    store.setState((s) => {
      const byAnchor = new Map<string, Range>();
      const byCell = new Map<string, string>();
      for (const merge of s.merges.byAnchor.values()) {
        const shifted =
          merge.sheet === affected.sheet && mergeIntersectsShiftBand(merge, affected, axis)
            ? shiftRange(merge, axis, delta)
            : merge;
        if (shifted.r0 < 0 || shifted.c0 < 0 || shifted.r1 > MAX_ROW || shifted.c1 > MAX_COL) {
          continue;
        }
        addMergeToMaps(byAnchor, byCell, shifted);
      }
      return { ...s, merges: { byAnchor, byCell } };
    });
  });
}

function canShiftMerges(state: State, affected: Range, axis: InsertCellsDirection): boolean {
  for (const merge of state.merges.byAnchor.values()) {
    if (merge.sheet !== affected.sheet) continue;
    if (!rangesIntersect(merge, affected)) continue;
    if (!mergeIntersectsShiftBand(merge, affected, axis)) continue;
    const fullyInsideBand =
      axis === 'down'
        ? merge.c0 >= affected.c0 && merge.c1 <= affected.c1 && merge.r0 >= affected.r0
        : merge.r0 >= affected.r0 && merge.r1 <= affected.r1 && merge.c0 >= affected.c0;
    if (!fullyInsideBand) {
      // eslint-disable-next-line no-console
      console.warn('formulon-cell: cell shift blocked — merge would be split');
      return false;
    }
  }
  return true;
}

function shiftFormatMap(
  formats: Map<string, CellFormat>,
  affected: Range,
  axis: InsertCellsDirection,
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
    const shifted = axis === 'down' ? { ...addr, row: addr.row + delta } : { ...addr, col: addr.col + delta };
    if (inShiftTarget(shifted, affected, axis)) next.set(addrKey(shifted), fmt);
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

function inShiftTarget(addr: Addr, affected: Range, axis: InsertCellsDirection): boolean {
  if (addr.sheet !== affected.sheet || addr.row < 0 || addr.col < 0) return false;
  if (addr.row > MAX_ROW || addr.col > MAX_COL) return false;
  return axis === 'down'
    ? addr.row >= affected.r0 && addr.col >= affected.c0 && addr.col <= affected.c1
    : addr.col >= affected.c0 && addr.row >= affected.r0 && addr.row <= affected.r1;
}

function rangesIntersect(a: Range, b: Range): boolean {
  return a.sheet === b.sheet && !(a.r1 < b.r0 || a.r0 > b.r1 || a.c1 < b.c0 || a.c0 > b.c1);
}

function mergeIntersectsShiftBand(
  merge: Range,
  affected: Range,
  axis: InsertCellsDirection,
): boolean {
  return axis === 'down'
    ? merge.r1 >= affected.r0 && merge.c1 >= affected.c0 && merge.c0 <= affected.c1
    : merge.c1 >= affected.c0 && merge.r1 >= affected.r0 && merge.r0 <= affected.r1;
}

function shiftRange(range: Range, axis: InsertCellsDirection, delta: number): Range {
  return axis === 'down'
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

function shiftFormulaRefsInBand(
  src: string,
  affected: Range,
  axis: InsertCellsDirection,
  delta: number,
): string {
  let out = '';
  let i = 0;
  while (i < src.length) {
    if (src[i] === '"') {
      const literal = consumeStringLiteral(src, i);
      out += literal.text;
      i = literal.end;
      continue;
    }
    const m = matchRef(src, i);
    if (!m) {
      out += src[i] ?? '';
      i += 1;
      continue;
    }
    const prev = i > 0 ? src[i - 1] : '';
    if (prev && /[A-Za-z0-9_]/.test(prev)) {
      out += src[i] ?? '';
      i += 1;
      continue;
    }
    const col = colLabelToIndex(m.label);
    const row = Number.parseInt(m.rowStr, 10) - 1;
    let nextRow = row;
    let nextCol = col;
    if (axis === 'down' && !m.absRow && row >= affected.r0 && col >= affected.c0 && col <= affected.c1) {
      nextRow = row + delta;
    } else if (
      axis === 'right' &&
      !m.absCol &&
      col >= affected.c0 &&
      row >= affected.r0 &&
      row <= affected.r1
    ) {
      nextCol = col + delta;
    }
    if (nextRow < 0 || nextCol < 0) out += '#REF!';
    else out += `${m.absCol ? '$' : ''}${colIndexToLabel(nextCol)}${m.absRow ? '$' : ''}${nextRow + 1}`;
    i = m.end;
  }
  return out;
}

function consumeStringLiteral(src: string, start: number): { text: string; end: number } {
  let i = start + 1;
  while (i < src.length) {
    if (src[i] === '"') {
      if (src[i + 1] === '"') {
        i += 2;
        continue;
      }
      i += 1;
      break;
    }
    i += 1;
  }
  return { text: src.slice(start, i), end: i };
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
  while (i < src.length && /[A-Za-z]/.test(src[i] ?? '')) i += 1;
  if (i === lettersStart) return null;
  const label = src.slice(lettersStart, i).toUpperCase();
  let absRow = false;
  if (src[i] === '$') {
    absRow = true;
    i += 1;
  }
  const digitsStart = i;
  while (i < src.length && /[0-9]/.test(src[i] ?? '')) i += 1;
  if (i === digitsStart || src[i] === '(') return null;
  return { absCol, label, absRow, rowStr: src.slice(digitsStart, i), end: i };
}

function colLabelToIndex(label: string): number {
  let n = 0;
  for (let i = 0; i < label.length; i += 1) n = n * 26 + (label.charCodeAt(i) - 64);
  return n - 1;
}

function colIndexToLabel(col: number): string {
  let n = col;
  let out = '';
  do {
    out = String.fromCharCode(65 + (n % 26)) + out;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return out;
}
