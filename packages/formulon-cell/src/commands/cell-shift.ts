import { addrKey, MAX_COL, MAX_ROW } from '../engine/address.js';
import type { Addr, CellValue, Range } from '../engine/types.js';
import { writeCell } from '../engine/value.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { addMergeToMaps } from '../store/merge-maps.js';
import { rangeContainsAddr, rangesIntersect } from '../store/selection-geometry.js';
import { type CellFormat, mutators, type SpreadsheetStore, type State } from '../store/store.js';
import { listComments, recordCommentChange } from './comment.js';
import { adjustFormulaForCellBandShift } from './formula-refs.js';
import type { History } from './history.js';
import { isSheetProtected } from './protection.js';
import { recordFormatChange, recordMergesChangeWithEngine } from './slice-history.js';

export type InsertCellsDirection = 'down' | 'right';
export type DeleteCellsDirection = 'up' | 'left';

export interface CellRecord {
  addr: Addr;
  value: CellValue;
  formula: string | null;
}

interface EngineCommentRecord {
  addr: Addr;
  author: string;
  text: string;
}

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
  if (delta > 0 && shiftWouldOverflow(store, wb, affected, direction, delta)) return false;
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
  if (!canShiftMerges(store.getState(), affected, axis)) {
    // eslint-disable-next-line no-console
    console.warn('formulon-cell: cell shift blocked — merge would be split');
    return false;
  }

  if (history) history.begin();
  try {
    const beforeCells =
      history && !history.isReplaying() ? collectSheetCells(wb, affected.sheet) : null;
    const beforeExternalFormulas =
      history && !history.isReplaying() ? collectExternalFormulaCells(wb, affected.sheet) : null;
    shiftCells(wb, affected, axis, delta);
    if (history && beforeCells && beforeExternalFormulas) {
      const afterCells = collectSheetCells(wb, affected.sheet);
      const afterExternalFormulas = collectExternalFormulaCells(wb, affected.sheet);
      history.push({
        undo: () => {
          restoreSheetCells(wb, affected.sheet, beforeCells);
          restoreFormulaCells(wb, beforeExternalFormulas);
        },
        redo: () => {
          restoreSheetCells(wb, affected.sheet, afterCells);
          restoreFormulaCells(wb, afterExternalFormulas);
        },
      });
    }
    const initialStoreCommentAddresses = listComments(store.getState(), affected.sheet)
      .filter((comment) => rangeContainsAddr(affected, comment.addr))
      .map((comment) => comment.addr);
    const initialStoreCommentKeys = new Set(initialStoreCommentAddresses.map(addrKey));
    const engineComments = collectEngineComments(wb, affected, initialStoreCommentAddresses);
    hydrateEngineOnlyComments(store, engineComments, initialStoreCommentKeys);
    const storeCommentAddresses = listComments(store.getState(), affected.sheet)
      .filter((comment) => rangeContainsAddr(affected, comment.addr))
      .map((comment) => comment.addr);
    const commentAddresses = collectShiftCommentAddresses(
      storeCommentAddresses,
      engineComments,
      affected,
      axis,
      delta,
    );
    recordCommentChange(history, store, wb, commentAddresses, () => {
      shiftFormats(store, history, affected, axis, delta);
      shiftEngineComments(wb, affected, axis, delta, engineComments);
    });
    shiftMerges(store, wb, history, affected, axis, delta);
    wb.recalcAuto();
  } finally {
    if (history) history.end();
  }
  return true;
}

function collectSheetCells(wb: WorkbookHandle, sheet: number): CellRecord[] {
  return Array.from(wb.physicalCells(sheet)).map<CellRecord>((c) => ({
    addr: c.addr,
    value: c.value,
    formula: c.formula,
  }));
}

function collectExternalFormulaCells(wb: WorkbookHandle, editedSheet: number): CellRecord[] {
  const formulas: CellRecord[] = [];
  for (let sheet = 0; sheet < wb.sheetCount; sheet += 1) {
    if (sheet === editedSheet) continue;
    for (const cell of wb.physicalCells(sheet)) {
      if (!cell.formula) continue;
      formulas.push({ addr: cell.addr, value: cell.value, formula: cell.formula });
    }
  }
  return formulas;
}

function restoreSheetCells(wb: WorkbookHandle, sheet: number, cells: readonly CellRecord[]): void {
  const existing = Array.from(wb.physicalCells(sheet));
  wb.withBatchedRecalc(() => {
    for (const cell of existing) wb.setBlank(cell.addr);
    for (const cell of cells) writeCell(wb, cell.addr, cell.value, cell.formula);
  });
  wb.recalcAuto();
}

function restoreFormulaCells(wb: WorkbookHandle, cells: readonly CellRecord[]): void {
  wb.withBatchedRecalc(() => {
    for (const cell of cells) {
      if (cell.formula) wb.setFormula(cell.addr, cell.formula);
    }
  });
  wb.recalcAuto();
}

function collectShiftCommentAddresses(
  storeCommentAddresses: readonly Addr[],
  engineComments: readonly EngineCommentRecord[],
  affected: Range,
  axis: InsertCellsDirection,
  delta: number,
): Addr[] {
  const tracked = new Map<string, Addr>();
  const add = (addr: Addr): void => {
    tracked.set(addrKey(addr), addr);
  };
  for (const addr of storeCommentAddresses) {
    add(addr);
    const shifted = shiftedCommentAddr(addr, axis, delta);
    if (inShiftTarget(shifted, affected, axis)) add(shifted);
  }
  for (const comment of engineComments) {
    add(comment.addr);
    const shifted = shiftedCommentAddr(comment.addr, axis, delta);
    if (inShiftTarget(shifted, affected, axis)) add(shifted);
  }
  return [...tracked.values()];
}

function shiftedCommentAddr(addr: Addr, axis: InsertCellsDirection, delta: number): Addr {
  return axis === 'down' ? { ...addr, row: addr.row + delta } : { ...addr, col: addr.col + delta };
}

function shiftEngineComments(
  wb: WorkbookHandle,
  affected: Range,
  axis: InsertCellsDirection,
  delta: number,
  entries: readonly EngineCommentRecord[],
): void {
  if (!wb.capabilities.comments) return;
  if (entries.length === 0) return;

  wb.withBatchedRecalc(() => {
    for (const entry of entries) {
      wb.setCommentEntry(entry.addr.sheet, entry.addr.row, entry.addr.col, '', '');
    }
    for (const entry of entries) {
      const shifted = shiftedCommentAddr(entry.addr, axis, delta);
      if (!inShiftTarget(shifted, affected, axis)) continue;
      wb.setCommentEntry(shifted.sheet, shifted.row, shifted.col, entry.author, entry.text);
    }
  });
}

function collectEngineComments(
  wb: WorkbookHandle,
  affected: Range,
  storeCommentAddresses: readonly Addr[],
): EngineCommentRecord[] {
  const entries = wb.capabilities.commentsEnumerable
    ? wb.getComments(affected.sheet).map((comment) => ({
        addr: { sheet: affected.sheet, row: comment.row, col: comment.col },
        author: comment.author,
        text: comment.text,
      }))
    : storeCommentAddresses.flatMap((addr) => {
        const comment = wb.getComment(addr.sheet, addr.row, addr.col);
        return comment ? [{ addr, author: comment.author, text: comment.text }] : [];
      });
  const unique = new Map<string, EngineCommentRecord>();
  for (const entry of entries) {
    if (rangeContainsAddr(affected, entry.addr)) unique.set(addrKey(entry.addr), entry);
  }
  return [...unique.values()];
}

function hydrateEngineOnlyComments(
  store: SpreadsheetStore,
  entries: readonly EngineCommentRecord[],
  storeCommentKeys: ReadonlySet<string>,
): void {
  for (const entry of entries) {
    if (storeCommentKeys.has(addrKey(entry.addr))) continue;
    mutators.setCellFormat(store, entry.addr, {
      comment: entry.text,
      commentAuthor: entry.author || undefined,
    });
  }
}

/** Move the cells inside `affected` by `delta` along `axis` and rewrite every
 *  formula on every sheet that references the moved band. */
export function shiftCells(
  wb: WorkbookHandle,
  affected: Range,
  axis: InsertCellsDirection,
  delta: number,
): void {
  wb.withBatchedRecalc(() => writeShiftedCells(wb, affected, axis, delta));
}

function writeShiftedCells(
  wb: WorkbookHandle,
  affected: Range,
  axis: InsertCellsDirection,
  delta: number,
): void {
  const allSheets = Array.from({ length: wb.sheetCount }, (_, sheet) =>
    Array.from(wb.physicalCells(sheet)).map<CellRecord>((c) => ({
      addr: c.addr,
      value: c.value,
      formula: c.formula,
    })),
  );
  const all = allSheets[affected.sheet] ?? [];
  const sheetNames = Array.from({ length: wb.sheetCount }, (_, sheet) => wb.sheetName(sheet));
  const moving = all.filter((cell) => rangeContainsAddr(affected, cell.addr));
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
    const formula = cell.formula
      ? adjustFormulaForCellBandShift(cell.formula, affected, axis, delta, {
          editedSheet: affected.sheet,
          formulaSheet: affected.sheet,
          sheetNames,
        })
      : null;
    writeCell(wb, next, cell.value, formula);
  }

  for (let sheet = 0; sheet < allSheets.length; sheet += 1) {
    const formulaSheet = allSheets[sheet] ?? [];
    for (const cell of formulaSheet) {
      if (!cell.formula || (sheet === affected.sheet && rangeContainsAddr(affected, cell.addr)))
        continue;
      const nextFormula = adjustFormulaForCellBandShift(cell.formula, affected, axis, delta, {
        editedSheet: affected.sheet,
        formulaSheet: sheet,
        sheetNames,
      });
      if (nextFormula !== cell.formula) wb.setFormula(cell.addr, nextFormula);
    }
  }
}

export function shiftFormats(
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

export function shiftMerges(
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

/** Whether every merge touching the shift band moves whole; false when a
 *  merge straddles the band edge and would be split. */
export function canShiftMerges(state: State, affected: Range, axis: InsertCellsDirection): boolean {
  for (const merge of state.merges.byAnchor.values()) {
    if (merge.sheet !== affected.sheet) continue;
    if (!rangesIntersect(merge, affected)) continue;
    if (!mergeIntersectsShiftBand(merge, affected, axis)) continue;
    const fullyInsideBand =
      axis === 'down'
        ? merge.c0 >= affected.c0 && merge.c1 <= affected.c1 && merge.r0 >= affected.r0
        : merge.r0 >= affected.r0 && merge.r1 <= affected.r1 && merge.c0 >= affected.c0;
    if (!fullyInsideBand) return false;
  }
  return true;
}

function shiftWouldOverflow(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  affected: Range,
  axis: InsertCellsDirection,
  delta: number,
): boolean {
  const overflows = (addr: Addr): boolean => {
    if (!rangeContainsAddr(affected, addr)) return false;
    return axis === 'down' ? addr.row + delta > MAX_ROW : addr.col + delta > MAX_COL;
  };
  for (const cell of wb.physicalCells(affected.sheet)) {
    if (overflows(cell.addr)) {
      // eslint-disable-next-line no-console
      console.warn('formulon-cell: cell shift blocked — content would leave the worksheet');
      return true;
    }
  }
  for (const key of store.getState().format.formats.keys()) {
    const parts = key.split(':');
    if (parts.length !== 3) continue;
    const addr: Addr = { sheet: Number(parts[0]), row: Number(parts[1]), col: Number(parts[2]) };
    if (overflows(addr)) {
      // eslint-disable-next-line no-console
      console.warn('formulon-cell: cell shift blocked — format would leave the worksheet');
      return true;
    }
  }
  if (wb.capabilities.commentsEnumerable) {
    for (const comment of wb.getComments(affected.sheet)) {
      if (overflows({ sheet: affected.sheet, row: comment.row, col: comment.col })) {
        // eslint-disable-next-line no-console
        console.warn('formulon-cell: cell shift blocked — comment would leave the worksheet');
        return true;
      }
    }
  }
  for (const merge of store.getState().merges.byAnchor.values()) {
    if (merge.sheet !== affected.sheet || !mergeIntersectsShiftBand(merge, affected, axis)) {
      continue;
    }
    const shifted = shiftRange(merge, axis, delta);
    if (shifted.r1 > MAX_ROW || shifted.c1 > MAX_COL) {
      // eslint-disable-next-line no-console
      console.warn('formulon-cell: cell shift blocked — merge would leave the worksheet');
      return true;
    }
  }
  return false;
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
    if (!rangeContainsAddr(affected, addr)) {
      next.set(key, fmt);
      continue;
    }
    const shifted =
      axis === 'down' ? { ...addr, row: addr.row + delta } : { ...addr, col: addr.col + delta };
    if (inShiftTarget(shifted, affected, axis)) next.set(addrKey(shifted), fmt);
  }
  return next;
}

function inShiftTarget(addr: Addr, affected: Range, axis: InsertCellsDirection): boolean {
  if (addr.sheet !== affected.sheet || addr.row < 0 || addr.col < 0) return false;
  if (addr.row > MAX_ROW || addr.col > MAX_COL) return false;
  return axis === 'down'
    ? addr.row >= affected.r0 && addr.col >= affected.c0 && addr.col <= affected.c1
    : addr.col >= affected.c0 && addr.row >= affected.r0 && addr.row <= affected.r1;
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

// Reference shifting for cell-band inserts/deletes is delegated to the shared
// `adjustFormulaForCellBandShift` (see commands/formula-refs.ts), which handles
// sheet qualifiers, function names, ranges, and grid bounds uniformly.
