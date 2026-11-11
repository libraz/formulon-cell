import { addrKey, MAX_COL, MAX_ROW } from '../engine/address.js';
import type { Addr, Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { State } from '../store/types.js';
import type { OperationEffect, PermissionDecision } from './interaction-policy.js';

const MAX_MATERIALIZED_CELLS = 100_000;

const isWellFormedRange = (range: Range): boolean =>
  Number.isInteger(range.sheet) &&
  Number.isInteger(range.r0) &&
  Number.isInteger(range.c0) &&
  Number.isInteger(range.r1) &&
  Number.isInteger(range.c1) &&
  range.sheet >= 0 &&
  range.r0 >= 0 &&
  range.c0 >= 0 &&
  range.r1 >= range.r0 &&
  range.c1 >= range.c0;

const rangeCellCount = (range: Range): number =>
  (range.r1 - range.r0 + 1) * (range.c1 - range.c0 + 1);

/** The band a row insert/delete shifts: from `firstRow` to the sheet end, full width. */
export const rowShiftBand = (sheet: number, firstRow: number): Range => ({
  sheet,
  r0: firstRow,
  c0: 0,
  r1: MAX_ROW,
  c1: MAX_COL,
});

/** The band a column insert/delete shifts: from `firstCol` to the sheet end, full height. */
export const columnShiftBand = (sheet: number, firstCol: number): Range => ({
  sheet,
  r0: 0,
  c0: firstCol,
  r1: MAX_ROW,
  c1: MAX_COL,
});

/** The band a cell shift of `range` moves: its columns down to the sheet end, or its rows right
 *  to the sheet end. */
export const cellShiftBand = (range: Range, vertical: boolean): Range =>
  vertical
    ? { sheet: range.sheet, r0: range.r0, c0: range.c0, r1: MAX_ROW, c1: range.c1 }
    : { sheet: range.sheet, r0: range.r0, c0: range.c0, r1: range.r1, c1: MAX_COL };

/** The range of a well-formed range effect too large to enumerate cell by
 *  cell (a row/column shift band); such a range is checked as a whole. */
export const bandRange = (effect: OperationEffect): Range | null => {
  if (effect.kind !== 'range' || !isWellFormedRange(effect.range)) return null;
  const area = rangeCellCount(effect.range);
  return Number.isSafeInteger(area) && area <= MAX_MATERIALIZED_CELLS ? null : effect.range;
};

/** The refusal for an effect `effectCells` cannot enumerate: a malformed range
 *  is invalid input, anything else names more cells than the bound. */
export const unenumerableEffect = (effect: OperationEffect): PermissionDecision =>
  effect.kind === 'range' && !isWellFormedRange(effect.range)
    ? { allowed: false, code: 'invalid', reason: 'affected range is malformed' }
    : {
        allowed: false,
        code: 'effectTooLarge',
        reason: 'affected range exceeds the authorization bound',
      };

export const effectCells = (effect: OperationEffect): readonly Addr[] | null => {
  if (effect.kind === 'cells')
    return effect.cells.length <= MAX_MATERIALIZED_CELLS ? effect.cells : null;
  // A sheet effect names no cells; the authorizer decides the sheet as a whole.
  if (effect.kind === 'workbook' || effect.kind === 'sheet') return [];
  const { range } = effect;
  if (!isWellFormedRange(range)) return null;
  const area = rangeCellCount(range);
  if (!Number.isSafeInteger(area) || area < 0 || area > MAX_MATERIALIZED_CELLS) return null;
  const cells: Addr[] = [];
  for (let row = range.r0; row <= range.r1; row += 1) {
    for (let col = range.c0; col <= range.c1; col += 1)
      cells.push({ sheet: range.sheet, row, col });
  }
  return cells;
};

/** Return the merge covering a cell without importing the merge command module
 * (the command module itself records history and would create a needless
 * dependency cycle at this policy boundary). */
const mergeRangeAt = (state: State, addr: Addr): Range | null => {
  const key = addrKey(addr);
  const anchorKey = state.merges.byCell.get(key) ?? key;
  return state.merges.byAnchor.get(anchorKey) ?? null;
};

export const mergeCellsFor = (state: State, addr: Addr): readonly Addr[] | null => {
  const merge = mergeRangeAt(state, addr);
  if (!merge) return [addr];
  const height = merge.r1 - merge.r0 + 1;
  const width = merge.c1 - merge.c0 + 1;
  const area = height * width;
  if (!Number.isSafeInteger(area) || height <= 0 || width <= 0 || area > MAX_MATERIALIZED_CELLS)
    return null;
  const cells: Addr[] = [];
  for (let row = merge.r0; row <= merge.r1; row += 1) {
    for (let col = merge.c0; col <= merge.c1; col += 1)
      cells.push({ sheet: merge.sheet, row, col });
  }
  return cells;
};

export const mergeAnchorFor = (state: State, addr: Addr): Addr => {
  const merge = mergeRangeAt(state, addr);
  return merge ? { sheet: merge.sheet, row: merge.r0, col: merge.c0 } : addr;
};

export const cellEffects = (
  cells: readonly Addr[],
  formulaCells: readonly Addr[] = [],
): readonly OperationEffect[] => {
  const effects: OperationEffect[] = [{ kind: 'cells', cells }];
  if (formulaCells.length > 0)
    effects.push({ kind: 'cells', cells: formulaCells, includesFormula: true });
  return effects;
};

const isWorkbookAddressWithSheetCount = (
  readSheetCount: () => number | null,
  addr: Addr,
): boolean => {
  try {
    if (
      !Number.isInteger(addr.sheet) ||
      !Number.isInteger(addr.row) ||
      !Number.isInteger(addr.col) ||
      addr.sheet < 0 ||
      addr.row < 0 ||
      addr.col < 0 ||
      addr.row > MAX_ROW ||
      addr.col > MAX_COL
    ) {
      return false;
    }
    const sheetCount = readSheetCount();
    return sheetCount !== null && addr.sheet < sheetCount;
  } catch {
    return false;
  }
};

export const isWorkbookAddress = (wb: WorkbookHandle, addr: Addr): boolean =>
  isWorkbookAddressWithSheetCount(() => wb.sheetCount, addr);

/** Create a decision-local address check that reads the workbook's sheet
 * count only after coordinates pass validation, then reuses that result. */
export const createWorkbookAddressCheck = (wb: WorkbookHandle): ((addr: Addr) => boolean) => {
  let sheetCountRead = false;
  let sheetCount: number | null = null;
  const readSheetCount = (): number | null => {
    if (!sheetCountRead) {
      sheetCountRead = true;
      try {
        sheetCount = wb.sheetCount;
      } catch {
        sheetCount = null;
      }
    }
    return sheetCount;
  };
  return (addr) => isWorkbookAddressWithSheetCount(readSheetCount, addr);
};
