import { addrKey, MAX_COL, MAX_ROW } from '../engine/address.js';
import type { Addr, Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { State } from '../store/types.js';
import type { OperationEffect } from './interaction-policy.js';

const MAX_MATERIALIZED_CELLS = 100_000;

export const effectCells = (effect: OperationEffect): readonly Addr[] | null => {
  if (effect.kind === 'cells')
    return effect.cells.length <= MAX_MATERIALIZED_CELLS ? effect.cells : null;
  if (effect.kind === 'workbook') return [];
  const { range } = effect;
  if (
    !Number.isInteger(range.sheet) ||
    !Number.isInteger(range.r0) ||
    !Number.isInteger(range.c0) ||
    !Number.isInteger(range.r1) ||
    !Number.isInteger(range.c1) ||
    range.sheet < 0 ||
    range.r0 < 0 ||
    range.c0 < 0 ||
    range.r1 < range.r0 ||
    range.c1 < range.c0
  )
    return null;
  const area = (range.r1 - range.r0 + 1) * (range.c1 - range.c0 + 1);
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

export const isWorkbookAddress = (wb: WorkbookHandle, addr: Addr): boolean => {
  try {
    return (
      Number.isInteger(addr.sheet) &&
      Number.isInteger(addr.row) &&
      Number.isInteger(addr.col) &&
      addr.sheet >= 0 &&
      addr.row >= 0 &&
      addr.col >= 0 &&
      addr.row <= MAX_ROW &&
      addr.col <= MAX_COL &&
      addr.sheet < wb.sheetCount
    );
  } catch {
    return false;
  }
};
