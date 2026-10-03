import { MAX_COL, MAX_ROW } from '../../engine/address.js';
import type { Range, SpreadsheetInstance } from '../../index.js';

export const normalizedSelectionRange = (instance: SpreadsheetInstance): Range => {
  const r = instance.store.getState().selection.range;
  return {
    sheet: r.sheet,
    r0: Math.min(r.r0, r.r1),
    c0: Math.min(r.c0, r.c1),
    r1: Math.max(r.r0, r.r1),
    c1: Math.max(r.c0, r.c1),
  };
};

export const isWholeRowSelection = (range: Range): boolean => range.c0 === 0 && range.c1 >= MAX_COL;

export const isWholeColumnSelection = (range: Range): boolean =>
  range.r0 === 0 && range.r1 >= MAX_ROW;

export const hasActiveCopy = (instance: SpreadsheetInstance): boolean => {
  const { copyRange, copyRanges } = instance.store.getState().ui;
  return Boolean(copyRange || copyRanges?.length);
};

export const addrFromKey = (key: string): { sheet: number; row: number; col: number } | null => {
  const [sheetRaw, rowRaw, colRaw] = key.split(':');
  const sheet = Number(sheetRaw);
  const row = Number(rowRaw);
  const col = Number(colRaw);
  if (!Number.isInteger(sheet) || !Number.isInteger(row) || !Number.isInteger(col)) return null;
  return { sheet, row, col };
};
