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

export const hasActiveCopy = (instance: SpreadsheetInstance): boolean => {
  const { copyRange, copyRanges } = instance.store.getState().ui;
  return Boolean(copyRange || copyRanges?.length);
};
