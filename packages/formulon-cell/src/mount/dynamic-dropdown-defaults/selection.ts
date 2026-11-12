import type { Range } from '../../engine/types.js';
import type { SpreadsheetInstance } from '../../mount.js';
import { normalizeRange } from '../../store/selection-geometry.js';

export const normalizedSelectionRange = (instance: SpreadsheetInstance): Range => {
  const r = instance.store.getState().selection.range;
  return normalizeRange(r);
};

export const hasActiveCopy = (instance: SpreadsheetInstance): boolean => {
  const { copyRange, copyRanges } = instance.store.getState().ui;
  return Boolean(copyRange || copyRanges?.length);
};
