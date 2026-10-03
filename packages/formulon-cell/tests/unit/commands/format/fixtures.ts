import { addrKey } from '../../../../src/engine/workbook-handle.js';
import type { CellFormat, SpreadsheetStore } from '../../../../src/store/store.js';

export const setRange = (
  store: SpreadsheetStore,
  r0: number,
  c0: number,
  r1: number,
  c1: number,
): void => {
  store.setState((s) => ({
    ...s,
    selection: {
      ...s.selection,
      range: { sheet: 0, r0, c0, r1, c1 },
    },
  }));
};

export const setSelection = (
  store: SpreadsheetStore,
  range: { sheet: number; r0: number; c0: number; r1: number; c1: number },
  extraRanges: (typeof range)[] = [],
): void => {
  store.setState((s) => ({
    ...s,
    selection: {
      ...s.selection,
      active: { sheet: range.sheet, row: range.r0, col: range.c0 },
      anchor: { sheet: range.sheet, row: range.r0, col: range.c0 },
      range,
      extraRanges,
    },
  }));
};

export const fmtAt = (store: SpreadsheetStore, row: number, col: number): CellFormat | undefined =>
  store.getState().format.formats.get(addrKey({ sheet: 0, row, col }));

export const effectiveFmtAt = (
  store: SpreadsheetStore,
  row: number,
  col: number,
): CellFormat | undefined => {
  const stored = fmtAt(store, row, col);
  const pending = store.getState().ui.pendingFormat;
  if (pending?.addr.sheet !== 0 || pending.addr.row !== row || pending.addr.col !== col) {
    return stored;
  }
  return { ...(stored ?? {}), ...pending.format };
};
