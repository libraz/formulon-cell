import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { autoSum } from '../../../src/commands/auto-sum.js';
import { applyMerge } from '../../../src/commands/merge.js';
import type { Addr, Range } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

const newWb = (): Promise<WorkbookHandle> => WorkbookHandle.createDefault({ preferStub: true });

const seedNumber = (
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  row: number,
  col: number,
  value: number,
): void => {
  const addr = { sheet: 0, row, col };
  wb.setNumber(addr, value);
  mutators.setCell(store, addr, { kind: 'number', value });
};

const setRange = (store: SpreadsheetStore, range: Range): void => {
  const active: Addr = { sheet: range.sheet, row: range.r0, col: range.c0 };
  store.setState((s) => ({
    ...s,
    selection: {
      ...s.selection,
      active,
      anchor: active,
      range,
    },
  }));
};

describe('autoSum with merged selections', () => {
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;

  beforeEach(async () => {
    store = createSpreadsheetStore();
    wb = await newWb();
  });

  afterEach(() => wb.dispose());

  it('places a range SUM below a horizontal merge instead of in its hidden body', () => {
    seedNumber(store, wb, 0, 0, 100);
    seedNumber(store, wb, 1, 0, 20);
    seedNumber(store, wb, 1, 1, 30);
    applyMerge(store, wb, null, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 });
    setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 });

    const result = autoSum(store.getState(), wb);

    expect(result).toEqual({ addr: { sheet: 0, row: 2, col: 0 }, formula: '=SUM(A1:B2)' });
    expect(wb.cellFormula({ sheet: 0, row: 2, col: 0 })).toBe('=SUM(A1:B2)');
    expect(wb.cellFormula({ sheet: 0, row: 0, col: 1 })).toBeNull();
  });

  it('places a range SUM below a vertical merge instead of in its hidden body', () => {
    seedNumber(store, wb, 0, 0, 100);
    seedNumber(store, wb, 0, 1, 20);
    seedNumber(store, wb, 1, 1, 30);
    applyMerge(store, wb, null, { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 });
    setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 });

    const result = autoSum(store.getState(), wb);

    expect(result).toEqual({ addr: { sheet: 0, row: 2, col: 0 }, formula: '=SUM(A1:B2)' });
    expect(wb.cellFormula({ sheet: 0, row: 2, col: 0 })).toBe('=SUM(A1:B2)');
    expect(wb.cellFormula({ sheet: 0, row: 1, col: 0 })).toBeNull();
  });

  it('uses the requested aggregate function with a merged selection', () => {
    seedNumber(store, wb, 0, 0, 100);
    seedNumber(store, wb, 1, 0, 20);
    seedNumber(store, wb, 1, 1, 30);
    applyMerge(store, wb, null, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 });
    setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 });

    const result = autoSum(store.getState(), wb, 'AVERAGE');

    expect(result).toEqual({ addr: { sheet: 0, row: 2, col: 0 }, formula: '=AVERAGE(A1:B2)' });
    expect(wb.cellFormula({ sheet: 0, row: 0, col: 1 })).toBeNull();
  });

  it('keeps ordinary range AutoSum placement unchanged', () => {
    seedNumber(store, wb, 0, 0, 10);
    seedNumber(store, wb, 0, 1, 20);
    seedNumber(store, wb, 1, 0, 30);
    seedNumber(store, wb, 1, 1, 40);
    setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 });

    const result = autoSum(store.getState(), wb);

    expect(result).toEqual({ addr: { sheet: 0, row: 2, col: 0 }, formula: '=SUM(A1:B2)' });
    expect(wb.cellFormula({ sheet: 0, row: 2, col: 0 })).toBe('=SUM(A1:B2)');
  });
});
