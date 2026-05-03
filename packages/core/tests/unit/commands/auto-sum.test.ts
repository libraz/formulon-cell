import { beforeEach, describe, expect, it } from 'vitest';
import { autoSum } from '../../../src/commands/auto-sum.js';
import type { Addr } from '../../../src/engine/types.js';
import { WorkbookHandle, addrKey } from '../../../src/engine/workbook-handle.js';
import { type SpreadsheetStore, createSpreadsheetStore } from '../../../src/store/store.js';

const newWb = (): Promise<WorkbookHandle> => WorkbookHandle.createDefault({ preferStub: true });

const seedNumber = (
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  row: number,
  col: number,
  value: number,
): void => {
  wb.setNumber({ sheet: 0, row, col }, value);
  store.setState((s) => {
    const map = new Map(s.data.cells);
    map.set(addrKey({ sheet: 0, row, col }), {
      value: { kind: 'number', value },
      formula: null,
    });
    return { ...s, data: { ...s.data, cells: map } };
  });
};

// Text neither counts as a number (autoSum won't extend a block into it) nor
// as empty (autoSum won't drop a SUM there) — it's the cleanest "occupied non-numeric" marker.
const seedText = (
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  row: number,
  col: number,
  value: string,
): void => {
  wb.setText({ sheet: 0, row, col }, value);
  store.setState((s) => {
    const map = new Map(s.data.cells);
    map.set(addrKey({ sheet: 0, row, col }), {
      value: { kind: 'text', value },
      formula: null,
    });
    return { ...s, data: { ...s.data, cells: map } };
  });
};

const setActive = (store: SpreadsheetStore, addr: Addr): void => {
  store.setState((s) => ({
    ...s,
    selection: {
      active: addr,
      anchor: addr,
      range: { sheet: addr.sheet, r0: addr.row, c0: addr.col, r1: addr.row, c1: addr.col },
    },
  }));
};

const setRangeOnly = (
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
      active: { sheet: 0, row: r0, col: c0 },
      anchor: { sheet: 0, row: r0, col: c0 },
      range: { sheet: 0, r0, c0, r1, c1 },
    },
  }));
};

describe('autoSum', () => {
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;

  beforeEach(async () => {
    store = createSpreadsheetStore();
    wb = await newWb();
  });

  it('writes SUM into an empty active cell using the column block above', () => {
    seedNumber(store, wb, 0, 0, 10);
    seedNumber(store, wb, 1, 0, 20);
    seedNumber(store, wb, 2, 0, 30);
    setActive(store, { sheet: 0, row: 3, col: 0 });

    const got = autoSum(store.getState(), wb);
    expect(got).toEqual({ addr: { sheet: 0, row: 3, col: 0 }, formula: '=SUM(A1:A3)' });
    expect(wb.cellFormula({ sheet: 0, row: 3, col: 0 })).toBe('=SUM(A1:A3)');
  });

  it('falls back to the row to the left when the column above is empty', () => {
    seedNumber(store, wb, 0, 0, 1);
    seedNumber(store, wb, 0, 1, 2);
    seedNumber(store, wb, 0, 2, 3);
    setActive(store, { sheet: 0, row: 0, col: 3 });

    const got = autoSum(store.getState(), wb);
    expect(got?.formula).toBe('=SUM(A1:C1)');
    expect(got?.addr).toEqual({ sheet: 0, row: 0, col: 3 });
  });

  it('returns null when the active cell is empty and has no neighbors', () => {
    setActive(store, { sheet: 0, row: 5, col: 5 });
    expect(autoSum(store.getState(), wb)).toBeNull();
  });

  it('places SUM in the cell directly below a numeric column block', () => {
    seedNumber(store, wb, 0, 0, 10);
    seedNumber(store, wb, 1, 0, 20);
    seedNumber(store, wb, 2, 0, 30);
    // Active sits inside the block.
    setActive(store, { sheet: 0, row: 1, col: 0 });

    const got = autoSum(store.getState(), wb);
    expect(got).toEqual({ addr: { sheet: 0, row: 3, col: 0 }, formula: '=SUM(A1:A3)' });
  });

  it('falls back to placing SUM to the right of a row block when below is occupied', () => {
    seedNumber(store, wb, 0, 0, 1);
    seedNumber(store, wb, 0, 1, 2);
    seedNumber(store, wb, 0, 2, 3);
    // The cell directly below the column block (row 1, col 0) is occupied by
    // a non-numeric value — autoSum can't drop SUM there, so it falls back
    // to the row-direction block.
    seedText(store, wb, 1, 0, 'lock');
    setActive(store, { sheet: 0, row: 0, col: 0 });

    const got = autoSum(store.getState(), wb);
    expect(got?.addr).toEqual({ sheet: 0, row: 0, col: 3 });
    expect(got?.formula).toBe('=SUM(A1:C1)');
  });

  it('handles a multi-cell range — places SUM directly below the range', () => {
    seedNumber(store, wb, 0, 0, 1);
    seedNumber(store, wb, 0, 1, 2);
    seedNumber(store, wb, 1, 0, 3);
    seedNumber(store, wb, 1, 1, 4);
    setRangeOnly(store, 0, 0, 1, 1);

    const got = autoSum(store.getState(), wb);
    expect(got?.addr).toEqual({ sheet: 0, row: 2, col: 0 });
    expect(got?.formula).toBe('=SUM(A1:B2)');
  });

  it('returns null when both column-target and row-target are occupied by non-numeric cells', () => {
    seedNumber(store, wb, 0, 0, 1);
    // Both candidate targets blocked by text — neither path can place SUM.
    seedText(store, wb, 1, 0, 'x');
    seedText(store, wb, 0, 1, 'y');
    setRangeOnly(store, 0, 0, 0, 0);
    expect(autoSum(store.getState(), wb)).toBeNull();
  });

  it('returns null for a multi-cell numeric range when both candidate targets are non-empty', () => {
    seedNumber(store, wb, 0, 0, 1);
    seedNumber(store, wb, 0, 1, 2);
    seedNumber(store, wb, 1, 0, 3);
    seedNumber(store, wb, 1, 1, 4);
    // Below + right of the 2x2 block: both occupied.
    seedText(store, wb, 2, 0, 'lock');
    seedText(store, wb, 0, 2, 'lock');
    setRangeOnly(store, 0, 0, 1, 1);
    expect(autoSum(store.getState(), wb)).toBeNull();
  });
});
