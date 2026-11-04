import { beforeEach, describe, expect, it, vi } from 'vitest';
import { autoSum, autoSumSuggestion } from '../../../src/commands/auto-sum.js';
import { applyValueFilter } from '../../../src/commands/filter.js';
import type { Addr } from '../../../src/engine/types.js';
import { addrKey, WorkbookHandle } from '../../../src/engine/workbook-handle.js';
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

  it('can insert the other AutoSum dropdown functions', () => {
    seedNumber(store, wb, 0, 0, 10);
    seedNumber(store, wb, 1, 0, 20);
    setActive(store, { sheet: 0, row: 2, col: 0 });

    const got = autoSum(store.getState(), wb, 'AVERAGE');
    expect(got).toEqual({ addr: { sheet: 0, row: 2, col: 0 }, formula: '=AVERAGE(A1:A2)' });
    expect(wb.cellFormula({ sheet: 0, row: 2, col: 0 })).toBe('=AVERAGE(A1:A2)');
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

  it('drops SUM into the empty trailing cell when a single-column selection includes blank rows below', () => {
    seedNumber(store, wb, 0, 0, 10);
    seedNumber(store, wb, 1, 0, 20);
    seedNumber(store, wb, 2, 0, 30);
    // A4, A5 are intentionally empty — included in the selection.
    setRangeOnly(store, 0, 0, 4, 0);

    const got = autoSum(store.getState(), wb);
    expect(got?.addr).toEqual({ sheet: 0, row: 3, col: 0 });
    expect(got?.formula).toBe('=SUM(A1:A3)');
    expect(wb.cellFormula({ sheet: 0, row: 3, col: 0 })).toBe('=SUM(A1:A3)');
    // The other trailing empty cell stays empty.
    expect(wb.cellFormula({ sheet: 0, row: 4, col: 0 })).toBe(null);
  });

  it('drops SUM into the empty trailing cell when a single-row selection includes blanks to the right', () => {
    seedNumber(store, wb, 0, 0, 1);
    seedNumber(store, wb, 0, 1, 2);
    seedNumber(store, wb, 0, 2, 3);
    setRangeOnly(store, 0, 0, 0, 4);

    const got = autoSum(store.getState(), wb);
    expect(got?.addr).toEqual({ sheet: 0, row: 0, col: 3 });
    expect(got?.formula).toBe('=SUM(A1:C1)');
  });

  it('fills a trailing-empty totals row across each column of a 2D selection', () => {
    // 3x3 numeric block plus an empty bottom row inside the selection.
    seedNumber(store, wb, 0, 0, 1);
    seedNumber(store, wb, 0, 1, 2);
    seedNumber(store, wb, 0, 2, 3);
    seedNumber(store, wb, 1, 0, 4);
    seedNumber(store, wb, 1, 1, 5);
    seedNumber(store, wb, 1, 2, 6);
    setRangeOnly(store, 0, 0, 2, 2);

    const got = autoSum(store.getState(), wb);
    expect(got?.addr).toEqual({ sheet: 0, row: 2, col: 0 });
    expect(got?.formula).toBe('=SUM(A1:A2)');
    expect(wb.cellFormula({ sheet: 0, row: 2, col: 0 })).toBe('=SUM(A1:A2)');
    expect(wb.cellFormula({ sheet: 0, row: 2, col: 1 })).toBe('=SUM(B1:B2)');
    expect(wb.cellFormula({ sheet: 0, row: 2, col: 2 })).toBe('=SUM(C1:C2)');
  });

  it('refuses huge multi-cell selections before scanning or writing formulas', () => {
    seedNumber(store, wb, 0, 0, 1);
    setRangeOnly(store, 0, 0, 100_000, 0);

    expect(autoSum(store.getState(), wb)).toBeNull();
    expect(wb.cellFormula({ sheet: 0, row: 1, col: 0 })).toBeNull();
  });

  it('skips locked protected targets in a multi-cell totals row', () => {
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});
    seedNumber(store, wb, 0, 0, 1);
    seedNumber(store, wb, 0, 1, 2);
    seedNumber(store, wb, 1, 0, 3);
    seedNumber(store, wb, 1, 1, 4);
    mutators.setCellFormat(store, { sheet: 0, row: 2, col: 1 }, { locked: false });
    mutators.setSheetProtected(store, 0, true);
    setRangeOnly(store, 0, 0, 2, 1);

    try {
      const got = autoSum(store.getState(), wb);

      expect(got?.addr).toEqual({ sheet: 0, row: 2, col: 1 });
      expect(wb.cellFormula({ sheet: 0, row: 2, col: 0 })).toBeNull();
      expect(wb.cellFormula({ sheet: 0, row: 2, col: 1 })).toBe('=SUM(B1:B2)');
      expect(warn).toHaveBeenCalled();
    } finally {
      warn.mockRestore();
    }
  });

  it('falls back to the row target when the column target is locked on a protected sheet', () => {
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});
    seedNumber(store, wb, 0, 0, 1);
    seedNumber(store, wb, 0, 1, 2);
    mutators.setCellFormat(store, { sheet: 0, row: 0, col: 2 }, { locked: false });
    mutators.setSheetProtected(store, 0, true);
    setActive(store, { sheet: 0, row: 0, col: 0 });

    try {
      const got = autoSum(store.getState(), wb);

      expect(got).toEqual({ addr: { sheet: 0, row: 0, col: 2 }, formula: '=SUM(A1:B1)' });
      expect(wb.cellFormula({ sheet: 0, row: 1, col: 0 })).toBeNull();
      expect(wb.cellFormula({ sheet: 0, row: 0, col: 2 })).toBe('=SUM(A1:B1)');
      expect(warn).toHaveBeenCalled();
    } finally {
      warn.mockRestore();
    }
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

  it('totals both directions of a 2D selection and sums the new total row in the corner', () => {
    // B10:D13 numbers, selection B10:E14 (one empty row below, one empty column right).
    for (let row = 9; row <= 12; row += 1)
      for (let col = 1; col <= 3; col += 1) seedNumber(store, wb, row, col, row + col);
    setRangeOnly(store, 9, 1, 13, 4);

    autoSum(store.getState(), wb);
    const f = (row: number, col: number) => wb.cellFormula({ sheet: 0, row, col });
    expect([f(13, 1), f(13, 2), f(13, 3)]).toEqual([
      '=SUM(B10:B13)',
      '=SUM(C10:C13)',
      '=SUM(D10:D13)',
    ]);
    expect([f(9, 4), f(10, 4), f(11, 4), f(12, 4)]).toEqual([
      '=SUM(B10:D10)',
      '=SUM(B11:D11)',
      '=SUM(B12:D12)',
      '=SUM(B13:D13)',
    ]);
    expect(f(13, 4)).toBe('=SUM(B14:D14)');
  });

  describe('autoSumSuggestion', () => {
    it('suggests the numbers above an empty cell, leaving the header out, without writing', () => {
      seedText(store, wb, 0, 1, 'Val');
      for (let row = 1; row <= 4; row += 1) seedNumber(store, wb, row, 1, row + 1);
      setActive(store, { sheet: 0, row: 5, col: 1 });
      expect(autoSumSuggestion(store.getState())).toEqual({
        addr: { sheet: 0, row: 5, col: 1 },
        formula: '=SUM(B2:B5)',
      });
      expect(wb.cellFormula({ sheet: 0, row: 5, col: 1 })).toBeNull();
    });

    it('stops at a blank gap above the cell', () => {
      seedText(store, wb, 0, 3, 'Gap');
      seedNumber(store, wb, 1, 3, 1);
      seedNumber(store, wb, 2, 3, 2);
      seedNumber(store, wb, 4, 3, 3);
      seedNumber(store, wb, 5, 3, 4);
      setActive(store, { sheet: 0, row: 6, col: 3 });
      expect(autoSumSuggestion(store.getState())?.formula).toBe('=SUM(D5:D6)');
    });

    it('suggests the row to the left when nothing is above', () => {
      for (let col = 5; col <= 8; col += 1) seedNumber(store, wb, 1, col, col);
      setActive(store, { sheet: 0, row: 1, col: 9 });
      expect(autoSumSuggestion(store.getState())?.formula).toBe('=SUM(F2:I2)');
    });

    it('suggests SUBTOTAL(9, ...) over the whole column below a filtered list', () => {
      seedText(store, wb, 9, 5, 'Cat');
      seedText(store, wb, 9, 6, 'N');
      for (let row = 10; row <= 13; row += 1) {
        seedText(store, wb, row, 5, row % 2 === 0 ? 'a' : 'b');
        seedNumber(store, wb, row, 6, row);
      }
      const range = { sheet: 0, r0: 9, c0: 5, r1: 13, c1: 6 };
      store.setState((s) => ({
        ...s,
        ui: {
          ...s.ui,
          filterRanges: [range],
          filterCriteria: [{ range, byCol: 5, hiddenValues: ['b'] }],
        },
      }));
      setActive(store, { sheet: 0, row: 14, col: 6 });
      expect(autoSumSuggestion(store.getState())?.formula).toBe('=SUBTOTAL(9,G11:G14)');
    });

    it('follows a value filter applied and cleared through the filter command', () => {
      seedText(store, wb, 9, 5, 'Cat');
      seedText(store, wb, 9, 6, 'N');
      for (let row = 10; row <= 13; row += 1) {
        seedText(store, wb, row, 5, row % 2 === 0 ? 'a' : 'b');
        seedNumber(store, wb, row, 6, row);
      }
      const range = { sheet: 0, r0: 9, c0: 5, r1: 13, c1: 6 };
      applyValueFilter(store.getState(), store, range, 5, ['b'], wb);
      expect(store.getState().layout.hiddenRows.has(11)).toBe(true);
      setActive(store, { sheet: 0, row: 14, col: 6 });
      expect(autoSumSuggestion(store.getState())?.formula).toBe('=SUBTOTAL(9,G11:G14)');

      applyValueFilter(store.getState(), store, range, 5, [], wb);
      expect(autoSumSuggestion(store.getState())?.formula).toBe('=SUM(G11:G14)');
    });

    it('keeps SUM below a filter range that has no criteria applied', () => {
      seedText(store, wb, 9, 6, 'N');
      for (let row = 10; row <= 13; row += 1) seedNumber(store, wb, row, 6, row);
      const range = { sheet: 0, r0: 9, c0: 6, r1: 13, c1: 6 };
      store.setState((s) => ({ ...s, ui: { ...s.ui, filterRanges: [range], filterCriteria: [] } }));
      setActive(store, { sheet: 0, row: 14, col: 6 });
      expect(autoSumSuggestion(store.getState())?.formula).toBe('=SUM(G11:G14)');
    });

    it('has no suggestion for a range selection or an occupied cell', () => {
      seedNumber(store, wb, 0, 0, 1);
      setActive(store, { sheet: 0, row: 0, col: 0 });
      expect(autoSumSuggestion(store.getState())).toBeNull();
      setRangeOnly(store, 0, 0, 1, 0);
      expect(autoSumSuggestion(store.getState())).toBeNull();
    });
  });
});
