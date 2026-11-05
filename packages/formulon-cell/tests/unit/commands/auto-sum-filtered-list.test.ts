import { beforeEach, describe, expect, it } from 'vitest';
import { autoSum, autoSumSuggestion } from '../../../src/commands/auto-sum.js';
import { addrKey, WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, type SpreadsheetStore } from '../../../src/store/store.js';

describe('autoSum inside a filtered list', () => {
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;

  beforeEach(async () => {
    store = createSpreadsheetStore();
    wb = await WorkbookHandle.createDefault({ preferStub: true });
    const cells = new Map(store.getState().data.cells);
    for (let row = 10; row <= 13; row += 1) {
      wb.setNumber({ sheet: 0, row, col: 6 }, row);
      cells.set(addrKey({ sheet: 0, row, col: 6 }), {
        value: { kind: 'number', value: row },
        formula: null,
      });
    }
    const range = { sheet: 0, r0: 9, c0: 5, r1: 13, c1: 6 };
    store.setState((s) => ({
      ...s,
      data: { ...s.data, cells },
      ui: {
        ...s.ui,
        filterRanges: [range],
        filterCriteria: [{ range, byCol: 5, hiddenValues: ['b'] }],
      },
    }));
  });

  it('writes SUBTOTAL when Sigma runs from a number cell in the list', () => {
    const active = { sheet: 0, row: 12, col: 6 };
    store.setState((s) => ({
      ...s,
      selection: { active, anchor: active, range: { sheet: 0, r0: 12, c0: 6, r1: 12, c1: 6 } },
    }));
    expect(autoSum(store.getState(), wb)?.formula).toBe('=SUBTOTAL(9,G11:G14)');
  });

  it('writes SUBTOTAL for a range selection with room below', () => {
    const range = { sheet: 0, r0: 10, c0: 6, r1: 13, c1: 6 };
    store.setState((s) => ({
      ...s,
      selection: {
        active: { sheet: 0, row: 10, col: 6 },
        anchor: { sheet: 0, row: 10, col: 6 },
        range,
      },
    }));
    expect(autoSum(store.getState(), wb)?.formula).toBe('=SUBTOTAL(9,G11:G14)');
  });
});

describe('autoSumSuggestion in and below a filtered list', () => {
  let store: SpreadsheetStore;

  const setup = (
    cells: Record<string, { n: number } | { s: string }>,
    list: { r0: number; r1: number },
    selection: { r0: number; c0: number; r1: number; c1: number },
  ): void => {
    store = createSpreadsheetStore();
    const map = new Map(store.getState().data.cells);
    for (const [key, v] of Object.entries(cells)) {
      const [row, col] = key.split(',').map(Number) as [number, number];
      map.set(addrKey({ sheet: 0, row, col }), {
        value: 'n' in v ? { kind: 'number', value: v.n } : { kind: 'text', value: v.s },
        formula: null,
      });
    }
    const range = { sheet: 0, r0: list.r0, c0: 0, r1: list.r1, c1: 4 };
    const active = { sheet: 0, row: selection.r0, col: selection.c0 };
    store.setState((s) => ({
      ...s,
      data: { ...s.data, cells: map },
      selection: { active, anchor: active, range: { sheet: 0, ...selection } },
      ui: {
        ...s.ui,
        filterRanges: [range],
        filterCriteria: [{ range, byCol: 4, hiddenValues: ['x'] }],
      },
    }));
  };

  it('puts the formula on the active number cell over the numbers above it', () => {
    setup(
      { '1,2': { n: 12 }, '2,2': { n: 5 }, '3,2': { n: 30 }, '4,2': { n: 8 } },
      { r0: 0, r1: 10 },
      {
        r0: 4,
        c0: 2,
        r1: 4,
        c1: 2,
      },
    );
    expect(autoSumSuggestion(store.getState())).toEqual({
      addr: { sheet: 0, row: 4, col: 2 },
      formula: '=SUBTOTAL(9,C2:C4)',
    });
  });

  it('stops at the gap above the active number cell', () => {
    setup(
      { '1,2': { n: 12 }, '2,2': { n: 5 }, '3,2': { n: 30 }, '6,2': { n: 40 } },
      { r0: 0, r1: 11 },
      {
        r0: 3,
        c0: 2,
        r1: 3,
        c1: 2,
      },
    );
    expect(autoSumSuggestion(store.getState())?.formula).toBe('=SUBTOTAL(9,C2:C3)');
  });

  it('opens an empty SUBTOTAL for a row selected below the list', () => {
    setup({ '11,0': { s: 'Lemon' } }, { r0: 0, r1: 11 }, { r0: 12, c0: 0, r1: 12, c1: 4 });
    expect(autoSumSuggestion(store.getState())).toEqual({
      addr: { sheet: 0, row: 12, col: 0 },
      formula: '=SUBTOTAL(9,)',
    });
  });
});
