import { describe, expect, it, vi } from 'vitest';
import { copy, planCopyRanges } from '../../../../src/commands/clipboard/copy.js';
import type { Range } from '../../../../src/engine/types.js';
import { createSpreadsheetStore, mutators } from '../../../../src/store/store.js';

const fullSheet: Range = {
  sheet: 0,
  r0: 0,
  c0: 0,
  r1: 1_048_575,
  c1: 16_383,
};

describe('clipboard copy range planning', () => {
  it('returns the same bounded copy metadata without reading cell values', () => {
    const store = createSpreadsheetStore();
    mutators.setRange(store, { sheet: 0, r0: 4, c0: 2, r1: 4, c1: 2 });
    mutators.setCell(store, { sheet: 0, row: 4, col: 2 }, { kind: 'text', value: 'source' }, null);
    mutators.setCellFormat(store, { sheet: 0, row: 9, col: 3 }, { bold: true });
    const state = store.getState();
    const read = vi.spyOn(state.data.cells, 'get');
    let plan: ReturnType<typeof planCopyRanges>;
    try {
      plan = planCopyRanges(state);
      expect(read).not.toHaveBeenCalled();
    } finally {
      read.mockRestore();
    }
    expect(plan).toEqual({
      range: { sheet: 0, r0: 4, c0: 2, r1: 4, c1: 2 },
      logicalRange: { sheet: 0, r0: 4, c0: 2, r1: 4, c1: 2 },
    });

    const result = copy(state);
    expect(result).toEqual({ ...plan, tsv: 'source' });
  });

  it('retains whole-band, full-sheet, disjoint and invalid copy shapes', () => {
    const store = createSpreadsheetStore();
    mutators.setCell(store, { sheet: 0, row: 2, col: 3 }, { kind: 'number', value: 7 }, null);
    mutators.setCellFormat(store, { sheet: 0, row: 10, col: 3 }, { italic: true });

    mutators.setRange(store, { sheet: 0, r0: 0, c0: 3, r1: 1_048_575, c1: 3 });
    const column = planCopyRanges(store.getState());
    expect(column?.range).toEqual({ sheet: 0, r0: 0, c0: 3, r1: 1_048_575, c1: 3 });
    expect(column?.payloadRanges).toEqual([{ sheet: 0, r0: 2, c0: 3, r1: 10, c1: 3 }]);

    mutators.setRange(store, fullSheet);
    expect(planCopyRanges(store.getState())?.payloadRanges).toEqual([
      { sheet: 0, r0: 2, c0: 3, r1: 10, c1: 3 },
    ]);

    mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 2 });
    store.setState((state) => ({
      ...state,
      selection: {
        ...state.selection,
        extraRanges: [{ sheet: 0, r0: 2, c0: 0, r1: 2, c1: 2 }],
      },
    }));
    expect(planCopyRanges(store.getState())).toMatchObject({
      range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 2 },
      ranges: [
        { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 2 },
        { sheet: 0, r0: 2, c0: 0, r1: 2, c1: 2 },
      ],
      payloadRanges: [
        { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 2 },
        { sheet: 0, r0: 2, c0: 0, r1: 2, c1: 2 },
      ],
    });

    const invalidStore = createSpreadsheetStore();
    invalidStore.setState((state) => ({
      ...state,
      selection: {
        ...state.selection,
        range: { ...state.selection.range, sheet: 1 },
      },
    }));
    expect(planCopyRanges(invalidStore.getState())).toBeNull();

    const overCapStore = createSpreadsheetStore();
    mutators.setRange(overCapStore, { sheet: 0, r0: 0, c0: 0, r1: 1000, c1: 1000 });
    expect(planCopyRanges(overCapStore.getState())).toBeNull();

    const emptyStore = createSpreadsheetStore();
    mutators.setRange(emptyStore, fullSheet);
    expect(planCopyRanges(emptyStore.getState())?.payloadRanges).toEqual([
      { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
    ]);
  });
});
