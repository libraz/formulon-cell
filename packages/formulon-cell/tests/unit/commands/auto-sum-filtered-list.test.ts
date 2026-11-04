import { beforeEach, describe, expect, it } from 'vitest';
import { autoSum } from '../../../src/commands/auto-sum.js';
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
