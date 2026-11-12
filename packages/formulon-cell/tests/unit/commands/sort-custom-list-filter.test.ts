import { describe, expect, it } from 'vitest';
import { applyValueFilter } from '../../../src/commands/filter.js';
import { sortRange } from '../../../src/commands/sort.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

describe('sorting a filtered range by a custom list', () => {
  it('remembers the custom-list level with the filter', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      ['Size', 'High', 'Low', 'Mid'].forEach((value, row) => {
        wb.setText({ sheet: 0, row, col: 0 }, value);
      });
      const store = createSpreadsheetStore();
      mutators.replaceCells(store, wb.cells(0));
      const range = { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 0 };
      applyValueFilter(store.getState(), store, range, 0, [], wb);
      const list = ['Low', 'Mid', 'High'];
      sortRange(store.getState(), store, wb, range, {
        byCol: 0,
        direction: 'asc',
        hasHeader: true,
        keys: [{ byCol: 0, direction: 'asc', sortOn: 'customList', customList: list }],
      });
      expect([1, 2, 3].map((row) => wb.getValue({ sheet: 0, row, col: 0 }))).toEqual(
        list.map((value) => ({ kind: 'text', value })),
      );
      expect(store.getState().ui.filterSorts?.[0]?.keys).toEqual([
        { byCol: 0, direction: 'asc', sortOn: 'customList', customList: list },
      ]);
    } finally {
      wb.dispose();
    }
  });
});
