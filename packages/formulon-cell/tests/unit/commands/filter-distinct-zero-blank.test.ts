import { describe, expect, it } from 'vitest';
import { distinctFilterItems } from '../../../src/commands/filter-items.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

/** The checklist items of column A over a header and `values`; null leaves the cell empty. */
const items = async (values: (number | string | null)[]) => {
  const wb = await WorkbookHandle.createDefault();
  try {
    wb.setText({ sheet: 0, row: 0, col: 0 }, 'H');
    values.forEach((value, index) => {
      const addr = { sheet: 0, row: index + 1, col: 0 };
      if (typeof value === 'number') wb.setNumber(addr, value);
      else if (typeof value === 'string') wb.setText(addr, value);
    });
    const store = createSpreadsheetStore();
    mutators.replaceCells(store, wb.cells(0));
    const range = { sheet: 0, r0: 0, c0: 0, r1: values.length, c1: 0 };
    return distinctFilterItems(store.getState(), range, 0, 'en-US', wb).map((item) => item.key);
  } finally {
    wb.dispose();
  }
};

describe('filter checklist with zero and blank cells', () => {
  it('lists zero with the numbers and the blanks last, whichever row comes first', async () => {
    expect(await items([5, null, 0])).toEqual(['0', '5', '']);
    expect(await items([null, null, 0, 'a'])).toEqual(['0', 'a', '']);
    expect(await items([0, null])).toEqual(['0', '']);
  });
});
