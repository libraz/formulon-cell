import { describe, expect, it } from 'vitest';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';

describe('SUBTOTAL after a row hidden-state change', () => {
  it('recalculates 109 over the rows that are visible now', async () => {
    const wb = await WorkbookHandle.createDefault();
    for (let r = 0; r < 4; r += 1) wb.setNumber({ sheet: 0, row: r, col: 0 }, 10 * (r + 1));
    wb.setFormula({ sheet: 0, row: 5, col: 0 }, '=SUBTOTAL(109,A1:A4)');
    expect(wb.getValue({ sheet: 0, row: 5, col: 0 })).toEqual({ kind: 'number', value: 100 });
    wb.setRowHidden(0, 1, true);
    expect(wb.getValue({ sheet: 0, row: 5, col: 0 })).toEqual({ kind: 'number', value: 80 });
    wb.setRowHidden(0, 1, false);
    expect(wb.getValue({ sheet: 0, row: 5, col: 0 })).toEqual({ kind: 'number', value: 100 });
  });
});
