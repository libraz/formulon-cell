import { afterEach, describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import { createPivotTableFromRange } from '../../../src/commands/pivot-table.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { en } from '../../../src/i18n/strings/en.js';
import { openPivotFieldSettingsAt } from '../../../src/interact/pivot-field-settings-dialog-apply.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

describe('Field Settings on report totals', () => {
  afterEach(() => {
    document.body.replaceChildren();
  });

  it('opens on the data field of a total value cell, not on a total label', async () => {
    const wb = await WorkbookHandle.createDefault();
    const rows = [
      ['Region', 'Product', 'Sales'],
      ['East', 'A', 1],
      ['East', 'B', 2],
      ['West', 'A', 3],
    ];
    for (const [r, row] of rows.entries()) {
      for (const [c, v] of row.entries()) {
        const addr = { sheet: 0, row: r, col: c };
        if (typeof v === 'number') wb.setNumber(addr, v);
        else wb.setText(addr, v);
      }
    }
    createPivotTableFromRange(wb, {
      source: { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 2 },
      destination: { sheet: 0, row: 7, col: 0 },
      rowField: 'Region',
      columnField: 'Product',
      valueField: 'Sales',
    });
    const host = document.createElement('div');
    document.body.appendChild(host);
    const ctx = {
      wb,
      host,
      strings: en,
      history: new History(),
      store: createSpreadsheetStore(),
      origin: 'contextMenu' as const,
      commandId: 'pivotFieldSettings',
      onChanged: () => {},
    };
    // Grand total row (11,3 and 11,1) and the grand total column (9,3) are value cells.
    expect(openPivotFieldSettingsAt(ctx, { sheet: 0, row: 11, col: 3 })).toBe(true);
    expect(openPivotFieldSettingsAt(ctx, { sheet: 0, row: 9, col: 3 })).toBe(true);
    // The Grand Total caption names no field.
    expect(openPivotFieldSettingsAt(ctx, { sheet: 0, row: 11, col: 0 })).toBe(false);
    wb.dispose();
  });
});
