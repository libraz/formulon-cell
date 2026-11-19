import { describe, expect, it } from 'vitest';
import { refreshPivotReportsAfterEdit } from '../../../src/commands/pivot-report-update.js';
import { createPivotTableFromRange } from '../../../src/commands/pivot-table.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

describe('refreshing the reports after a report edit', () => {
  it('reloads the sheet cells once', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      const rows: (string | number)[][] = [
        ['Region', 'Amt'],
        ['East', 12],
        ['West', 5],
      ];
      for (const [r, row] of rows.entries()) {
        for (const [c, value] of row.entries()) {
          const addr = { sheet: 0, row: r, col: c };
          if (typeof value === 'number') wb.setNumber(addr, value);
          else wb.setText(addr, value);
        }
      }
      expect(
        createPivotTableFromRange(wb, {
          source: { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 1 },
          destination: { sheet: 0, row: 6, col: 0 },
          rowField: 'Region',
          valueField: 'Amt',
        }).ok,
      ).toBe(true);
      const store = createSpreadsheetStore();
      mutators.replaceCells(store, wb.cells(0));
      let reloads = 0;
      let cells = store.getState().data.cells;
      const stop = store.subscribe((state) => {
        if (state.data.cells === cells) return;
        cells = state.data.cells;
        reloads += 1;
      });
      refreshPivotReportsAfterEdit(store, wb);
      stop();
      expect(reloads).toBe(1);
      expect(
        [...store.getState().data.cells.values()].some(
          (c) => c.value.kind === 'text' && c.value.value === 'East',
        ),
      ).toBe(true);
    } finally {
      wb.dispose();
    }
  });
});
