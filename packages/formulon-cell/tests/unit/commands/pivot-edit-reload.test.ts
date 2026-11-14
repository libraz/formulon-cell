import { describe, expect, it, vi } from 'vitest';
import {
  createPivotTableFromRange,
  refreshPivotReportsAfterEdit,
} from '../../../src/commands/pivot-table.js';
import { pivotStyleLayer } from '../../../src/engine/pivot-style.js';
import { PivotAggregation } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

/** Kind / Amt at K1:L3; report (rows Kind) at A1. */
const setup = async () => {
  const wb = await WorkbookHandle.createDefault();
  const rows: (string | number)[][] = [
    ['Kind', 'Amt'],
    ['a', 1],
    ['b', 2],
  ];
  for (const [r, row] of rows.entries()) {
    for (const [c, v] of row.entries()) {
      const addr = { sheet: 0, row: r, col: 10 + c };
      if (typeof v === 'number') wb.setNumber(addr, v);
      else wb.setText(addr, v);
    }
  }
  const created = createPivotTableFromRange(wb, {
    source: { sheet: 0, r0: 0, c0: 10, r1: 2, c1: 11 },
    destination: { sheet: 0, row: 0, col: 0 },
    rowField: 'Kind',
    valueField: 'Amt',
    aggregation: PivotAggregation.Sum,
  });
  expect(created.ok).toBe(true);
  const store = createSpreadsheetStore();
  mutators.replaceCells(store, wb.cells(0));
  pivotStyleLayer(wb, store.getState());
  return { wb, store };
};

describe('refreshPivotReportsAfterEdit', () => {
  it('reloads the sheet cells into the store once per edit', async () => {
    const { wb, store } = await setup();
    wb.setNumber({ sheet: 0, row: 1, col: 11 }, 5);
    const replace = vi.spyOn(mutators, 'replaceCells');
    const refresh = vi.spyOn(mutators, 'refreshActiveSheetCells');
    const read = vi.spyOn(wb, 'cells');
    refreshPivotReportsAfterEdit(store, wb);
    expect(read.mock.calls.length).toBe(1);
    expect(replace.mock.calls.length + refresh.mock.calls.length).toBe(1);
    vi.restoreAllMocks();
    wb.dispose();
  });

  it('leaves the report values in the store after the reload', async () => {
    const { wb, store } = await setup();
    refreshPivotReportsAfterEdit(store, wb);
    const shown = store.getState().data.cells;
    for (const cell of wb.pivotCells(0)) {
      expect(shown.get(`0:${cell.addr.row}:${cell.addr.col}`)?.value).toEqual(cell.value);
    }
    expect(shown.get('0:3:1')?.value).toEqual({ kind: 'number', value: 3 });
    wb.dispose();
  });
});
