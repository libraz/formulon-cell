import { describe, expect, it } from 'vitest';
import { computeAutofitColWidth } from '../../../src/commands/autofit-measurement.js';
import {
  autofitPivotTableColumns,
  createPivotTableFromRange,
} from '../../../src/commands/pivot-table.js';
import { PivotAggregation } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const setup = async () => {
  const wb = await WorkbookHandle.createDefault();
  const rows: (string | number)[][] = [
    ['Region', 'Amt'],
    ['East', 500],
    ['North', 300],
  ];
  for (const [r, row] of rows.entries()) {
    for (const [c, v] of row.entries()) {
      const addr = { sheet: 0, row: r, col: c };
      if (typeof v === 'number') wb.setNumber(addr, v);
      else wb.setText(addr, v);
    }
  }
  const created = createPivotTableFromRange(wb, {
    source: { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 1 },
    destination: { sheet: 0, row: 5, col: 0 },
    rowField: 'Region',
    valueField: 'Amt',
    aggregation: PivotAggregation.Sum,
  });
  expect(created.ok).toBe(true);
  const store = createSpreadsheetStore();
  mutators.replaceCells(store, wb.cells(0));
  return { wb, store };
};

describe('PivotTable column autofit', () => {
  it('measures the caption with the field button and the style bold', async () => {
    const { wb, store } = await setup();
    const span = { from: 5, to: 5 };
    const styled = computeAutofitColWidth(store.getState(), 0, null, { workbook: wb, span });
    const plain = computeAutofitColWidth(store.getState(), 0, null, { span });
    // The 17 px button plus its inset, on top of the caption itself.
    expect(styled - plain).toBeGreaterThanOrEqual(18);
    wb.dispose();
  });

  it('fits the report columns, shrinking them, and leaves the others alone', async () => {
    const { wb, store } = await setup();
    mutators.setColWidth(store, 1, 300);
    mutators.setColWidth(store, 3, 33);
    autofitPivotTableColumns(store, wb);
    const widths = store.getState().layout.colWidths;
    expect(widths.get(1)).toBeLessThan(300);
    expect(widths.get(0)).toBeGreaterThan(0);
    expect(widths.get(3)).toBe(33);
    wb.dispose();
  });
});
