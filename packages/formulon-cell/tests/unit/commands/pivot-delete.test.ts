import { describe, expect, it } from 'vitest';
import { clearSelectedContents } from '../../../src/commands/clear-contents.js';
import { History } from '../../../src/commands/history.js';
import { setPivotGuardSink } from '../../../src/commands/pivot-guard.js';
import { createPivotTableFromRange } from '../../../src/commands/pivot-table.js';
import { PivotReportLayout } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const setup = async () => {
  const wb = await WorkbookHandle.createDefault();
  const rows: (string | number)[][] = [
    ['Region', 'Product', 'Sales', 'Qty'],
    ['East', 'Desk', 12, 1],
    ['East', 'Chair', 8, 2],
    ['West', 'Desk', 5, 3],
  ];
  for (const [r, row] of rows.entries()) {
    for (const [c, value] of row.entries()) {
      const addr = { sheet: 0, row: r, col: c };
      if (typeof value === 'number') wb.setNumber(addr, value);
      else wb.setText(addr, value);
    }
  }
  const created = createPivotTableFromRange(wb, {
    source: { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 3 },
    destination: { sheet: 0, row: 8, col: 1 },
    name: 'Report',
    rowField: 'Region',
    filterFields: ['Product', 'Sales'],
    valueField: 'Qty',
    showColumnGrandTotals: false,
    valueNumberFormat: '#,##0.00',
  });
  expect(created.ok).toBe(true);
  wb.setPivotReportLayout(0, 0, PivotReportLayout.Tabular);
  const store = createSpreadsheetStore();
  mutators.replaceCells(store, wb.cells(0));
  const bounds = wb.pivotBounds(0)[0]?.range;
  if (!bounds) throw new Error('report has no extent');
  return { wb, store, bounds };
};

const snapshot = (wb: WorkbookHandle) =>
  [...wb.pivotCells(0)].map((c) => [c.addr.row, c.addr.col, c.value, c.numberFormat]);

describe('Delete over a whole PivotTable report', () => {
  it('removes the report as one undo step that rebuilds its definition', async () => {
    const { wb, store, bounds } = await setup();
    const history = new History();
    const alerts: string[] = [];
    setPivotGuardSink(wb, (alert) => alerts.push(alert.message));
    const before = snapshot(wb);
    const definition = wb.getPivotDefinition(0, 0);
    expect(definition).not.toBeNull();

    mutators.setRange(store, {
      sheet: 0,
      r0: bounds.r0 - 1,
      c0: bounds.c0 - 1,
      r1: bounds.r1 + 2,
      c1: bounds.c1 + 2,
    });
    const result = clearSelectedContents({ store, workbook: wb, history });
    expect(result.status).toBe('applied');
    expect(alerts).toEqual([]);
    expect(wb.getPivotTables()).toEqual([]);
    expect([...wb.pivotCells(0)]).toEqual([]);
    expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'text', value: 'Region' });

    expect(history.undo()).toBe(true);
    expect(wb.getPivotTables()).toHaveLength(1);
    expect(wb.getPivotDefinition(0, 0)).toEqual(definition);
    expect(snapshot(wb)).toEqual(before);
    expect(history.undo()).toBe(false);

    expect(history.redo()).toBe(true);
    expect(wb.getPivotTables()).toEqual([]);
    expect(history.undo()).toBe(true);
    expect(snapshot(wb)).toEqual(before);
    wb.dispose();
  });

  it('removes the report without a history', async () => {
    const { wb, store, bounds } = await setup();
    mutators.setRange(store, bounds);
    expect(clearSelectedContents({ store, workbook: wb }).status).toBe('applied');
    expect(wb.getPivotTables()).toEqual([]);
    wb.dispose();
  });

  it('refuses a selection that holds only part of the report', async () => {
    const { wb, store, bounds } = await setup();
    const alerts: string[] = [];
    setPivotGuardSink(wb, (alert) => alerts.push(alert.message));

    mutators.setRange(store, { ...bounds, r1: bounds.r1 - 1 });
    expect(clearSelectedContents({ store, workbook: wb }).status).toBe('rejected');
    expect(alerts).toHaveLength(1);

    // Two ranges that together cover the report remove it.
    mutators.setRange(store, { ...bounds, r1: bounds.r0 });
    mutators.addExtraRange(store, { ...bounds, r0: bounds.r0 + 1 });
    expect(wb.getPivotTables()).toHaveLength(1);
    expect(clearSelectedContents({ store, workbook: wb }).status).toBe('applied');
    expect(wb.getPivotTables()).toEqual([]);
    wb.dispose();
  });
});
