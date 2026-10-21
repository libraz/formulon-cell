import { readFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  createPivotTableFromRange,
  executeRibbonPivotTableAction,
  refreshPivotTable,
} from '../../../src/commands/pivot-table.js';
import { PivotAggregation, PivotAxis, PivotReportLayout } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const fixture = resolve(
  dirname(fileURLToPath(import.meta.url)),
  '../../fixtures/charts/chart_pivot.xlsx',
);

const seed = (wb: WorkbookHandle): void => {
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
};

describe('pivot definition read-back', () => {
  it('advertises the capability', async () => {
    const wb = await WorkbookHandle.createDefault();
    expect(wb.capabilities.pivotDefinitionRead).toBe(true);
    wb.dispose();
  });

  it('reads back what createPivotTableFromRange wrote', async () => {
    const wb = await WorkbookHandle.createDefault();
    seed(wb);
    const created = createPivotTableFromRange(wb, {
      source: { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 3 },
      destination: { sheet: 0, row: 8, col: 0 },
      name: 'Report',
      rowField: 'Region',
      columnField: 'Product',
      filterField: 'Sales',
      valueField: 'Qty',
      aggregation: PivotAggregation.Average,
      showRowGrandTotals: false,
      valueNumberFormat: '#,##0.00',
    });
    expect(created.ok).toBe(true);
    wb.setPivotReportLayout(0, 0, PivotReportLayout.Outline);

    const definition = wb.getPivotDefinition(0, 0);
    expect(definition).toMatchObject({
      name: 'Report',
      grandTotals: { rows: false, cols: true },
      layout: PivotReportLayout.Outline,
      rowFieldOrder: [0],
      colFieldOrder: [1],
      dataFields: [
        {
          name: 'Average of Qty',
          fieldIndex: 3,
          aggregation: PivotAggregation.Average,
          numberFormat: '#,##0.00',
        },
      ],
    });
    expect(definition?.fields.map((f) => [f.sourceName, f.axis])).toEqual([
      ['Region', PivotAxis.Row],
      ['Product', PivotAxis.Col],
      ['Sales', PivotAxis.Page],
      ['Qty', PivotAxis.Value],
    ]);
    wb.dispose();
  });

  it('binds a lone value field to its own cache column', async () => {
    const wb = await WorkbookHandle.createDefault();
    seed(wb);
    createPivotTableFromRange(wb, {
      source: { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 3 },
      destination: { sheet: 0, row: 8, col: 0 },
      rowField: 'Region',
      valueField: 'Sales',
    });
    const numbers = [...wb.pivotCells(0)]
      .filter((c) => c.value.kind === 'number')
      .map((c) => (c.value as { value: number }).value);
    expect(numbers).toEqual([20, 5, 25]);
    // Fields left out of the report keep their cache position.
    expect(wb.getPivotDefinition(0, 0)?.dataFields[0]).toMatchObject({
      name: 'Sum of Sales',
      fieldIndex: 2,
    });
    wb.dispose();
  });

  it('reads a report loaded from a file, including its cache and refresh source', async () => {
    const wb = await WorkbookHandle.loadBytes(new Uint8Array(readFileSync(fixture)));
    const sheet = wb.getPivotTables()[0]?.sheetIndex ?? -1;
    expect(sheet).toBeGreaterThanOrEqual(0);

    const cacheId = wb.pivotTableCacheId(sheet, 0);
    expect(cacheId).toBeGreaterThanOrEqual(0);
    const definition = wb.getPivotDefinition(sheet, 0);
    expect(definition).toMatchObject({
      name: 'PT1',
      cacheId,
      grandTotals: { rows: true, cols: true },
      rowFieldOrder: [0],
      colFieldOrder: [],
    });
    expect(definition?.fields.map((f) => [f.sourceName, f.axis])).toEqual([
      ['Region', PivotAxis.Row],
      ['Item', null],
      ['Sales', PivotAxis.Value],
    ]);
    expect(definition?.fields[0]?.items.map((i) => [i.label, i.visible])).toEqual([
      ['East', true],
      ['North', true],
      ['West', true],
    ]);
    expect(definition?.dataFields).toMatchObject([
      { name: '合計 / Sales', fieldIndex: 2, aggregation: PivotAggregation.Sum },
    ]);

    const first = wb.getPivotTables()[0];
    const info = wb.pivotCellInfo({ sheet, row: (first?.top ?? 0) + 1, col: first?.left ?? 0 });
    expect(info?.cacheId).toBe(cacheId);
    expect(refreshPivotTable(wb, { sheet, pivotIndex: 0 })).toEqual({ ok: true, cacheId });
    wb.dispose();
  });

  it('refreshes a loaded report from the ribbon', async () => {
    const wb = await WorkbookHandle.loadBytes(new Uint8Array(readFileSync(fixture)));
    const first = wb.getPivotTables()[0];
    if (!first) throw new Error('fixture has no report');
    const store = createSpreadsheetStore();
    mutators.setActive(store, { sheet: first.sheetIndex, row: first.top + 1, col: first.left });
    const strings = {
      pivotTable: 'PivotTable',
      pivotTableNewSheet: 'New sheet',
      pivotTableRefreshData: 'Refresh',
      pivotTableRefreshUnavailable: 'Unavailable',
      recommendedPivotTables: 'Recommended',
      pivotAuthoringDetail: 'Detail',
      workbookStructureProtectedBlocked: 'Protected',
    };
    expect(
      executeRibbonPivotTableAction({ store, workbook: wb, action: 'refresh', strings }),
    ).toEqual({ kind: 'refreshed', sheet: first.sheetIndex });
    wb.dispose();
  });
});
