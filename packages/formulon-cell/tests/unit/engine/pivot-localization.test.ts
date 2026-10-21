import { describe, expect, it } from 'vitest';
import { createPivotTableFromRange } from '../../../src/commands/pivot-table.js';
import { PivotAggregation } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';

const makeWorkbook = async (): Promise<WorkbookHandle> => {
  const wb = await WorkbookHandle.createDefault();
  const rows: (string | number)[][] = [
    ['Region', 'Product', 'Sales', 'Qty'],
    ['East', 'Desk', 12, 1],
    ['East', 'Chair', 8, 2],
    ['West', 'Desk', 5, 3],
    ['West', '', 3, 4],
  ];
  for (const [r, row] of rows.entries()) {
    for (const [c, value] of row.entries()) {
      if (value === '') continue;
      const addr = { sheet: 0, row: r, col: c };
      if (typeof value === 'number') wb.setNumber(addr, value);
      else wb.setText(addr, value);
    }
  }
  return wb;
};

const source = { sheet: 0, r0: 0, c0: 0, r1: 4, c1: 3 };

const texts = (wb: WorkbookHandle): Map<string, string> => {
  const out = new Map<string, string>();
  for (const cell of wb.pivotCells(0)) {
    if (cell.value.kind === 'text') out.set(`${cell.addr.row}:${cell.addr.col}`, cell.value.value);
  }
  return out;
};

describe('pivot captions', () => {
  it('follows the workbook profile, as the engine localizes them', async () => {
    const wb = await makeWorkbook();
    const result = createPivotTableFromRange(wb, {
      source,
      destination: { sheet: 0, row: 7, col: 0 },
      rowField: 'Region',
      columnField: 'Product',
      valueField: 'Sales',
    });
    expect(result.ok).toBe(true);
    expect([...texts(wb).values()]).toEqual(
      expect.arrayContaining(['Row Labels', 'Column Labels', 'Grand Total', '(blank)']),
    );

    expect(wb.setSpreadsheetProfileId('windows-ja_JP')).toBe(true);
    const localized = [...texts(wb).values()];
    expect(localized).toEqual(
      expect.arrayContaining(['行ラベル', '列ラベル', '総計', '(空白)', 'East', 'West', 'Desk']),
    );
    // The stored data-field name is not a caption.
    expect(localized).toContain('Sum of Sales');
    wb.dispose();
  });

  it('names new value fields from the workbook profile', async () => {
    const wb = await makeWorkbook();
    expect(wb.setSpreadsheetProfileId('windows-ja_JP')).toBe(true);
    createPivotTableFromRange(wb, {
      source,
      destination: { sheet: 0, row: 7, col: 0 },
      rowField: 'Region',
      valueField: 'Sales',
    });
    expect([...texts(wb).values()]).toContain('合計 / Sales');
    expect(wb.getPivotDefinition(0, 0)?.dataFields[0]?.name).toBe('合計 / Sales');
    wb.dispose();
  });

  it('names a field per aggregation when the name is omitted', async () => {
    const wb = await makeWorkbook();
    createPivotTableFromRange(wb, {
      source,
      destination: { sheet: 0, row: 7, col: 0 },
      rowField: 'Region',
      valueField: 'Sales',
    });
    const spec = { fieldIndex: 2, aggregation: PivotAggregation.StdDevP };
    expect(wb.addPivotDataField(0, 0, spec)).toBeGreaterThan(0);
    expect(
      wb.addPivotDataField(0, 0, { ...spec, aggregation: PivotAggregation.CountNumbers }),
    ).toBeGreaterThan(0);
    const names = wb.getPivotDefinition(0, 0)?.dataFields.map((d) => d.name);
    expect(names).toEqual(['Sum of Sales', 'StdDevp of Sales', 'Count of Sales']);
    wb.dispose();
  });
});

describe('pivot occupancy', () => {
  it('reports cells inside a report and describes the clicked cell', async () => {
    const wb = await makeWorkbook();
    createPivotTableFromRange(wb, {
      source,
      destination: { sheet: 0, row: 7, col: 0 },
      rowField: 'Region',
      valueField: 'Sales',
    });
    expect(wb.pivotOverlapsRange({ sheet: 0, r0: 7, c0: 0, r1: 7, c1: 0 })).toBe(true);
    expect(wb.pivotOverlapsRange({ sheet: 0, r0: 0, c0: 0, r1: 4, c1: 3 })).toBe(false);
    expect(wb.pivotOverlapsRange({ sheet: 0, r0: 0, c0: 0, r1: 20, c1: 0 })).toBe(true);

    const label = wb.pivotCellInfo({ sheet: 0, row: 8, col: 0 });
    expect(label).toMatchObject({
      pivotIndex: 0,
      fieldName: 'Region',
      fieldIndex: 0,
      isValue: false,
    });
    const value = wb.pivotCellInfo({ sheet: 0, row: 8, col: 1 });
    expect(value?.isValue).toBe(true);
    expect(wb.pivotCellInfo({ sheet: 0, row: 0, col: 0 })).toBeNull();
    wb.dispose();
  });
});
