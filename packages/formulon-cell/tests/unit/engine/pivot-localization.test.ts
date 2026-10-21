import { describe, expect, it } from 'vitest';
import { createPivotTableFromRange } from '../../../src/commands/pivot-table.js';
import { ENGLISH_PIVOT_CAPTIONS, pivotDataFieldName } from '../../../src/engine/pivot-captions.js';
import { PivotAggregation, PivotAxis, PivotReportLayout } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { en, ja } from '../../../src/i18n/strings.js';

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

describe('pivot caption localization', () => {
  it('translates engine captions by kind and leaves item names alone', async () => {
    const wb = await makeWorkbook();
    const result = createPivotTableFromRange(wb, {
      source,
      destination: { sheet: 0, row: 7, col: 0 },
      rowField: 'Region',
      columnField: 'Product',
      valueField: 'Sales',
    });
    expect(result.ok).toBe(true);
    expect([...texts(wb).values()].sort()).toEqual(
      [
        '(blank)',
        'Chair',
        'Column Labels',
        'Desk',
        'East',
        'Grand Total',
        'Grand Total',
        'Row Labels',
        'Sum of Sales',
        'West',
      ].sort(),
    );

    wb.setPivotCaptions(ja.pivotTableDialog.pivotCaptions);
    const localized = [...texts(wb).values()];
    expect(localized).toEqual(
      expect.arrayContaining(['行ラベル', '列ラベル', '総計', '(空白)', 'East', 'West', 'Desk']),
    );
    expect(localized).not.toContain('Row Labels');
    expect(localized).not.toContain('Grand Total');
    // The stored data-field name is not a caption.
    expect(localized).toContain('Sum of Sales');

    wb.setPivotCaptions(null);
    expect(texts(wb).get('8:0')).toBe('Row Labels');
    wb.dispose();
  });

  it('translates subtotal captions', async () => {
    const wb = await makeWorkbook();
    createPivotTableFromRange(wb, {
      source,
      destination: { sheet: 0, row: 7, col: 0 },
      rowField: 'Region',
      columnField: 'Product',
      valueField: 'Sales',
    });
    wb.addPivotField(0, 0, { sourceName: 'Qty', axis: PivotAxis.Row });
    wb.setPivotReportLayout(0, 0, PivotReportLayout.Tabular);
    expect([...texts(wb).values()]).toContain('East Total');
    wb.setPivotCaptions(ja.pivotTableDialog.pivotCaptions);
    expect([...texts(wb).values()]).toContain('East 集計');
    wb.dispose();
  });

  it('names new value fields in the UI language', async () => {
    const wb = await makeWorkbook();
    createPivotTableFromRange(wb, {
      source,
      destination: { sheet: 0, row: 7, col: 0 },
      rowField: 'Region',
      valueField: 'Sales',
      dataFieldNaming: ja.pivotTableDialog.dataFieldNaming,
    });
    expect([...texts(wb).values()]).toContain('合計 / Sales');
    wb.dispose();
  });

  it('formats every aggregation name', () => {
    const naming = en.pivotTableDialog.dataFieldNaming;
    expect(pivotDataFieldName(PivotAggregation.StdDevP, 'Sales', naming)).toBe('StdDevp of Sales');
    expect(pivotDataFieldName(PivotAggregation.CountNumbers, 'Sales')).toBe('Count of Sales');
    const jaNaming = ja.pivotTableDialog.dataFieldNaming;
    expect(pivotDataFieldName(PivotAggregation.StdDev, 'Sales', jaNaming)).toBe(
      '標本標準偏差 / Sales',
    );
    expect(pivotDataFieldName(PivotAggregation.VarP, '$&', jaNaming)).toBe('分散 / $&');
  });

  it('keeps the English table in step with the engine wording', () => {
    expect(en.pivotTableDialog.pivotCaptions).toEqual(ENGLISH_PIVOT_CAPTIONS);
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
