import { describe, expect, it } from 'vitest';
import { createPivotTableFromRange } from '../../../../src/commands/pivot-table.js';
import { PivotShowValuesAs } from '../../../../src/engine/types.js';
import { WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import { columnFormats } from './fixture.js';

const build = async (
  showValuesAs: PivotShowValuesAs,
  numberFormat?: string,
): Promise<WorkbookHandle> => {
  const wb = await WorkbookHandle.createDefault();
  const rows: [string, number][] = [
    ['East', 500],
    ['North', 300],
    ['South', 700],
  ];
  wb.setText({ sheet: 0, row: 0, col: 0 }, 'Region');
  wb.setText({ sheet: 0, row: 0, col: 1 }, 'Amt');
  rows.forEach(([name, n], i) => {
    wb.setText({ sheet: 0, row: i + 1, col: 0 }, name);
    wb.setNumber({ sheet: 0, row: i + 1, col: 1 }, n);
  });
  const created = createPivotTableFromRange(wb, {
    source: { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 1 },
    destination: { sheet: 0, row: 6, col: 0 },
    rowField: 'Region',
    valueField: 'Amt',
    valueFieldSettings: [{ fieldName: 'Amt', showValuesAs, numberFormat }],
  });
  expect(created.ok).toBe(true);
  return wb;
};

describe('show values as number format', () => {
  // Desktop app: % of total / column / row set "0.00%" on the data column;
  // Index stays General.
  it.each([
    [PivotShowValuesAs.PercentOfTotal, '0.00%'],
    [PivotShowValuesAs.PercentOfCol, '0.00%'],
    [PivotShowValuesAs.PercentOfRow, '0.00%'],
    [PivotShowValuesAs.Index, ''],
    [PivotShowValuesAs.Normal, ''],
  ])('show-as %i formats values as "%s"', async (showValuesAs, format) => {
    const wb = await build(showValuesAs);
    try {
      expect(columnFormats(wb, 1).slice(1)).toEqual([format, format, format, format]);
    } finally {
      wb.dispose();
    }
  });

  it('keeps an explicit format over the percent default', async () => {
    const wb = await build(PivotShowValuesAs.PercentOfTotal, '0%');
    try {
      expect(columnFormats(wb, 1).slice(1)).toEqual(['0%', '0%', '0%', '0%']);
    } finally {
      wb.dispose();
    }
  });
});
