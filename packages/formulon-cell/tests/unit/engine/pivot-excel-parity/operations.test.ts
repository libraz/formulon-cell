import { describe, expect, it } from 'vitest';
import {
  PivotAggregation,
  PivotShowValuesAs,
  PivotSortOrder,
} from '../../../../src/engine/types.js';
import { WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import { column, columnFormats, loadFixture, pivotGrid } from './fixture.js';

const withFixture = async (
  file: string,
  run: (wb: WorkbookHandle) => void | Promise<void>,
): Promise<void> => {
  const wb = await loadFixture(file);
  try {
    await run(wb);
  } finally {
    wb.dispose();
  }
};

const setValueField = (
  wb: WorkbookHandle,
  patch: {
    aggregation?: PivotAggregation;
    showValuesAs?: PivotShowValuesAs;
    numberFormat?: string;
  },
): void => {
  const current = wb.getPivotDefinition(0, 0)?.dataFields[0];
  if (!current) throw new Error('no data field');
  const { showAsBaseField: _f, showAsBaseItem: _i, ...rest } = current;
  const renamed = patch.aggregation !== undefined && patch.aggregation !== current.aggregation;
  expect(
    wb.setPivotDataField(0, 0, 0, {
      ...rest,
      ...(renamed ? { name: '' } : {}),
      ...patch,
    }),
  ).toBe(true);
};

// pivot_blank_item: North 100 + 50, (空白) 200, South 300.
const FUNCTIONS: [PivotAggregation, string, unknown[]][] = [
  [PivotAggregation.Sum, '合計 / Amt', [150, 300, 200, 650]],
  [PivotAggregation.Count, '個数 / Amt', [2, 1, 1, 4]],
  [PivotAggregation.CountNumbers, '数値の個数 / Amt', [2, 1, 1, 4]],
  [PivotAggregation.Max, '最大 / Amt', [100, 300, 200, 300]],
  [PivotAggregation.Min, '最小 / Amt', [50, 300, 200, 50]],
  [PivotAggregation.Product, '積 / Amt', [5000, 300, 200, 300000000]],
  [PivotAggregation.Average, '平均 / Amt', [75, 300, 200, 162.5]],
  [PivotAggregation.StdDevP, '標準偏差 / Amt', [25, 0, 0, 96.0143218483576]],
  [PivotAggregation.VarP, '分散 / Amt', [625, 0, 0, 9218.75]],
];

describe('pivot operations match the desktop app', () => {
  it.each(FUNCTIONS)('summarizes by %i as %s', (aggregation, caption, values) =>
    withFixture('pivot_blank_item.xlsx', (wb) => {
      setValueField(wb, { aggregation });
      const grid = pivotGrid(wb);
      expect(grid.B3).toBe(caption);
      expect([grid.B4, grid.B5, grid.B6, grid.B7]).toEqual(values);
    }),
  );

  it.each([
    [PivotAggregation.StdDev, '標本標準偏差 / Amt', 35.3553390593, 110.8677891],
    [PivotAggregation.Var, '標本分散 / Amt', 1250, 12291.6666667],
  ])('keeps the sample statistic %i of a two-value group', (aggregation, caption, north, total) =>
    withFixture('pivot_blank_item.xlsx', (wb) => {
      setValueField(wb, { aggregation });
      const grid = pivotGrid(wb);
      expect(grid.B3).toBe(caption);
      expect(grid.B4).toBeCloseTo(north, 6);
      expect(grid.B7).toBeCloseTo(total, 6);
    }),
  );

  it.each([
    [PivotShowValuesAs.PercentOfTotal, [1 / 3, 1 / 3, 0.2, 1 / 15, 2 / 15, 7 / 15, 0.2, 4 / 15, 1]],
    [PivotShowValuesAs.PercentOfCol, [1 / 3, 1 / 3, 0.2, 1 / 15, 2 / 15, 7 / 15, 0.2, 4 / 15, 1]],
    [PivotShowValuesAs.PercentOfRow, [1, 1, 1, 1, 1, 1, 1, 1, 1]],
    [PivotShowValuesAs.Index, [1, 1, 1, 1, 1, 1, 1, 1, 1]],
  ])('shows values as %i', (showValuesAs, values) =>
    withFixture('ja_d7_two_rows_values_cols.xlsx', (wb) => {
      setValueField(wb, { showValuesAs });
      const grid = pivotGrid(wb);
      expect(grid.B3).toBe('Sum of Amt');
      const shown = column(grid, 'B').slice(1) as number[];
      shown.forEach((v, i) => {
        expect(v).toBeCloseTo(values[i] as number, 9);
      });
    }),
  );

  it('formats the data column with its number format, leaving the others General', () =>
    withFixture('ja_c0_rows.xlsx', (wb) => {
      setValueField(wb, { numberFormat: '#,##0' });
      expect(columnFormats(wb, 1).slice(1)).toEqual(['#,##0', '#,##0', '#,##0', '#,##0']);
      expect(columnFormats(wb, 2).slice(1)).toEqual(['', '', '', '']);
    }));

  it.each([
    [
      'descending by Sum of Amt',
      PivotSortOrder.Descending,
      'Sum of Amt',
      ['South', 'A', 'B', 'East', 'A', 'North', 'A', 'B'],
    ],
    [
      'ascending by label',
      PivotSortOrder.Ascending,
      '',
      ['East', 'A', 'North', 'A', 'B', 'South', 'A', 'B'],
    ],
    [
      'descending by label',
      PivotSortOrder.Descending,
      '',
      ['South', 'A', 'B', 'North', 'A', 'B', 'East', 'A'],
    ],
  ] as const)('sorts Region %s', (_name, order, by, labels) =>
    withFixture('ja_d7_two_rows_values_cols.xlsx', (wb) => {
      expect(wb.setPivotFieldSort(0, 0, 0, order, by)).toBe(true);
      expect(column(pivotGrid(wb), 'A').slice(1, -1)).toEqual(labels);
    }),
  );

  it('keeps a sort by values through a save and reload', () =>
    withFixture('ja_d7_two_rows_values_cols.xlsx', async (wb) => {
      expect(wb.setPivotFieldSort(0, 0, 0, PivotSortOrder.Ascending, 'Sum of Amt')).toBe(true);
      const labels = ['North', 'A', 'B', 'East', 'A', 'South', 'A', 'B'];
      expect(column(pivotGrid(wb), 'A').slice(1, -1)).toEqual(labels);
      const reloaded = await WorkbookHandle.loadBytes(wb.saveWithDiagnostics('xlsx').bytes);
      try {
        expect(reloaded.getPivotField(0, 0, 0)).toMatchObject({
          sortOrder: PivotSortOrder.Ascending,
          sortByField: 'Sum of Amt',
        });
        expect(column(pivotGrid(reloaded), 'A').slice(1, -1)).toEqual(labels);
      } finally {
        reloaded.dispose();
      }
    }));

  it('drops a hidden item and recomputes the grand total', () =>
    withFixture('ja_d7_two_rows_values_cols.xlsx', (wb) => {
      const north = wb.getPivotField(0, 0, 0)?.items.findIndex((i) => i.label === 'North') ?? -1;
      expect(wb.setPivotFieldItemVisible(0, 0, 0, north, false)).toBe(true);
      const grid = pivotGrid(wb);
      expect(column(grid, 'A').slice(1)).toEqual(['East', 'A', 'South', 'A', 'B', '総計']);
      expect([grid.B9, grid.C9]).toEqual([1200, 12]);
    }));

  it('drops the bottom grand total with column grand totals off', () =>
    withFixture('ja_d10_tabular_two_rows.xlsx', (wb) => {
      expect(wb.setPivotTableGrandTotals(0, 0, true, false)).toBe(true);
      expect(column(pivotGrid(wb), 'A')).not.toContain('総計');
      expect(wb.setPivotTableGrandTotals(0, 0, false, true)).toBe(true);
      expect(column(pivotGrid(wb), 'A').at(-1)).toBe('総計');
    }));

  it.each([
    [
      [PivotAggregation.Sum, PivotAggregation.Count],
      ['North 合計', 'North 個数'],
      [300, 2],
    ],
    [
      [PivotAggregation.Average, PivotAggregation.Max],
      ['North 平均', 'North 最大'],
      [150, 200],
    ],
  ])('labels custom subtotals "<item> <function>" (%o)', (fns, labels, values) =>
    withFixture('ja_d10_tabular_two_rows.xlsx', (wb) => {
      wb.clearPivotFieldSubtotalFns(0, 0, 0);
      for (const fn of fns) wb.addPivotFieldSubtotalFn(0, 0, 0, fn);
      const grid = pivotGrid(wb);
      const rows = Object.entries(grid).filter(
        ([ref, v]) => ref.startsWith('A') && labels.includes(v as string),
      );
      expect(rows.map(([, v]) => v)).toEqual(labels);
      expect(rows.map(([ref]) => grid[`C${ref.slice(1)}`])).toEqual(values);
    }),
  );
});
