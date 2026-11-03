import { describe, expect, it } from 'vitest';
import { PivotCalendar, PivotDateGrouping } from '../../../../src/engine/types.js';
import { column, loadFixture, pivotGrid } from './fixture.js';

// Sheet5 rows Date 2024-01-15 100, 01-25 80, 02-08 150, 02-20 50, 03-10 200;
// the desktop app buckets them as below. Bucket labels are the engine's own.
describe('pivot date grouping matches the desktop app', () => {
  it.each([
    [PivotDateGrouping.Month, [180, 200, 200, 580]],
    [PivotDateGrouping.Quarter, [580, 580]],
    [PivotDateGrouping.Year, [580, 580]],
    [PivotDateGrouping.Day, [100, 80, 150, 50, 200, 580]],
  ])('groups by %i', async (granularity, sums) => {
    const wb = await loadFixture('pivot_value_date_filters.xlsx');
    try {
      const sheet = Array.from({ length: wb.sheetCount }, (_, i) => i).find(
        (i) => wb.sheetName(i) === 'Sheet5',
      ) as number;
      const date = wb.getPivotField(sheet, 0, 2);
      expect(date?.sourceName).toBe('Date');
      expect(wb.setPivotFieldDateGroup(sheet, 0, 2, granularity, PivotCalendar.Gregorian)).toBe(
        true,
      );
      const grid = pivotGrid(wb, sheet);
      expect(column(grid, 'B').slice(1)).toEqual(sums);
      expect(column(grid, 'A').at(-1)).toBe('総計');
    } finally {
      wb.dispose();
    }
  });
});
