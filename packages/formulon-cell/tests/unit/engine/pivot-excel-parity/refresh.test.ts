import { describe, expect, it } from 'vitest';
import { refreshPivotTable } from '../../../../src/commands/pivot-cache-source.js';
import { resolveXfForFormat } from '../../../../src/engine/cell-xf-writeback.js';
import { formatCodeToNumFmt } from '../../../../src/engine/format-writeback.js';
import { column, columnFormats, loadFixture, pivotGrid } from './fixture.js';

describe('pivot refresh matches the desktop app', () => {
  it('holds source edits until refresh, then appends the new item after the old ones', async () => {
    const wb = await loadFixture('ja_c0_rows.xlsx');
    try {
      const source = wb.getPivotCacheWorksheetSource(wb.pivotTableCacheId(0, 0));
      const sheet = source?.range?.sheet;
      expect(sheet).toBeDefined();
      // The measured edit: A2 North/A 100/1 becomes a new region.
      const at = { sheet: sheet as number, row: 1, col: 0 };
      const before = wb.getValue(at);
      expect(before.kind).toBe('text');
      wb.setText(at, 'VeryLongRegionNameForWidth');
      const numFmt = formatCodeToNumFmt('#,##0.00') ?? undefined;
      const xf = resolveXfForFormat(wb, { numFmt }, new Map(), new Map(), false, null).xfIndex;
      expect(wb.setCellXfIndex(at.sheet, 1, 1, xf)).toBe(true);
      wb.recalc();
      expect(column(pivotGrid(wb), 'A')).toEqual(['行ラベル', 'East', 'North', 'South', '総計']);
      expect(refreshPivotTable(wb, { sheet: 0, pivotIndex: 0 }).ok).toBe(true);
      expect(pivotGrid(wb)).toMatchObject({
        A4: 'East',
        B4: 500,
        C4: 5,
        A5: 'North',
        B5: 200,
        C5: 2,
        A6: 'South',
        B6: 700,
        C6: 7,
        A7: 'VeryLongRegionNameForWidth',
        B7: 100,
        C7: 1,
        A8: '総計',
        B8: 1500,
        C8: 15,
      });
      // Reverting the edit drops the new item again.
      wb.setText(at, 'North');
      wb.recalc();
      expect(refreshPivotTable(wb, { sheet: 0, pivotIndex: 0 }).ok).toBe(true);
      expect(column(pivotGrid(wb), 'A')).toEqual(['行ラベル', 'East', 'North', 'South', '総計']);
      // Source formats do not reach the report.
      expect(new Set(columnFormats(wb, 1))).toEqual(new Set(['']));
    } finally {
      wb.dispose();
    }
  });
});
