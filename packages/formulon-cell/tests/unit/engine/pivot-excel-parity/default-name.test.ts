import { describe, expect, it } from 'vitest';
import { nextPivotTableName } from '../../../../src/commands/pivot-table.js';
import { loadFixture } from './fixture.js';

// The desktop app (ja) names reports ピボットテーブルN with the first free N.
describe('default pivot report name', () => {
  it.each([
    // ピボットテーブル1..6 are taken.
    ['pivot_value_date_filters.xlsx', 'ピボットテーブル7'],
    // Only ピボットテーブル2 is taken.
    ['pivot_blank_item.xlsx', 'ピボットテーブル1'],
    // A ja workbook holding only an English-named report.
    ['ja_c0_rows.xlsx', 'ピボットテーブル1'],
  ])('names the next report in %s %s', async (file, name) => {
    const wb = await loadFixture(file);
    try {
      expect(nextPivotTableName(wb)).toBe(name);
    } finally {
      wb.dispose();
    }
  });

  it('uses the English stem outside a Japanese workbook and skips taken names', async () => {
    const wb = await loadFixture('ja_c0_rows.xlsx');
    try {
      wb.setSpreadsheetProfileId('windows-en_US');
      expect(nextPivotTableName(wb)).toBe('PivotTable2');
    } finally {
      wb.dispose();
    }
  });
});
