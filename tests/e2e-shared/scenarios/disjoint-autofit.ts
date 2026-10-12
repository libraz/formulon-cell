import { expect, type Page } from '@playwright/test';
import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/mount/types.js';
import { gridCellCenter, SpreadsheetPage } from '../pages/SpreadsheetPage.js';

export async function runDisjointAutofit(page: Page): Promise<void> {
  const spreadsheet = new SpreadsheetPage(page);
  for (const locale of ['en', 'ja', 'de', 'fr', 'zh', 'ko', 'th']) {
    await spreadsheet.mount({ locale, platform: 'default' });
    await spreadsheet.expectNoStub();
    expect(
      await page.evaluate(() => {
        const instance = (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst;
        return instance.applyChanges([
          { addr: { sheet: 0, row: 0, col: 0 }, input: 'one\ntwo\nthree' },
          { addr: { sheet: 0, row: 2, col: 0 }, input: 'one\ntwo\nthree\nfour' },
          {
            addr: { sheet: 0, row: 0, col: 1 },
            input: 'A long selected column B text for AutoFit',
          },
          {
            addr: { sheet: 0, row: 0, col: 3 },
            input: 'An even longer selected column D text for AutoFit',
          },
        ]).status;
      }),
    ).toBe('applied');

    const selectHeader = async (axis: 'row' | 'col', index: number, extra = false) => {
      const box = await page.locator('.fc-host__grid').first().boundingBox();
      if (!box) throw new Error('Missing grid');
      const cell = await gridCellCenter(
        page,
        axis === 'row' ? index : 0,
        axis === 'col' ? index : 0,
      );
      if (extra) await page.keyboard.down('Control');
      try {
        await page.mouse.click(
          axis === 'row' ? box.x + 10 : cell.x,
          axis === 'row' ? cell.y : box.y + 10,
        );
      } finally {
        if (extra) await page.keyboard.up('Control');
      }
    };
    const runFit = async (axis: 'row' | 'col') => {
      await page.locator('[data-ribbon-command="formatCellsHome"]').first().click();
      await page.locator(`[data-cell-format="${axis}-autofit"]:visible`).click();
    };
    const layout = () =>
      page.evaluate(() => {
        const instance = (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst;
        const state = instance.store.getState();
        return {
          rows: [...state.layout.rowHeights],
          cols: [...state.layout.colWidths],
          rowDefault: state.layout.defaultRowHeight,
          colDefault: state.layout.defaultColWidth,
          nativeRows: instance.workbook.getRowLayouts(0),
          nativeCols: instance.workbook.getColumnLayouts(0),
        };
      });
    const before = await layout();
    await selectHeader('row', 0);
    await selectHeader('row', 2, true);
    await runFit('row');
    const fittedRows = await layout();
    for (const index of [0, 2]) {
      const height = new Map(fittedRows.rows).get(index);
      expect(height, `${locale} row ${index + 1}`).toBeGreaterThan(before.rowDefault);
      expect(fittedRows.nativeRows.find((row) => row.row === index)?.height).toBeCloseTo(
        height ?? 0,
      );
    }
    expect(new Map(fittedRows.rows).get(1)).toBe(new Map(before.rows).get(1));
    expect(
      await page.evaluate(() =>
        (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst.undo(),
      ),
    ).toBe(true);
    expect((await layout()).rows).toEqual(before.rows);
    expect(
      await page.evaluate(() =>
        (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst.redo(),
      ),
    ).toBe(true);
    expect((await layout()).rows).toEqual(fittedRows.rows);
    await selectHeader('col', 1);
    await selectHeader('col', 3, true);
    await runFit('col');
    const fittedCols = await layout();
    for (const index of [1, 3]) {
      const width = new Map(fittedCols.cols).get(index);
      expect(width, `${locale} column ${index + 1}`).toBeGreaterThan(before.colDefault);
      const native = fittedCols.nativeCols.find((col) => col.first <= index && col.last >= index);
      expect(native?.width).toBeCloseTo(width ?? 0);
    }
    expect(new Map(fittedCols.cols).get(2)).toBe(new Map(before.cols).get(2));
  }
}
