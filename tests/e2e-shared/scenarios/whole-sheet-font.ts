import { expect, type Page } from '@playwright/test';
import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/mount/types.js';
import { gridCellCenter, SpreadsheetPage } from '../pages/SpreadsheetPage.js';

export async function runWholeSheetFont(page: Page, platform: 'default' | 'mac'): Promise<void> {
  const spreadsheet = new SpreadsheetPage(page);
  await spreadsheet.mount({ locale: 'ja', platform });
  await spreadsheet.expectNoStub();
  const clickA1 = async (): Promise<void> => {
    const point = await gridCellCenter(page, 0, 0);
    await page.mouse.click(point.x, point.y);
  };
  await clickA1();
  await page.keyboard.press('F2');
  await page.keyboard.type('全セルのフォント');
  await page.keyboard.press('Enter');
  expect(
    await page.evaluate(() =>
      (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst.workbook.getValue({
        sheet: 0,
        row: 0,
        col: 0,
      }),
    ),
  ).toEqual({ kind: 'text', value: '全セルのフォント' });
  const nativeSize = (row = 0, col = 0) =>
    page.evaluate(
      ({ row, col }) => {
        const instance = (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst;
        const effective = instance.workbook.getEffectiveStyle(0, row, col);
        const xf = effective && instance.workbook.getCellXf(effective.xfIndex);
        return xf && instance.workbook.getFontRecord(xf.fontIndex)?.size;
      },
      { row, col },
    );
  const originalSize = await nativeSize();
  // A new ja workbook uses 游ゴシック 11 on the default profile, 12 on Mac.
  expect(originalSize).toBe(platform === 'mac' ? 12 : 11);

  // The top-left header corner is the user's select-all entry point.
  await page
    .locator('.fc-host__grid')
    .first()
    .click({ position: { x: 10, y: 10 } });
  const range = await page.evaluate(
    () =>
      (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst.store.getState().selection
        .range,
  );
  expect(range).toMatchObject({ r0: 0, c0: 0, r1: 1048575, c1: 16383 });
  const sizePicker = page.locator('[data-ribbon-command="fontSize"] button').first();
  await sizePicker.click();
  await page.locator('.fc-tb__rb-dd__opt[data-value="18"]').click();
  expect(await nativeSize()).toBe(18);
  expect(await nativeSize(1048575, 16383)).toBe(18);
  await expect(page.locator('[data-ribbon-select="fontSize"] .fc-tb__rb-dd__value')).toHaveText(
    '18',
  );

  expect(
    await page.evaluate(() =>
      (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst.undo(),
    ),
  ).toBe(true);
  expect(await nativeSize()).toBe(originalSize);
  expect(
    await page.evaluate(() =>
      (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst.redo(),
    ),
  ).toBe(true);
  expect(await nativeSize()).toBe(18);

  const familyPicker = page.locator('[data-ribbon-command="fontFamily"] button').first();
  await familyPicker.click();
  const current = page.locator('.fc-tb__rb-dd__opt[data-value="游ゴシック"]');
  await expect(current).toContainText('游ゴシック Regular');
  await expect(page.locator('.fc-tb__rb-dd__submenu')).toHaveCount(0);
  await current.click();
  await expect(familyPicker).toContainText('游ゴシック Regular');
  await familyPicker.click();
  await expect(current).toHaveAttribute('aria-selected', 'true');
  await current.click();

  // A later direct choice must override the sheet font on just one cell.
  await clickA1();
  await sizePicker.click();
  await page.locator('.fc-tb__rb-dd__opt[data-value="24"]').click();
  expect(await nativeSize()).toBe(24);
  expect(await nativeSize(1048575, 16383)).toBe(18);
  await page.keyboard.press('ArrowRight');
  await page.keyboard.type('new cell');
  await page.keyboard.press('Enter');
  expect(await nativeSize(0, 1)).toBe(18);

  const moduleUrl = `/@fs${
    new URL('../../../packages/formulon-cell/src/engine/workbook-handle.ts', import.meta.url)
      .pathname
  }`;
  const reloadedFonts = await page.evaluate(async (moduleUrl) => {
    const instance = (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst;
    const { WorkbookHandle } = await import(moduleUrl);
    const reloaded = await WorkbookHandle.loadBytes(instance.workbook.save());
    try {
      return [
        [0, 0],
        [0, 1],
        [1048575, 16383],
      ].map(([row, col]) => {
        const style = reloaded.getEffectiveStyle(0, row, col);
        const xf = reloaded.getCellXf(style.xfIndex);
        const font = reloaded.getFontRecord(xf.fontIndex);
        return { size: font.size, name: font.name };
      });
    } finally {
      reloaded.dispose();
    }
  }, moduleUrl);
  expect(reloadedFonts).toEqual([
    { size: 24, name: '游ゴシック' },
    { size: 18, name: '游ゴシック' },
    { size: 18, name: '游ゴシック' },
  ]);

  // Re-select one cell before checking the dialog's complete family list.
  await clickA1();
  await page.evaluate(() =>
    (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst.openFormatDialog('font'),
  );
  const dialog = page.locator('.fc-fmtdlg:not([hidden])');
  const families = dialog.locator('.fc-fmtdlg__font-list--family');
  const selected = families.locator('[data-fc-font-family="游ゴシック"]');
  await expect(selected).toHaveAttribute('aria-selected', 'true');
  await expect(selected).toContainText('游ゴシック Regular');
  const last = families.locator('button').last();
  await last.scrollIntoViewIfNeeded();
  await expect(last).toBeInViewport();
  await dialog.locator('.fc-fmtdlg__footer .fc-fmtdlg__btn:not(.fc-fmtdlg__btn--primary)').click();
}
