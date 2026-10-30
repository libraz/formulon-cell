import { expect, type Page } from '@playwright/test';
import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/mount/types.js';
import { UserJourneyPage } from '../pages/UserJourneyPage.js';

export async function runWholeSheetCopy(page: Page, platform: 'default' | 'mac'): Promise<void> {
  const sheet = new UserJourneyPage(page);
  await sheet.mount({ locale: 'en', platform });
  await sheet.expectNoStub();
  await page.evaluate(() => {
    const instance = (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst;
    const wb = instance.workbook;
    if (wb.addSheet('Target') !== 1) throw new Error('Target worksheet was not created');
    wb.applyCellPatchAtomic([
      { addr: { sheet: 0, row: 4, col: 3 }, value: { kind: 'number', value: 2 }, formula: null },
      { addr: { sheet: 0, row: 4, col: 4 }, value: { kind: 'number', value: 6 }, formula: '=D5*3' },
      { addr: { sheet: 1, row: 19, col: 25 }, value: { kind: 'number', value: 9 }, formula: null },
    ]);
  });

  await page
    .locator('.fc-host__grid')
    .first()
    .click({ position: { x: 10, y: 10 } });
  await expect
    .poll(() =>
      page.evaluate(
        () =>
          (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst.store.getState()
            .selection.range,
      ),
    )
    .toEqual({ sheet: 0, r0: 0, c0: 0, r1: 1048575, c1: 16383 });
  await sheet.shortcut('c');
  await expect.poll(() => page.evaluate(() => navigator.clipboard.readText())).toBe('2\t=D5*3');

  await sheet.chooseSheet('Target');
  await sheet.goTo('A1');
  await sheet.shortcut('v');
  await expect.poll(() => sheet.readValue('D5', 1)).toEqual({ kind: 'number', value: 2 });
  await expect.poll(() => sheet.readValue('E5', 1)).toEqual({ kind: 'number', value: 6 });
  expect(await sheet.readFormula('E5', 1)).toBe('=D5*3');
  expect(await sheet.readValue('A1', 1)).toEqual({ kind: 'blank' });
  expect(await sheet.readValue('Z20', 1)).toEqual({ kind: 'blank' });

  await sheet.shortcut('z');
  await expect.poll(() => sheet.readValue('D5', 1)).toEqual({ kind: 'blank' });
  expect(await sheet.readValue('Z20', 1)).toEqual({ kind: 'number', value: 9 });
  await sheet.shortcut('Shift+z');
  await expect.poll(() => sheet.readValue('D5', 1)).toEqual({ kind: 'number', value: 2 });
  expect(await sheet.readFormula('E5', 1)).toBe('=D5*3');
  expect(await sheet.readValue('Z20', 1)).toEqual({ kind: 'blank' });
  await sheet.expectNoConsoleErrors();
}
