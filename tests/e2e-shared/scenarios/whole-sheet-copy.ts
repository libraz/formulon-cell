import { expect, type Page } from '@playwright/test';
import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/api/spreadsheet/instance.js';
import { expectSystemClipboardText } from '../pages/system-clipboard.js';
import { UserJourneyPage } from '../pages/UserJourneyPage.js';

export async function runWholeSheetCopy(page: Page, platform: 'default' | 'mac'): Promise<void> {
  const sheet = new UserJourneyPage(page);
  await sheet.mount({ locale: 'en', platform });
  await sheet.expectNoStub();
  await page.evaluate(() => {
    const instance = (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst;
    const wb = instance.workbook;
    if (wb.addSheet('Target').index !== 1) throw new Error('Target worksheet was not created');
    wb.batch(() => {
      wb.sheet(0).setNumber('D5', 2);
      wb.sheet(0).setFormula('E5', '=D5*3');
      wb.sheet(1).setNumber('Z20', 9);
    });
    instance.history.clear();
  });

  await page
    .locator('.fc-host__grid')
    .first()
    .click({ position: { x: 10, y: 10 } });
  await expect
    .poll(() =>
      page.evaluate(() =>
        (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst.selection.get(),
      ),
    )
    .toMatchObject({
      sheet: 0,
      ranges: [{ firstRow: 0, firstCol: 0, lastRow: 1048575, lastCol: 16383 }],
    });
  await sheet.shortcut('c');
  await expectSystemClipboardText(page, '2\t=D5*3');

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
