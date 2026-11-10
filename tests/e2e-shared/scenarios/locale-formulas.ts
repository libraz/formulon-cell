import { expect, type Page } from '@playwright/test';
import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/api/spreadsheet/instance.js';
import { SpreadsheetPage } from '../pages/SpreadsheetPage.js';

export async function runLocaleFormulaScenario(
  page: Page,
  platform: 'mac' | 'default' = 'default',
): Promise<void> {
  const sp = new SpreadsheetPage(page);
  const errors = sp.collectConsoleErrors();
  const readCell = (row: number) =>
    page.evaluate((r) => {
      const instance = (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst;
      const sheet = instance.workbook.sheet(0);
      return {
        formula: sheet.getFormula(r, 0),
        value: sheet.getValue(r, 0),
      };
    }, row);

  for (const profile of [
    { locale: 'de', formula: '=SUMME(1,5;2)' },
    { locale: 'fr', formula: '=SOMME(1,5;2)' },
  ]) {
    await sp.mount({ locale: profile.locale, platform });
    await sp.expectNoStub();
    await sp.typeIntoActiveCell(profile.formula);
    const expected = { formula: '=SUM(1.5,2)', value: 3.5 };
    await expect.poll(() => readCell(0)).toEqual(expected);
    await page.keyboard.press('ArrowUp');
    await expect.poll(() => sp.formulaBarValue()).toBe(profile.formula);
    await page.keyboard.press('F2');
    await expect(page.locator('.fc-host__editor')).toHaveValue(profile.formula);
    await page.keyboard.press('Enter');
    await expect.poll(() => readCell(0)).toEqual(expected);
    await page.keyboard.press('ArrowUp');

    await page.keyboard.press('F2');
    await page.locator('.fc-host__editor').evaluate((element) => {
      const editor = element as HTMLTextAreaElement;
      editor.setSelectionRange(1, editor.value.length);
    });
    await page.keyboard.press('F9');
    await expect(page.locator('.fc-host__editor')).toHaveValue('=3,5');
    await page.keyboard.press('Escape');
    await expect.poll(() => readCell(0)).toEqual(expected);

    const copied = await page.evaluate(() => {
      const data = new DataTransfer();
      document
        .querySelector('.fc-host')
        ?.dispatchEvent(
          new ClipboardEvent('copy', { clipboardData: data, bubbles: true, cancelable: true }),
        );
      return data.getData('text/plain');
    });
    expect(copied).toBe(profile.formula);
    await page.keyboard.press('ArrowDown');
    await page.keyboard.press('ArrowDown');
    // An independently supplied formula exercises the external text path.
    const pastedFormula = profile.formula.replace(';2)', ';3)');
    await page.evaluate((text) => {
      const data = new DataTransfer();
      data.setData('text/plain', text);
      document
        .querySelector('.fc-host')
        ?.dispatchEvent(
          new ClipboardEvent('paste', { clipboardData: data, bubbles: true, cancelable: true }),
        );
    }, pastedFormula);
    await expect
      .poll(() => readCell(2))
      .toEqual({
        formula: '=SUM(1.5,3)',
        value: 4.5,
      });

    await page.evaluate(() => {
      const instance = (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst;
      instance.update({ locale: 'en' });
    });
    await expect.poll(() => readCell(0)).toEqual(expected);
    await expect.poll(() => sp.formulaBarValue()).toBe('=SUM(1.5,3)');
  }

  // YEN is the Japanese display alias of canonical DOLLAR; canonical DOLLAR
  // must not be interpreted a second time as the alias of USDOLLAR.
  await sp.mount({ locale: 'ja', platform });
  await sp.expectNoStub();
  await sp.typeIntoActiveCell('=YEN(1234.5)');
  const yen = { formula: '=DOLLAR(1234.5)', value: '¥1,235' };
  await expect.poll(() => readCell(0)).toEqual(yen);
  await page.keyboard.press('ArrowUp');
  await expect.poll(() => sp.formulaBarValue()).toBe('=YEN(1234.5)');
  await page.keyboard.press('F2');
  await expect(page.locator('.fc-host__editor')).toHaveValue('=YEN(1234.5)');
  await page.keyboard.press('Enter');
  await expect.poll(() => readCell(0)).toEqual(yen);
  expect(errors.read()).toEqual([]);
}
