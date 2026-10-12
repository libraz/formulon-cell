import { expect, type Page } from '@playwright/test';
import { SpreadsheetPage } from '../pages/SpreadsheetPage.js';

/** Exercise real editor and both clipboard payloads without the system clipboard. */
export async function runLocaleInputScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  const errors = sp.collectConsoleErrors();
  const readCell = (row: number) =>
    page.evaluate((r) => {
      const instance = (
        window as unknown as {
          __fcInst: {
            workbook: {
              getValue: (addr: { sheet: number; row: number; col: number }) => unknown;
            };
          };
        }
      ).__fcInst;
      return instance.workbook.getValue({ sheet: 0, row: r, col: 0 });
    }, row);

  for (const profile of [
    { locale: 'de', trueText: 'WAHR' },
    { locale: 'fr', trueText: 'VRAI' },
  ]) {
    await sp.mount({ locale: profile.locale });
    await sp.expectNoStub();
    await sp.typeIntoActiveCell('1,5');
    await page.keyboard.press('ArrowUp');
    await expect.poll(() => sp.formulaBarValue()).toBe('1,5');
    await page.keyboard.press('F2');
    await expect(page.locator('.fc-host__editor')).toHaveValue('1,5');
    await page.keyboard.press('Enter');
    await expect.poll(() => readCell(0)).toEqual({ kind: 'number', value: 1.5 });
    await page.keyboard.press('ArrowUp');

    const payload = await page.evaluate(() => {
      const host = document.querySelector('.fc-host');
      if (!host) throw new Error('Spreadsheet host missing');
      const data = new DataTransfer();
      host.dispatchEvent(
        new ClipboardEvent('copy', { clipboardData: data, bubbles: true, cancelable: true }),
      );
      const html = data.getData('text/html');
      const parsed = new DOMParser().parseFromString(html, 'text/html');
      return {
        text: data.getData('text/plain'),
        htmlText: parsed.querySelector('td')?.textContent,
      };
    });
    expect(payload).toEqual({ text: '1,5', htmlText: '1,5' });
    await page.keyboard.press('ArrowDown');
    await page.keyboard.press('ArrowDown');
    await page.evaluate((text) => {
      const data = new DataTransfer();
      data.setData('text/plain', text);
      document
        .querySelector('.fc-host')
        ?.dispatchEvent(
          new ClipboardEvent('paste', { clipboardData: data, bubbles: true, cancelable: true }),
        );
    }, payload.text);
    await expect.poll(() => readCell(2)).toEqual({ kind: 'number', value: 1.5 });

    await sp.typeIntoActiveCell(profile.trueText);
    await page.keyboard.press('ArrowUp');
    await expect.poll(() => readCell(0)).toEqual({ kind: 'bool', value: true });
    await expect.poll(() => sp.formulaBarValue()).toBe(profile.trueText);
    await page.keyboard.press('F2');
    await expect(page.locator('.fc-host__editor')).toHaveValue(profile.trueText);
    await page.keyboard.press('Enter');
    await expect.poll(() => readCell(0)).toEqual({ kind: 'bool', value: true });
  }

  expect(errors.read()).toEqual([]);
}
