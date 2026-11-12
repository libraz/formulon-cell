import { expect, type Page } from '@playwright/test';

import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/api/spreadsheet/instance.js';
import { filterButtonPoint, hiddenRows, readAutoFilterRange } from '../pages/autofilter.js';
import { UserJourneyPage } from '../pages/UserJourneyPage.js';

type InstanceWindow = Window & { __fcInst: SpreadsheetInstance };

/** The Mac column filter's By color: Cell Color offers the column's fills, a swatch keeps only the
 *  rows filled with it, the reopened panel shows that swatch pressed, and no choice clears it. */
export async function runMacAutoFilterByColorScenario(page: Page): Promise<void> {
  const sheet = new UserJourneyPage(page);
  await sheet.mount({ locale: 'en', platform: 'mac' });
  await sheet.expectNoStub();
  await sheet.seed({ A1: 'Item', A2: 'Pen', A3: 'Ink', A4: 'Pad', A5: 'Cap' });
  await page.evaluate(() => {
    const s = (window as unknown as InstanceWindow).__fcInst.workbook.sheet(0);
    for (const ref of ['A3', 'A5'])
      s.updateCellFormat(ref, { fill: { pattern: 1, foreground: '#FFFF00' } });
  });
  await sheet.goTo('A2');
  await sheet.shortcut('Shift+L');
  await expect
    .poll(() => readAutoFilterRange(page))
    .toEqual({ firstRow: 0, firstCol: 0, lastRow: 4, lastCol: 0 });

  const open = async () => {
    const point = await filterButtonPoint(page, 0, 0);
    await page.mouse.click(point.x, point.y);
    const dropdown = page.locator('.fc-filter-dropdown:visible').first();
    await expect(dropdown).toBeVisible();
    return dropdown;
  };
  let dropdown = await open();
  const kind = dropdown.locator('select.fc-filter-dropdown__color-kind');
  // The column has fills but no icon set, so Cell Icon is greyed.
  await expect(kind.locator('option[value="cellColor"]')).not.toHaveAttribute('disabled');
  await expect(kind.locator('option[value="cellIcon"]')).toHaveAttribute('disabled', '');
  await kind.selectOption('cellColor');
  const yellow = dropdown.locator('.fc-filter-dropdown__swatch[data-fc-color*="ffff00"]');
  await expect(yellow).toBeVisible();
  await yellow.click();
  // Pen and Pad have no fill and drop out.
  await expect.poll(() => hiddenRows(page)).toEqual([1, 3]);

  dropdown = page.locator('.fc-filter-dropdown:visible').first();
  if (!(await dropdown.isVisible())) dropdown = await open();
  await expect(
    dropdown.locator('.fc-filter-dropdown__swatch[data-fc-color*="ffff00"]'),
  ).toHaveAttribute('aria-pressed', 'true');
  await dropdown.locator('select.fc-filter-dropdown__color-kind').selectOption('');
  await expect.poll(() => hiddenRows(page)).toEqual([]);
  await page.keyboard.press('Escape');
  await sheet.expectNoConsoleErrors();
}
