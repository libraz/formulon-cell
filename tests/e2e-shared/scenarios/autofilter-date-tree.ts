import { expect, type Page } from '@playwright/test';

import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/api/spreadsheet/instance.js';
import { filterButtonPoint, hiddenRows, readAutoFilterRange } from '../pages/autofilter.js';
import { UserJourneyPage } from '../pages/UserJourneyPage.js';

type InstanceWindow = Window & { __fcInst: SpreadsheetInstance };

/** A date column's filter checklist groups by year, month and day: unchecking May under 2026
 *  hides the May rows, and the reopened list shows May unchecked under a mixed 2026. */
export async function runAutoFilterDateTreeScenario(
  page: Page,
  platform: 'default' | 'mac',
): Promise<void> {
  const sheet = new UserJourneyPage(page);
  await sheet.mount({ locale: 'en', platform });
  await sheet.expectNoStub();
  await page.evaluate(() => {
    const s = (window as unknown as InstanceWindow).__fcInst.workbook.sheet(0);
    s.setText(0, 0, 'Date');
    // 2026-05-08, 05-18, 05-28 and 06-07.
    [46150, 46160, 46170, 46180].forEach((serial, i) => {
      s.setNumber(i + 1, 0, serial);
    });
    s.updateCellFormat('A2:A5', { numberFormat: 'yyyy-mm-dd' });
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
    // Groups start collapsed: only the year shows until it is expanded.
    const year = dropdown.locator('.fc-filter-dropdown__disclosure[data-fc-date-path="2026"]');
    await expect(year).toHaveAttribute('aria-expanded', 'false');
    await expect(dropdown.locator('input[value="group:2026-5"]')).toBeHidden();
    await year.click();
    await expect(year).toHaveAttribute('aria-expanded', 'true');
    return dropdown;
  };

  let dropdown = await open();
  const may = dropdown.locator('input[value="group:2026-5"]');
  await expect(may).toBeChecked();
  await may.uncheck();
  // The Mac applies each change; elsewhere Apply commits the list.
  if (platform === 'default') await dropdown.locator('.fc-filter-dropdown__apply').click();
  await expect.poll(() => hiddenRows(page)).toEqual([1, 2, 3]);
  if (await dropdown.isVisible()) await page.keyboard.press('Escape');
  await expect(dropdown).toBeHidden();

  dropdown = await open();
  await expect(dropdown.locator('input[value="group:2026-5"]')).not.toBeChecked();
  await expect(dropdown.locator('input[value="group:2026-6"]')).toBeChecked();
  expect(
    await dropdown
      .locator('input[value="group:2026"]')
      .evaluate((el) => (el as HTMLInputElement).indeterminate),
  ).toBe(true);
  await page.keyboard.press('Escape');
  await sheet.expectNoConsoleErrors();
}
