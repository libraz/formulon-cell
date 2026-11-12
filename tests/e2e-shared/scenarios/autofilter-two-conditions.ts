import { expect, type Page } from '@playwright/test';

import { filterButtonPoint, hiddenRows, readAutoFilterRange } from '../pages/autofilter.js';
import { UserJourneyPage } from '../pages/UserJourneyPage.js';

/** The Mac column filter joins a second comparison with Or: less than 2 or greater than 8 over
 *  1, 4, 6 and 9 keeps 1 and 9, and the reopened filter shows the pair. */
export async function runMacAutoFilterOrPairScenario(page: Page): Promise<void> {
  const sheet = new UserJourneyPage(page);
  await sheet.mount({ locale: 'en', platform: 'mac' });
  await sheet.expectNoStub();
  await sheet.seed({ A1: 'Qty', A2: 1, A3: 4, A4: 6, A5: 9 });
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
  await dropdown.locator('select.fc-filter-dropdown__condition-op').selectOption('lessThan');
  await dropdown.locator('.fc-filter-dropdown__condition-value').fill('2');
  const join = dropdown.locator('.fc-filter-dropdown__condition-join');
  await expect(join).toBeVisible();
  await expect(join.locator('input[value="and"]')).toBeChecked();
  await join.locator('input[value="or"]').check();
  await dropdown.locator('select.fc-filter-dropdown__condition-op2').selectOption('greaterThan');
  const second = dropdown.locator('.fc-filter-dropdown__condition-second');
  await expect(second).toBeVisible();
  await second.fill('8');
  await second.press('Enter');
  await expect.poll(() => hiddenRows(page)).toEqual([2, 3]);
  await page.keyboard.press('Escape');
  await expect(dropdown).toBeHidden();

  dropdown = await open();
  await expect(
    dropdown.locator('.fc-filter-dropdown__condition-join input[value="or"]'),
  ).toBeChecked();
  await expect(dropdown.locator('select.fc-filter-dropdown__condition-op2')).toHaveValue(
    'greaterThan',
  );
  await expect(dropdown.locator('.fc-filter-dropdown__condition-second')).toHaveValue('8');
  await page.keyboard.press('Escape');
  await sheet.expectNoConsoleErrors();
}
