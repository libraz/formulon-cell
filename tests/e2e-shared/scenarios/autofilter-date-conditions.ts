import { expect, type Page } from '@playwright/test';

import type { E2eProbes } from '../../../apps/demo-shared/e2e-probes.js';
import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/api/spreadsheet/instance.js';
import { filterButtonPoint, hiddenRows, readAutoFilterRange } from '../pages/autofilter.js';
import { UserJourneyPage } from '../pages/UserJourneyPage.js';

type ProbeWindow = Window & { __fcInst: SpreadsheetInstance; __fcProbes: E2eProbes };

/** Each locale's way of typing 2026-05-20. */
const BEFORE: Record<'en' | 'de' | 'ja', string> = {
  en: '5/20/2026',
  de: '20.5.2026',
  ja: '2026/5/20',
};

const openDateFilter = async (page: Page) => {
  const point = await filterButtonPoint(page, 0, 0);
  await page.mouse.click(point.x, point.y);
  const dropdown = page.locator('.fc-filter-dropdown:visible').first();
  await expect(dropdown).toBeVisible();
  await expect(dropdown).toHaveAttribute('data-fc-date-column', 'true');
  return dropdown;
};

/** The Mac column filter on a date column: Before reads the bound with the locale's date grammar,
 *  and All Dates in the Period hides the value field and applies the month its pop-up picks. */
export async function runMacAutoFilterDateScenario(
  page: Page,
  locale: 'en' | 'de' | 'ja',
): Promise<void> {
  const sheet = new UserJourneyPage(page);
  await sheet.mount({ locale, platform: 'mac' });
  await sheet.expectNoStub();
  await page.evaluate(() => {
    const s = (window as unknown as ProbeWindow).__fcInst.workbook.sheet(0);
    s.setText(0, 0, 'Date');
    s.setText(0, 1, 'Sales');
    // 2026-05-08, 05-18, 05-28 and 06-07.
    [46150, 46160, 46170, 46180].forEach((serial, i) => {
      s.setNumber(i + 1, 0, serial);
      s.setNumber(i + 1, 1, (i + 1) * 10);
    });
    s.updateCellFormat('A2:A5', { numberFormat: 'yyyy-mm-dd' });
  });
  await sheet.goTo('A2');
  await sheet.shortcut('Shift+L');
  await expect
    .poll(() => readAutoFilterRange(page))
    .toEqual({ firstRow: 0, firstCol: 0, lastRow: 4, lastCol: 1 });

  let dropdown = await openDateFilter(page);
  const op = dropdown.locator('select.fc-filter-dropdown__condition-op');
  await op.selectOption('lessThan');
  const value = dropdown.locator('.fc-filter-dropdown__condition-value');
  await value.fill(BEFORE[locale]);
  await value.press('Enter');
  // Rows dated 05-28 and 06-07 are filtered out.
  await expect.poll(() => hiddenRows(page)).toEqual([3, 4]);
  await page.keyboard.press('Escape');
  await expect(dropdown).toBeHidden();

  dropdown = await openDateFilter(page);
  // Months sit behind All Dates in the Period, whose own pop-up picks one.
  await dropdown
    .locator('select.fc-filter-dropdown__condition-op')
    .selectOption('allDatesInPeriod');
  await expect(dropdown.locator('.fc-filter-dropdown__condition-value')).toBeHidden();
  await dropdown.locator('select.fc-filter-dropdown__condition-period').selectOption('june');
  await expect.poll(() => hiddenRows(page)).toEqual([1, 2, 3]);
  await page.keyboard.press('Escape');
  await sheet.expectNoConsoleErrors();
}
