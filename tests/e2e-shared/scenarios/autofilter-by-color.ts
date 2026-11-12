import { expect, type Locator, type Page } from '@playwright/test';

import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/api/spreadsheet/instance.js';
import { filterButtonPoint, hiddenRows, readAutoFilterRange } from '../pages/autofilter.js';
import { UserJourneyPage } from '../pages/UserJourneyPage.js';

type InstanceWindow = Window & { __fcInst: SpreadsheetInstance };

// The panel's two By color pop-ups sit on either side of the rule between Sort and Filter.
const sortSide = (dropdown: Locator) =>
  dropdown.locator('.fc-filter-dropdown__color-filter:not(.fc-filter-dropdown__rule ~ *)');
const filterSide = (dropdown: Locator) =>
  dropdown.locator('.fc-filter-dropdown__rule ~ .fc-filter-dropdown__color-filter');
const yellowSwatch = (side: Locator) =>
  side.locator('.fc-filter-dropdown__swatch[data-fc-color*="ffff00"]');

/** Seeds Pen, Ink, Pad and Cap under Item with Ink and Cap filled yellow, and turns AutoFilter on. */
const seedFilledColumn = async (page: Page, sheet: UserJourneyPage) => {
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
};

const openColumnFilter = async (page: Page) => {
  const point = await filterButtonPoint(page, 0, 0);
  await page.mouse.click(point.x, point.y);
  const dropdown = page.locator('.fc-filter-dropdown:visible').first();
  await expect(dropdown).toBeVisible();
  return dropdown;
};

/** The panel reopens itself after a swatch applies; open it again only if it closed. */
const reopenedColumnFilter = async (page: Page) => {
  const dropdown = page.locator('.fc-filter-dropdown:visible').first();
  return (await dropdown.isVisible()) ? dropdown : openColumnFilter(page);
};

/** The Mac column filter's By color: Cell Color offers the column's fills, a swatch keeps only the
 *  rows filled with it, the reopened panel shows that swatch pressed, and no choice clears it. */
export async function runMacAutoFilterByColorScenario(page: Page): Promise<void> {
  const sheet = new UserJourneyPage(page);
  await sheet.mount({ locale: 'en', platform: 'mac' });
  await sheet.expectNoStub();
  await seedFilledColumn(page, sheet);

  let dropdown = await openColumnFilter(page);
  const kind = filterSide(dropdown).locator('select.fc-filter-dropdown__color-kind');
  // The column has fills but no icon set, so Cell Icon is greyed.
  await expect(kind.locator('option[value="cellColor"]')).not.toHaveAttribute('disabled');
  await expect(kind.locator('option[value="cellIcon"]')).toHaveAttribute('disabled', '');
  await kind.selectOption('cellColor');
  await yellowSwatch(filterSide(dropdown)).click();
  // Pen and Pad have no fill and drop out.
  await expect.poll(() => hiddenRows(page)).toEqual([1, 3]);

  dropdown = await reopenedColumnFilter(page);
  await expect(yellowSwatch(filterSide(dropdown))).toHaveAttribute('aria-pressed', 'true');
  await filterSide(dropdown).locator('select.fc-filter-dropdown__color-kind').selectOption('');
  await expect.poll(() => hiddenRows(page)).toEqual([]);
  await page.keyboard.press('Escape');
  await sheet.expectNoConsoleErrors();
}

/** The Mac panel's sort-side By color: Cell Color and the yellow swatch bring the filled rows to
 *  the top, and the reopened panel shows that swatch pressed. */
export async function runMacSortByColorScenario(page: Page): Promise<void> {
  const sheet = new UserJourneyPage(page);
  await sheet.mount({ locale: 'en', platform: 'mac' });
  await sheet.expectNoStub();
  await seedFilledColumn(page, sheet);

  let dropdown = await openColumnFilter(page);
  await sortSide(dropdown)
    .locator('select.fc-filter-dropdown__color-kind')
    .selectOption('cellColor');
  await yellowSwatch(sortSide(dropdown)).click();
  const column = async () =>
    (await Promise.all(['A2', 'A3', 'A4', 'A5'].map((ref) => sheet.readValue(ref)))).map((cell) =>
      cell.kind === 'text' ? cell.value : null,
    );
  // Filled rows first, the rest after them in their old order.
  await expect.poll(column).toEqual(['Ink', 'Cap', 'Pen', 'Pad']);

  dropdown = await reopenedColumnFilter(page);
  await expect(yellowSwatch(sortSide(dropdown))).toHaveAttribute('aria-pressed', 'true');
  await page.keyboard.press('Escape');
  await sheet.expectNoConsoleErrors();
}
