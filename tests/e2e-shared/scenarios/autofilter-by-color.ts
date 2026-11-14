import { expect, type Locator, type Page } from '@playwright/test';

import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/api/spreadsheet/instance.js';
import { filterButtonPoint, hiddenRows, readAutoFilterRange } from '../pages/autofilter.js';
import { UserJourneyPage } from '../pages/UserJourneyPage.js';

type InstanceWindow = Window & { __fcInst: SpreadsheetInstance };

// The panel has a By color pop-up on its sort side and on its filter side.
const sortSide = (dropdown: Locator) => dropdown.locator('[data-fc-color-side="sort"]');
const filterSide = (dropdown: Locator) => dropdown.locator('[data-fc-color-side="filter"]');
const yellowSwatch = (side: Locator) =>
  side.locator('.fc-filter-dropdown__swatch[data-fc-color*="ffff00"]');
const swatches = (side: Locator) => side.locator('.fc-filter-dropdown__swatches > *');

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

/** The Mac column filter's By color: every kind is offered, Cell Color lists the column's fills
 *  then No Fill, a swatch keeps only the rows filled with it, the reopened panel shows that swatch
 *  pressed, no choice clears it, and No Fill keeps only the unfilled rows. */
export async function runMacAutoFilterByColorScenario(page: Page): Promise<void> {
  const sheet = new UserJourneyPage(page);
  await sheet.mount({ locale: 'en', platform: 'mac' });
  await sheet.expectNoStub();
  await seedFilledColumn(page, sheet);

  let dropdown = await openColumnFilter(page);
  const kind = filterSide(dropdown).locator('select.fc-filter-dropdown__color-kind');
  // Every kind is offered, even Cell Icon on a column with no icon set.
  for (const value of ['cellColor', 'fontColor', 'cellIcon'])
    await expect(kind.locator(`option[value="${value}"]`)).not.toHaveAttribute('disabled');
  await kind.selectOption('cellColor');
  // The used fill, then No Fill because Pen and Pad have none.
  await expect(swatches(filterSide(dropdown))).toHaveCount(2);
  await expect(swatches(filterSide(dropdown)).first()).toHaveAttribute('data-fc-color', /ffff00/i);
  const noFill = swatches(filterSide(dropdown)).last();
  await expect(noFill).toHaveClass(/fc-filter-dropdown__swatch--none/);
  await expect(noFill).toHaveText('No Fill');
  await yellowSwatch(filterSide(dropdown)).click();
  // Pen and Pad have no fill and drop out.
  await expect.poll(() => hiddenRows(page)).toEqual([1, 3]);

  dropdown = await reopenedColumnFilter(page);
  await expect(yellowSwatch(filterSide(dropdown))).toHaveAttribute('aria-pressed', 'true');
  await filterSide(dropdown).locator('select.fc-filter-dropdown__color-kind').selectOption('');
  await expect.poll(() => hiddenRows(page)).toEqual([]);

  dropdown = await reopenedColumnFilter(page);
  await filterSide(dropdown)
    .locator('select.fc-filter-dropdown__color-kind')
    .selectOption('cellColor');
  await filterSide(dropdown).locator('.fc-filter-dropdown__swatch--none').click();
  // Ink and Cap are filled and drop out.
  await expect.poll(() => hiddenRows(page)).toEqual([2, 4]);
  await page.keyboard.press('Escape');
  await sheet.expectNoConsoleErrors();
}

/** The Mac panel's sort-side By color: Cell Color and the yellow swatch bring the filled rows to
 *  the top, the reopened panel shows that swatch pressed, and Custom Sort..., last in the list,
 *  opens the Sort dialog. */
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
  const customSort = swatches(sortSide(dropdown)).last();
  await expect(customSort).toHaveClass(/fc-filter-dropdown__custom-sort/);
  await customSort.click();
  await expect(dropdown).toBeHidden();
  const sortDialog = page.getByRole('dialog').filter({ has: page.locator('.fc-sortdlg__level') });
  await expect(sortDialog).toBeVisible();
  await page.keyboard.press('Escape');
  await expect(sortDialog).toBeHidden();
  await sheet.expectNoConsoleErrors();
}
