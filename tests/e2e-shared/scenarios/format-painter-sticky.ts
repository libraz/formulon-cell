import { expect, type Page } from '@playwright/test';

import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/api/spreadsheet/instance.js';
import { selectRibbonTab } from '../pages/ribbon-tabs.js';
import { gridCellCenter } from '../pages/SpreadsheetPage.js';
import { UserJourneyPage } from '../pages/UserJourneyPage.js';

type InstanceWindow = Window & { __fcInst: SpreadsheetInstance };

/** Format Painter from the Home ribbon button: a click paints once, a double-click keeps painting
 *  until the button is clicked again or Escape is pressed. */
export async function runStickyFormatPainterScenario(
  page: Page,
  platform: 'default' | 'mac',
): Promise<void> {
  const sheet = new UserJourneyPage(page);
  await sheet.mount({ locale: 'en', platform });
  await sheet.expectNoStub();
  await sheet.seed({ A1: 'source', B2: 'x', C3: 'x', B4: 'x', B5: 'x', B6: 'x', B7: 'x' });
  await page.evaluate(() =>
    (window as unknown as InstanceWindow).__fcInst.workbook
      .sheet(0)
      .updateCellFormat('A1', { font: { bold: true } }),
  );
  await selectRibbonTab(page, 'home');
  const button = page.locator('[data-ribbon-command="formatPainter"]').first();
  const host = page.locator('.fc-host').first();
  const painting = async (on: boolean) =>
    on
      ? await expect(host).toHaveClass(/fc-host--paintbrush/)
      : await expect(host).not.toHaveClass(/fc-host--paintbrush/);
  const clickCell = async (row: number, col: number) => {
    const point = await gridCellCenter(page, row, col);
    await page.mouse.click(point.x, point.y);
  };
  const bold = async (ref: string) => (await sheet.readFormat(ref))?.bold === true;

  // One click paints the next target only.
  await sheet.goTo('A1');
  await button.click();
  await painting(true);
  await clickCell(1, 1);
  await expect.poll(() => bold('B2')).toBe(true);
  await painting(false);
  await clickCell(2, 2);
  expect(await bold('C3')).toBe(false);

  // A double-click keeps painting until the button is clicked again.
  await sheet.goTo('A1');
  await button.dblclick();
  await painting(true);
  await clickCell(3, 1);
  await expect.poll(() => bold('B4')).toBe(true);
  await painting(true);
  await clickCell(4, 1);
  await expect.poll(() => bold('B5')).toBe(true);
  await painting(true);
  await button.click();
  await painting(false);
  await clickCell(5, 1);
  expect(await bold('B6')).toBe(false);

  // Escape also ends a sticky painter.
  await sheet.goTo('A1');
  await button.dblclick();
  await painting(true);
  await page.keyboard.press('Escape');
  await painting(false);
  await clickCell(6, 1);
  expect(await bold('B7')).toBe(false);
  await sheet.expectNoConsoleErrors();
}
