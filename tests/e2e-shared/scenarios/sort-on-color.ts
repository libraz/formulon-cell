import { expect, type Page } from '@playwright/test';

import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/api/spreadsheet/instance.js';
import { selectRibbonTab } from '../pages/ribbon-tabs.js';
import { UserJourneyPage } from '../pages/UserJourneyPage.js';

type InstanceWindow = Window & { __fcInst: SpreadsheetInstance };

/** Data > Sort with Sort On set to Cell Color: the level offers On Top and On Bottom and the
 *  column's fills with the first one pressed, and OK brings the filled row to the top. */
export async function runSortOnCellColorScenario(
  page: Page,
  platform: 'default' | 'mac',
): Promise<void> {
  const sheet = new UserJourneyPage(page);
  await sheet.mount({ locale: 'en', platform });
  await sheet.expectNoStub();
  await sheet.seed({ A1: 'Item', A2: 'Pen', A3: 'Ink', A4: 'Pad', A5: 'Cap' });
  await page.evaluate(() => {
    const inst = (window as unknown as InstanceWindow).__fcInst;
    inst.workbook.sheet(0).updateCellFormat('A4', { fill: { pattern: 1, foreground: '#FFFF00' } });
    inst.selection.set('A1:A5');
  });
  await selectRibbonTab(page, 'data');
  await sheet.clickRibbon(platform === 'mac' ? 'mac.data.sortCustom' : 'sortData');
  const dialog = page.getByRole('dialog').filter({ has: page.locator('.fc-sortdlg__level') });
  await expect(dialog).toBeVisible();

  const level = dialog.locator('.fc-sortdlg__level').first();
  const sortOn = level.locator('select.fc-sortdlg__sort-on');
  // The column has a fill but no icon set.
  await expect(sortOn.locator('option[value="cellColor"]')).not.toHaveAttribute('disabled');
  await expect(sortOn.locator('option[value="cellIcon"]')).toHaveAttribute('disabled', '');
  await sortOn.selectOption('cellColor');
  const direction = level.locator('select.fc-sortdlg__direction');
  await expect(direction.locator('option')).toHaveCount(2);
  await expect(direction).toHaveValue('asc');
  const swatch = level.locator('.fc-sortdlg__swatch[data-fc-color*="ffff00"]');
  await expect(swatch).toHaveAttribute('aria-pressed', 'true');
  await dialog.locator('.fc-fmtdlg__btn--primary').click();
  await expect(dialog).toBeHidden();

  const column = async () =>
    (await Promise.all(['A2', 'A3', 'A4', 'A5'].map((ref) => sheet.readValue(ref)))).map((cell) =>
      cell.kind === 'text' ? cell.value : null,
    );
  // On Top: Pad leads, the rest keep their order.
  await expect.poll(column).toEqual(['Pad', 'Pen', 'Ink', 'Cap']);
  await sheet.expectNoConsoleErrors();
}
