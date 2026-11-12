import { expect, type Page } from '@playwright/test';

import type { E2eProbes } from '../../../apps/demo-shared/e2e-probes.js';
import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/api/spreadsheet/instance.js';
import { gridCellCenter } from '../pages/SpreadsheetPage.js';
import { UserJourneyPage } from '../pages/UserJourneyPage.js';

type ProbeWindow = Window & { __fcInst: SpreadsheetInstance; __fcProbes: E2eProbes };

/** Each locale's way of typing 2026-05-15 and 2026-05-31. */
const RANGE: Record<'en' | 'de' | 'ja', readonly [string, string]> = {
  en: ['5/15/2026', '5/31/2026'],
  de: ['15.5.2026', '31.5.2026'],
  ja: ['2026/5/15', '2026/5/31'],
};

// PivotFilterKind values the panel's options carry.
const DATE_BETWEEN = '5';
const MAY = '60';

/** Seeds Date and Sales rows and builds a report at E1 with Date on rows and Sum of Sales. */
const seedDateReport = (page: Page, mac: boolean) =>
  page.evaluate((mac) => {
    const w = window as unknown as ProbeWindow;
    const inst = w.__fcInst;
    const s = inst.workbook.sheet(0);
    s.setText(0, 0, 'Date');
    s.setText(0, 1, 'Sales');
    // 2026-05-08, 05-18, 05-28 and 06-07.
    [46150, 46160, 46170, 46180].forEach((serial, i) => {
      s.setNumber(i + 1, 0, serial);
      s.setNumber(i + 1, 1, (i + 1) * 10);
    });
    s.updateCellFormat('A2:A5', { numberFormat: 'yyyy-mm-dd' });
    const handle = w.__fcProbes['api/workbook'].workbookHandleOf(inst.workbook);
    const made = w.__fcProbes['commands/pivot-table'].createPivotTableFromRange(handle, {
      source: { sheet: 0, r0: 0, c0: 0, r1: 4, c1: 1 },
      destination: { sheet: 0, row: 0, col: 4 },
      name: 'DateReport',
      rowField: 'Date',
      valueField: 'Sales',
      aggregation: 0,
    });
    if (!made.ok) throw new Error(`pivot creation failed: ${made.reason}`);
    inst.history.clear();
    const t = inst.i18n.strings as unknown as {
      pivotTableDialog: { filterDialogs: { date: { title: string; macTitle: string } } };
    };
    // The Mac names the dialog alone; elsewhere it carries the field.
    const { title, macTitle } = t.pivotTableDialog.filterDialogs.date;
    return mac ? macTitle : title.replace('{field}', 'Date');
  }, mac);

/** The report's item rows as [date serial, Sum of Sales], then the grand total. */
const reportRows = (page: Page) =>
  page.evaluate(() => {
    const s = (window as unknown as ProbeWindow).__fcInst.workbook.sheet(0);
    const rows: [unknown, unknown][] = [];
    for (let r = 1; r < 8; r++) {
      const value = s.getValue(r, 5);
      if (typeof value === 'number') rows.push([s.getValue(r, 4), value]);
    }
    return rows;
  });

/** Right-clicks the first date label and opens Filter > Date Filters... */
const openDateFilters = async (page: Page) => {
  const label = await gridCellCenter(page, 1, 4);
  await page.mouse.click(label.x, label.y, { button: 'right' });
  await page.locator('.fc-ctxmenu [data-fc-action="filter"]').first().hover();
  await page.locator('.fc-ctxmenu__sub [data-fc-action="pivotFilterLabel"]').click();
  const panel = page.locator('.fc-pivot-filter-panel');
  await expect(panel).toBeVisible();
  return panel.locator('select[data-pivot-filter-by="label"]');
};

/** A date row field's Date Filters: Between opens the Date Filter dialog, which reads both bounds
 *  with the locale's date grammar; a period such as May applies at once. */
export async function runPivotDateBetweenScenario(
  page: Page,
  locale: 'en' | 'de' | 'ja',
  platform: 'default' | 'mac',
): Promise<void> {
  const sheet = new UserJourneyPage(page);
  await sheet.mount({ locale, platform });
  await sheet.expectNoStub();
  const dialogTitle = await seedDateReport(page, platform === 'mac');
  await expect.poll(() => reportRows(page)).toHaveLength(5);

  await (await openDateFilters(page)).selectOption(DATE_BETWEEN);
  const filter = page.getByRole('dialog', { name: dialogTitle });
  await expect(filter).toBeVisible();
  await expect(filter.locator('[data-pivot-filter-operator="true"]')).toHaveValue(DATE_BETWEEN);
  const [low, high] = RANGE[locale];
  await filter.locator('[data-pivot-filter-value="true"]').fill(low);
  await filter.locator('[data-pivot-filter-value-high="true"]').fill(high);
  await filter.locator('.fc-fmtdlg__btn--primary').click();
  await expect(filter).toBeHidden();
  await expect
    .poll(() => reportRows(page))
    .toEqual([
      [46160, 20],
      [46170, 30],
      [expect.any(String), 50],
    ]);

  await page.keyboard.press('Escape');
  await (await openDateFilters(page)).selectOption(MAY);
  await expect(filter).toBeHidden();
  await expect
    .poll(() => reportRows(page))
    .toEqual([
      [46150, 10],
      [46160, 20],
      [46170, 30],
      [expect.any(String), 60],
    ]);
  await sheet.expectNoConsoleErrors();
}
