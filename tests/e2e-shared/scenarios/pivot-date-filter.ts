import { expect, type Page } from '@playwright/test';

import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/api/spreadsheet/instance.js';
import { UserJourneyPage } from '../pages/UserJourneyPage.js';

type InstanceWindow = Window & { __fcInst: SpreadsheetInstance };

type PivotStrings = {
  title: string;
  fieldSettingsFor: string;
  filterDialog: string;
  filterDialogs: { date: { title: string } };
};

/** Each locale's way of typing 2026-05-15 and 2026-05-31. */
const RANGE: Record<'en' | 'de' | 'ja', readonly [string, string]> = {
  en: ['5/15/2026', '5/31/2026'],
  de: ['15.5.2026', '31.5.2026'],
  ja: ['2026/5/15', '2026/5/31'],
};

/** Date is between, set from Create PivotTable > Field Settings > Filter...: the Date Filter
 *  dialog reads both bounds with the locale's date grammar, and the report keeps only the rows
 *  dated inside them. */
export async function runPivotDateBetweenScenario(
  page: Page,
  locale: 'en' | 'de' | 'ja',
): Promise<void> {
  const sheet = new UserJourneyPage(page);
  await sheet.mount({ locale });
  await sheet.expectNoStub();
  const t = await page.evaluate(() => {
    const inst = (window as unknown as InstanceWindow).__fcInst;
    const s = inst.workbook.sheet(0);
    ['Region', 'Date', 'Sales'].forEach((label, col) => {
      s.setText(0, col, label);
    });
    // 2026-05-08, 05-18, 05-28 and 06-07.
    [46150, 46160, 46170, 46180].forEach((serial, i) => {
      s.setText(i + 1, 0, i % 2 ? 'West' : 'East');
      s.setNumber(i + 1, 1, serial);
      s.setNumber(i + 1, 2, (i + 1) * 10);
    });
    s.updateCellFormat('B2:B5', { numberFormat: 'yyyy-mm-dd' });
    inst.history.clear();
    inst.selection.set('A1:C5');
    inst.dialogs.open('pivotTable');
    return (inst.i18n.strings as unknown as { pivotTableDialog: PivotStrings }).pivotTableDialog;
  });

  const create = page.getByRole('dialog', { name: t.title });
  await expect(create).toBeVisible();
  for (const field of ['Region', 'Date', 'Sales'])
    await create.locator(`[data-pivot-field-list-field="${field}"]`).check();
  await create
    .locator('[data-pivot-area-field="Date"]')
    .dragTo(create.locator('[data-pivot-area="filters"]'));
  await expect(
    create.locator('[data-pivot-area="filters"] [data-pivot-area-field="Date"]'),
  ).toBeVisible();
  await create
    .getByRole('button', { name: t.fieldSettingsFor.replace('{field}', 'Date'), exact: true })
    .click();
  await create.locator('[data-pivot-filter-category="true"]').selectOption('date');
  // 5 is PivotFilterKind.DateBetween.
  await create.locator('[data-pivot-filter-condition="true"]').selectOption('5');
  await create.getByRole('button', { name: t.filterDialog, exact: true }).click();

  const filter = page.getByRole('dialog', {
    name: t.filterDialogs.date.title.replace('{field}', 'Date'),
  });
  await expect(filter).toBeVisible();
  await expect(filter.locator('[data-pivot-filter-operator="true"]')).toHaveValue('5');
  const [low, high] = RANGE[locale];
  await filter.locator('[data-pivot-filter-value="true"]').fill(low);
  await filter.locator('[data-pivot-filter-value-high="true"]').fill(high);
  await filter.locator('.fc-fmtdlg__btn--primary').click();
  await expect(filter).toBeHidden();
  await create.locator('.fc-fmtdlg__btn--primary').click();
  await expect(create).toBeHidden();

  // The report lands on a new sheet: East keeps 05-28 (30), West keeps 05-18 (20).
  const report = () =>
    page.evaluate(() => {
      const wb = (window as unknown as InstanceWindow).__fcInst.workbook;
      for (let i = 0; i < wb.listSheets().length; i++) {
        const s = wb.sheet(i);
        if (s.getValue(0, 0) === 'Region') continue;
        const rows: [unknown, unknown][] = [];
        for (let r = 0; r < 12; r++) rows.push([s.getValue(r, 0), s.getValue(r, 1)]);
        return rows.filter(([label, value]) => typeof value === 'number' && label !== null);
      }
      return null;
    });
  await expect.poll(report).toEqual([
    ['East', 30],
    ['West', 20],
    [expect.any(String), 50],
  ]);
  await sheet.expectNoConsoleErrors();
}
