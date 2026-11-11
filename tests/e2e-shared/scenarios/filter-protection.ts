import { expect, type Page } from '@playwright/test';

import type { E2eProbes } from '../../../apps/demo-shared/e2e-probes.js';
import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/api/spreadsheet/instance.js';
import { selectRibbonTab } from '../pages/ribbon-tabs.js';
import { UserJourneyPage } from '../pages/UserJourneyPage.js';

type InstanceWindow = Window & { __fcInst: SpreadsheetInstance };
type ProbeWindow = InstanceWindow & { __fcProbes: E2eProbes };

const readAutoFilterRange = (page: Page) =>
  page.evaluate(
    () =>
      (window as unknown as InstanceWindow).__fcInst.workbook.sheet(0).getAutoFilter()?.range ??
      null,
  );

const seedRegion = (sheet: UserJourneyPage) =>
  sheet.seed({ A1: 'Item', B1: 'Qty', A2: 'Pen', B2: 3, A3: 'Ink', B3: 5, A4: 'Pad', B4: 2 });

/** Ctrl+Shift+L turns AutoFilter on over the region around the active cell and off again, as the
 *  ribbon's Filter toggle does. */
export async function runFilterShortcutScenario(page: Page): Promise<void> {
  const sheet = new UserJourneyPage(page);
  await sheet.mount({ locale: 'en' });
  await sheet.expectNoStub();
  await seedRegion(sheet);
  await sheet.goTo('B3');

  await sheet.shortcut('Shift+L');
  await expect
    .poll(() => readAutoFilterRange(page))
    .toEqual({ firstRow: 0, firstCol: 0, lastRow: 3, lastCol: 1 });

  await sheet.shortcut('Shift+L');
  await expect.poll(() => readAutoFilterRange(page)).toBeNull();
  await sheet.expectNoConsoleErrors();
}

/** On a protected sheet the Page Layout Background command is unavailable and Ctrl+Shift+L does not
 *  turn AutoFilter on; both come back once the sheet is unprotected. */
export async function runProtectedSheetFilterAndBackgroundScenario(page: Page): Promise<void> {
  const sheet = new UserJourneyPage(page);
  await sheet.mount({ locale: 'en' });
  await sheet.expectNoStub();
  await seedRegion(sheet);
  await page.evaluate(() =>
    (window as unknown as InstanceWindow).__fcInst.workbook.sheet(0).protect(),
  );

  await selectRibbonTab(page, 'pageLayout');
  const background = page.locator('[data-ribbon-command="sheetBackground"]').first();
  await expect(background).toBeDisabled();

  await sheet.goTo('B3');
  await sheet.shortcut('Shift+L');
  await expect.poll(() => readAutoFilterRange(page)).toBeNull();

  await page.evaluate(() =>
    (window as unknown as InstanceWindow).__fcInst.workbook.sheet(0).unprotect(),
  );
  await expect(background).toBeEnabled();
  await sheet.goTo('B3');
  await sheet.shortcut('Shift+L');
  await expect
    .poll(() => readAutoFilterRange(page))
    .toEqual({ firstRow: 0, firstCol: 0, lastRow: 3, lastCol: 1 });
  await sheet.expectNoConsoleErrors();
}

/** Page point of the AutoFilter button in the header cell at (`row`, `col`), left-to-right sheet. */
const filterButtonPoint = async (page: Page, row: number, col: number) => {
  const box = await page.locator('.fc-host__grid').first().boundingBox();
  if (!box) throw new Error('Missing grid');
  const point = await page.evaluate(
    ({ row, col }) => {
      const w = window as unknown as ProbeWindow;
      const state = w.__fcProbes['mount/internals']
        .spreadsheetInternals(w.__fcInst)
        .store.getState();
      const { cellRectUnclamped, layoutForView } = w.__fcProbes['render/geometry'];
      const rect = cellRectUnclamped(layoutForView(state), state.viewport, row, col);
      // The 17 px button sits one px inside the cell's trailing edge and one below its top.
      return { x: rect.x + rect.w - 1 - 17 / 2, y: rect.y + 1 + 17 / 2 };
    },
    { row, col },
  );
  return { x: box.x + point.x, y: box.y + point.y };
};

const readActive = (page: Page) =>
  page.evaluate(() => (window as unknown as InstanceWindow).__fcInst.selection.get().active);

/** Clicks the B1 filter button and asserts nothing opens, nothing prompts and nothing moves. */
const expectFilterButtonInert = async (page: Page) => {
  const active = await readActive(page);
  const point = await filterButtonPoint(page, 0, 1);
  await page.mouse.click(point.x, point.y);
  // A refused click changes nothing; give it time to act before reading.
  await page.waitForTimeout(300);
  await expect(page.locator('.fc-filter-dropdown:visible')).toHaveCount(0);
  await expect(page.locator('.fc-valdlg:visible')).toHaveCount(0);
  expect(await readActive(page)).toEqual(active);
};

/** On a protected sheet an existing AutoFilter button is inert without Use AutoFilter: the click
 *  opens no dropdown and raises no prompt. With Use AutoFilter it opens the dropdown, while
 *  Ctrl+Shift+L still cannot turn the AutoFilter off. */
export async function runProtectedSheetFilterButtonScenario(page: Page): Promise<void> {
  const sheet = new UserJourneyPage(page);
  await sheet.mount({ locale: 'en' });
  await sheet.expectNoStub();
  await seedRegion(sheet);
  await sheet.goTo('B3');
  await sheet.shortcut('Shift+L');
  const range = { firstRow: 0, firstCol: 0, lastRow: 3, lastCol: 1 };
  await expect.poll(() => readAutoFilterRange(page)).toEqual(range);

  await page.evaluate(() =>
    (window as unknown as InstanceWindow).__fcInst.workbook.sheet(0).protect(),
  );
  await sheet.goTo('B3');
  await expectFilterButtonInert(page);
  expect(await readAutoFilterRange(page)).toEqual(range);

  await page.evaluate(async () => {
    const target = (window as unknown as InstanceWindow).__fcInst.workbook.sheet(0);
    await target.unprotect();
    await target.protect({ allowAutoFilter: true });
  });
  const point = await filterButtonPoint(page, 0, 1);
  await page.mouse.click(point.x, point.y);
  const dropdown = page.locator('.fc-filter-dropdown:visible').first();
  await expect(dropdown).toBeVisible();
  await page.keyboard.press('Escape');
  await expect(dropdown).toBeHidden();

  await sheet.goTo('B3');
  await sheet.shortcut('Shift+L');
  await page.waitForTimeout(300);
  expect(await readAutoFilterRange(page)).toEqual(range);
  await sheet.expectNoConsoleErrors();
}
