import type { Page } from '@playwright/test';
import { expect } from '@playwright/test';

import { SpreadsheetPage } from '../pages/SpreadsheetPage.js';
import { UserJourneyPage } from '../pages/UserJourneyPage.js';

/** T01 — adding a sheet via the `+` button in the sheetbar increases the
 *  tab count and makes the new tab the selected one.
 *
 *  Uses only core chrome (`.fc-host__sheetbar-*`), so it runs identically
 *  across the three demo apps. */
export async function runSheetTabsScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await sp.mount();
  await sp.expectNoStub();

  const tabs = page.locator('.fc-host__sheetbar-tab');
  const initial = await tabs.count();
  expect(initial, 'demo apps boot with at least one sheet').toBeGreaterThanOrEqual(1);

  await page.locator('.fc-host__sheetbar-add').click();

  // The new tab gets aria-selected="true". Wait until the count changes so we
  // don't race the re-render in `update()` (sheet-tabs-controller.ts:221).
  await expect(tabs).toHaveCount(initial + 1);

  const selected = page.locator('.fc-host__sheetbar-tab[aria-selected="true"]');
  await expect(selected).toHaveCount(1);
}

/** T02 — clicking a sibling tab switches the active sheet. */
export async function runSheetTabSwitchScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await sp.mount();
  await sp.expectNoStub();

  // Make sure we have at least two sheets.
  const addBtn = page.locator('.fc-host__sheetbar-add');
  const tabs = page.locator('.fc-host__sheetbar-tab');
  if ((await tabs.count()) < 2) {
    await addBtn.click();
    await expect(tabs).toHaveCount(2);
  }

  // Click the first tab and verify selection moves.
  await tabs.first().click();
  await expect(tabs.first()).toHaveAttribute('aria-selected', 'true');
}

/** T03 — Move or Copy with "Create a copy" from the tab menu adds a named copy holding the source's
 *  cells and activates it; undo removes the copy and redo restores it. */
export async function runSheetTabCopyScenario(page: Page): Promise<void> {
  const sheet = new UserJourneyPage(page);
  await sheet.mount({ locale: 'en' });
  await sheet.expectNoStub();
  await sheet.seed({ A1: 42, B2: '=A1*2' });

  const tabs = page.locator('.fc-host__sheetbar-tab');
  const initial = await sheet.sheetCount();
  const source = await sheet.sheetName(0);
  await sheet.chooseSheet(source);

  await tabs.filter({ hasText: source }).first().click({ button: 'right' });
  const menu = page.locator('.fc-sheetmenu');
  await expect(menu).toBeVisible();
  await menu.getByRole('menuitem', { name: 'Move or Copy...' }).click();

  const dialog = page.locator('.fc-movecopy');
  await expect(dialog).toBeVisible();
  await dialog.locator('.fc-movecopy__item').last().click();
  await dialog.locator('.fc-movecopy__copy input[type="checkbox"]').check();
  await dialog.getByRole('button', { name: 'OK' }).click();
  await expect(dialog).toBeHidden();

  const copyName = `${source} (2)`;
  await expect.poll(() => sheet.sheetCount()).toBe(initial + 1);
  expect(await sheet.sheetName(initial)).toBe(copyName);
  await expect(page.locator('.fc-host__sheetbar-tab[aria-selected="true"]')).toHaveText(copyName);
  expect(await sheet.readValue('A1', initial)).toMatchObject({ kind: 'number', value: 42 });
  expect(await sheet.readFormula('B2', initial)).toBe('=A1*2');

  await sheet.focusHost();
  await sheet.shortcut('z');
  await expect.poll(() => sheet.sheetCount()).toBe(initial);
  await expect(tabs.filter({ hasText: copyName })).toHaveCount(0);

  await sheet.shortcut('y');
  await expect.poll(() => sheet.sheetCount()).toBe(initial + 1);
  expect(await sheet.sheetName(initial)).toBe(copyName);
  expect(await sheet.readValue('A1', initial)).toMatchObject({ kind: 'number', value: 42 });
}

/** Shift+Tab moves backwards out of the sheet bar: from the Add sheet button and
 *  from the active sheet tab, focus leaves for another control instead of being
 *  held by the grid's key handling. */
export async function runSheetBarShiftTabScenario(
  page: Page,
  platform: 'default' | 'mac',
): Promise<void> {
  const sp = new SpreadsheetPage(page);
  const consoleErrors = sp.collectConsoleErrors();
  await sp.mount({ platform });
  await sp.expectNoStub();
  const sheetbar = page.locator('.fc-host__sheetbar').first();
  const starts = {
    'Add sheet button': page.locator('.fc-host__sheetbar-add').first(),
    'active sheet tab': page.locator('.fc-host__sheetbar-tab[aria-selected="true"]').first(),
  };
  for (const [name, start] of Object.entries(starts)) {
    await start.focus();
    await expect(start, `${name} takes focus`).toBeFocused();
    await page.keyboard.press('Shift+Tab');
    await expect(start, `Shift+Tab moves focus off the ${name}`).not.toBeFocused();
    expect(
      await page.evaluate(() => {
        const active = document.activeElement;
        return !!active && active !== document.body && active !== document.documentElement;
      }),
      `Shift+Tab from the ${name} lands on a control`,
    ).toBe(true);
  }
  expect(
    await sheetbar.evaluate((el) => el.contains(document.activeElement)),
    'Shift+Tab from the active sheet tab leaves the sheet bar',
  ).toBe(false);
  expect(consoleErrors.read()).toEqual([]);
}
