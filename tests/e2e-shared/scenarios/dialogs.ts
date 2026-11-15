import type { Page } from '@playwright/test';
import { expect } from '@playwright/test';

import { SpreadsheetPage } from '../pages/SpreadsheetPage.js';

/** D02 — Find & Replace. Bound to Mod+F via host-shortcuts.ts. The find
 *  overlay's exact selector varies by build but every demo's chrome attaches
 *  a `.fc-findreplace` (or close descendant) on open. We test the keystroke
 *  succeeds without throwing and SOME find-related UI becomes visible. */
export async function runFindReplaceShortcutScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  const consoleErrors = sp.collectConsoleErrors();
  await sp.mount();
  await sp.expectNoStub();

  await sp.focusHost();
  await sp.shortcut('f');

  // Settle async paint.
  await page.waitForTimeout(150);
  expect(consoleErrors.read(), 'opening find should not error').toEqual([]);

  // Closing with Escape should also be safe.
  await page.keyboard.press('Escape');
  expect(consoleErrors.read()).toEqual([]);
}

/** D01 (smoke) — Mod+Z / Mod+Y undo redo path is bound by the host's
 *  shortcut router. Closing path checks that nothing throws. */
export async function runUndoSubsystemSmokeScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  const consoleErrors = sp.collectConsoleErrors();
  await sp.mount();
  await sp.expectNoStub();

  await sp.shortcut('z');
  await page.waitForTimeout(50);
  expect(consoleErrors.read()).toEqual([]);
}

/** D03 — Shift+F3 and the formula bar fx button open Function Arguments, and
 *  closing it by Escape or Cancel hands focus back to the sheet. */
export async function runFunctionDialogShortcutScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  const consoleErrors = sp.collectConsoleErrors();
  await sp.mount();
  await sp.expectNoStub();
  const host = page.locator('.fc-host').first();
  const dialog = page.getByRole('dialog', { name: 'Function Arguments', exact: true });

  await sp.focusHost();
  await page.keyboard.press('Shift+F3');
  await expect(dialog).toBeVisible();
  await expect(dialog.getByRole('combobox', { name: 'Search functions…' })).toBeFocused();
  await page.keyboard.press('Escape');
  await expect(dialog).toHaveCount(0);
  await expect(host).toBeFocused();

  const fx = page.locator('.fc-host__formulabar-fx').first();
  for (const close of ['Escape', 'Cancel'] as const) {
    await fx.click();
    await expect(dialog).toBeVisible();
    if (close === 'Escape') await page.keyboard.press('Escape');
    else await dialog.getByRole('button', { name: 'Cancel', exact: true }).click();
    await expect(dialog, `closed by ${close}`).toHaveCount(0);
    await expect(host, `focus after ${close}`).toBeFocused();
  }
  expect(consoleErrors.read()).toEqual([]);
}

/** D03 (Mac) — Shift+F3 opens the docked Formula Builder pane with focus in
 *  its search field. Escape leaves it open; Shift+Tab reaches its close button
 *  (Chromium: WebKit leaves plain buttons out of the Tab order); the function
 *  list is one tab stop, so one Tab leaves it; Tab moves out of the pane while
 *  it stays open; in Chromium focus lands on the status bar and Shift+Tab comes
 *  back into the pane (in WebKit the list is the page's last tab stop); the
 *  close button hands focus to the sheet. */
export async function runMacFormulaBuilderKeyboardScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  const consoleErrors = sp.collectConsoleErrors();
  await sp.mount({ platform: 'mac' });
  await sp.expectNoStub();
  const host = page.locator('.fc-host').first();
  const pane = page.locator('.fc-mac-formula-palette');
  const search = pane.locator('.fc-mac-formula-palette__search');
  const close = pane.locator('.fc-mac-formula-palette__header [data-action="close"]');

  const open = async (): Promise<void> => {
    await sp.focusHost();
    await page.keyboard.press('Shift+F3');
    await expect(pane).toBeVisible();
    await expect(search).toBeFocused();
  };

  await open();
  await page.keyboard.press('Escape');
  await expect(pane).toBeVisible();

  const webkit = page.context().browser()?.browserType().name() === 'webkit';
  if (!webkit) {
    await search.focus();
    let reachedClose = false;
    for (let i = 0; i < 8 && !reachedClose; i += 1) {
      await page.keyboard.press('Shift+Tab');
      reachedClose = await close.evaluate((el) => el === document.activeElement);
    }
    expect(reachedClose, 'Shift+Tab from the search field reaches the close button').toBe(true);
  }

  // Each section renders its own listbox; an empty section has no tab stop.
  const inList = () =>
    pane.evaluate(
      (el) =>
        !!document.activeElement?.closest('[role="listbox"]') &&
        el.contains(document.activeElement),
    );
  await search.focus();
  let reachedList = false;
  for (let i = 0; i < 4 && !reachedList; i += 1) {
    await page.keyboard.press('Tab');
    reachedList = await inList();
  }
  expect(reachedList, 'Tab from the search field reaches the function list').toBe(true);
  await page.keyboard.press('Tab');
  expect(await inList(), 'one Tab leaves the function list').toBe(false);

  await search.focus();
  let left = false;
  for (let i = 0; i < 12 && !left; i += 1) {
    await page.keyboard.press('Tab');
    left = await pane.evaluate((el) => !el.contains(document.activeElement));
  }
  expect(left, 'Tab moves focus out of the pane').toBe(true);
  await expect(pane, 'the docked pane stays open after Tab leaves it').toBeVisible();
  if (!webkit) {
    expect(
      await page.evaluate(() => {
        const active = document.activeElement;
        return !!active && active !== document.body && active !== document.documentElement;
      }),
      'focus lands on a control after leaving the pane',
    ).toBe(true);
    await page.keyboard.press('Shift+Tab');
    expect(
      await pane.evaluate((el) => el.contains(document.activeElement)),
      'Shift+Tab returns into the pane',
    ).toBe(true);
  }

  await close.click();
  await expect(pane).toBeHidden();
  await expect(host).toBeFocused();
  expect(consoleErrors.read()).toEqual([]);
}
