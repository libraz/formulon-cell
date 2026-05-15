import { expect, test } from '@playwright/test';

import {
  runSheetTabSwitchScenario,
  runSheetTabsScenario,
} from '../../../tests/e2e-shared/scenarios/sheet-tabs.js';

test('T01 (playground): the + button adds a new sheet and makes it active', async ({ page }) => {
  await runSheetTabsScenario(page);
});

test('T02 (playground): clicking a tab switches the active sheet', async ({ page }) => {
  await runSheetTabSwitchScenario(page);
});

test('T03 (playground): sheet tab menu supports menu keys and Escape focus return', async ({
  page,
}) => {
  await page.goto('/');
  await page.waitForSelector('.fc-host', { state: 'attached', timeout: 30_000 });
  await page.locator('#btn-sheet-add').click();
  await expect(page.locator('.app__tab')).toHaveCount(2);
  const firstTab = page.locator('.app__tab').first();
  await firstTab.focus();
  await firstTab.click({ button: 'right' });

  const menu = page.getByRole('menu', { name: 'Sheet tab' });
  await expect(menu).toBeVisible();
  await expect(menu.getByRole('menuitem', { name: 'Rename…' })).toBeFocused();
  await page.keyboard.press('ArrowDown');
  await expect(menu.getByRole('menuitem', { name: 'Delete' })).toBeFocused();
  await page.keyboard.press('End');
  await expect(menu.getByRole('menuitem', { name: 'Move right' })).toBeFocused();
  await page.keyboard.press('Escape');
  await expect(menu).toBeHidden();
  await expect(firstTab).toBeFocused();
});
