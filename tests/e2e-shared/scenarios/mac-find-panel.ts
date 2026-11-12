import { expect, type Page } from '@playwright/test';

import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/api/spreadsheet/instance.js';
import { UserJourneyPage } from '../pages/UserJourneyPage.js';

type InstanceWindow = Window & { __fcInst: SpreadsheetInstance };

const readActive = (page: Page) =>
  page.evaluate(() => {
    const { row, col } = (window as unknown as InstanceWindow).__fcInst.selection.get().active;
    return { row, col };
  });

const inputFocused = (page: Page) =>
  page.evaluate(() => document.activeElement?.classList.contains('fc-findpanel__input') ?? false);

/** The Mac ⌘F find panel: it floats at the top right under the ribbon, Return and Shift+Return
 *  both step to the next match and wrap, the panel stays open on a match, Escape closes it only
 *  while it has focus, and a search with no match raises the not-found alert. */
export async function runMacFindPanelScenario(page: Page): Promise<void> {
  const sheet = new UserJourneyPage(page);
  await sheet.mount({ locale: 'ja', platform: 'mac' });
  await sheet.expectNoStub();
  await sheet.seed({ A1: 'x', B2: 'apple', C3: 'pear', D4: 'apple' });
  await sheet.goTo('A1');

  await sheet.shortcut('f');
  const panel = page.locator('.fc-findpanel');
  const input = panel.locator('.fc-findpanel__input');
  await expect(panel).toBeVisible();
  await expect(page.locator('.fc-find')).toBeHidden();
  await expect.poll(() => inputFocused(page)).toBe(true);

  const box = await panel.boundingBox();
  const ribbon = await page.locator('[data-ribbon-tab="home"]').first().boundingBox();
  const grid = await page.locator('.fc-host__grid').first().boundingBox();
  if (!box || !ribbon || !grid) throw new Error('find panel, ribbon or grid not laid out');
  expect(box.y, 'under the ribbon').toBeGreaterThanOrEqual(ribbon.y + ribbon.height);
  expect(grid.x + grid.width - (box.x + box.width), 'at the right edge').toBeLessThan(40);
  expect(box.x, 'on the right half').toBeGreaterThan(grid.x + grid.width / 2);

  await input.fill('apple');
  const steps: [string, { row: number; col: number }][] = [
    ['Enter', { row: 1, col: 1 }],
    ['Enter', { row: 3, col: 3 }],
    ['Enter', { row: 1, col: 1 }],
    ['Shift+Enter', { row: 3, col: 3 }],
  ];
  for (const [key, active] of steps) {
    await page.keyboard.press(key);
    await expect.poll(() => readActive(page), key).toEqual(active);
    await expect(panel).toBeVisible();
    expect(await inputFocused(page), key).toBe(true);
  }

  // With the grid focused, Escape leaves the panel open; with the panel focused, it closes it.
  await page
    .locator('.fc-host__grid')
    .first()
    .click({ position: { x: 60, y: 120 } });
  await expect.poll(() => inputFocused(page)).toBe(false);
  await page.keyboard.press('Escape');
  await page.waitForTimeout(300);
  await expect(panel).toBeVisible();
  await input.click();
  await page.keyboard.press('Escape');
  await expect(panel).toBeHidden();

  await sheet.goTo('A1');
  await sheet.shortcut('f');
  await expect(panel).toBeVisible();
  await input.fill('banana');
  await page.keyboard.press('Enter');
  const alert = page.locator('.fc-tb__dlg--alert');
  await expect(alert).toBeVisible();
  await expect(alert.locator('.fc-tb__dlg__headline')).not.toBeEmpty();
  expect(await readActive(page)).toEqual({ row: 0, col: 0 });
  await alert.getByRole('button').first().click();
  await expect(alert).toBeHidden();
  await sheet.expectNoConsoleErrors();
}
