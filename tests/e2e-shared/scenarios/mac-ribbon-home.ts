import { expect, type Page } from '@playwright/test';

import { UserJourneyPage } from '../pages/UserJourneyPage.js';

export async function runMacHomeMenusScenario(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount({ platform: 'mac' });
  await sp.expectNoStub();
  await sp.enter('A1', 'Formatted');
  await sp.goTo('A1');
  const underline = page.locator('.fc-tb__rb[data-ribbon-command="underline"]');
  await underline.locator('.fc-tb__rb-split-chevron').click();
  const menu = page.locator('#menu-underline');
  await expect(menu).toBeVisible();
  await menu.locator('[data-underline-action="double"]').click();
  const read = () =>
    page.evaluate(
      () =>
        (
          window as unknown as {
            __fcInst: {
              store: { getState(): { format: { formats: Map<string, { underline?: string }> } } };
            };
          }
        ).__fcInst.store
          .getState()
          .format.formats.get('0:0:0')?.underline,
    );
  await expect.poll(read).toBe('double');
  await page.locator('.fc-host').focus();
  await page.keyboard.press('Meta+z');
  await expect.poll(read).not.toBe('double');
  await page.locator('.fc-tb__rb[data-ribbon-command="merge"] .fc-tb__rb-split-chevron').click();
  await expect(page.locator('#menu-merge')).toBeVisible();
  await page.keyboard.press('Escape');
  await expect(page.locator('#menu-merge')).toBeHidden();
  await sp.expectNoConsoleErrors();
}
