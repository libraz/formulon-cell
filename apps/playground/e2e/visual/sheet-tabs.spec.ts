import { expect, test } from '@playwright/test';

import { mountVisualPage } from './helpers.js';

test('@visual sheet tabs — tab color and palette', async ({ page }) => {
  await mountVisualPage(page, '/?theme=light&locale=en');
  await page.locator('#btn-sheet-add').click();
  await expect(page.locator('.app__tab')).toHaveCount(2);

  const firstTab = page.locator('.app__tab').first();
  await firstTab.click({ button: 'right' });
  const menu = page.getByRole('menu', { name: 'Sheet tab' });
  await expect(menu).toBeVisible();
  await menu.getByRole('menuitemradio', { name: 'Tab Color #c00000' }).click();

  const bottomBar = page.locator('.app__bottombar');
  await expect(page.locator('.app__tab').first()).toHaveAttribute('data-sheet-tab-color', 'true');
  await expect(bottomBar).toHaveScreenshot('sheet-tabs-colored-tab.png', {
    maxDiffPixels: 80,
    animations: 'disabled',
  });

  await page.locator('.app__tab').first().click({ button: 'right' });
  await expect(menu).toBeVisible();
  await expect(menu).toHaveScreenshot('sheet-tabs-tab-color-menu.png', {
    maxDiffPixels: 80,
    animations: 'disabled',
  });
});
