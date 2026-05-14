import { expect, test } from '@playwright/test';

import { mountVisualPage } from './helpers.js';

const ribbonTabs = [
  { id: 'home', label: 'Home' },
  { id: 'insert', label: 'Insert' },
  { id: 'pageLayout', label: 'Page Layout' },
  { id: 'data', label: 'Data' },
  { id: 'review', label: 'Review' },
  { id: 'view', label: 'View' },
] as const;

for (const tab of ribbonTabs) {
  test(`@visual ribbon baseline — ${tab.id}`, async ({ page }) => {
    await mountVisualPage(page, '/?theme=light&locale=en');
    await page.getByRole('tab', { name: tab.label, exact: true }).click();

    const ribbon = page.locator('.app__ribbon-shell').first();
    await expect(ribbon).toBeVisible();
    await expect(page.locator('.demo__ribbon:not([hidden])')).toHaveAttribute(
      'data-ribbon-panel',
      tab.id,
    );

    await expect(ribbon).toHaveScreenshot(`ribbon-${tab.id}.png`, {
      maxDiffPixels: 80,
      animations: 'disabled',
    });
  });
}
