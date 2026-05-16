import { expect, test } from '@playwright/test';

import { mountVisualPage } from './helpers.js';

const ribbonTabs = [
  { id: 'file', label: 'File' },
  { id: 'home', label: 'Home' },
  { id: 'insert', label: 'Insert' },
  { id: 'draw', label: 'Draw' },
  { id: 'pageLayout', label: 'Page Layout' },
  { id: 'formulas', label: 'Formulas' },
  { id: 'data', label: 'Data' },
  { id: 'review', label: 'Review' },
  { id: 'view', label: 'View' },
  { id: 'automate', label: 'Automate' },
  { id: 'acrobat', label: 'Acrobat' },
] as const;

for (const tab of ribbonTabs) {
  test(`@visual ribbon baseline — ${tab.id}`, async ({ page }) => {
    await mountVisualPage(page, '/?theme=light&locale=en');
    await page.getByRole('tab', { name: tab.label, exact: true }).click();

    const ribbon = page.locator('.app__ribbon-shell').first();
    await expect(ribbon).toBeVisible();
    if (tab.id === 'file') {
      await expect(page.locator('.demo__backstage')).toBeVisible();
      await expect(page).toHaveScreenshot('ribbon-file.png', {
        maxDiffPixels: 120,
        animations: 'disabled',
      });
      return;
    }
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

test('@visual ribbon collapsed — tabs only', async ({ page }) => {
  await mountVisualPage(page, '/?theme=light&locale=en');
  const home = page.getByRole('tab', { name: 'Home', exact: true });
  await home.focus();
  await page.keyboard.press('Control+F1');

  const ribbon = page.locator('.app__ribbon-shell').first();
  await expect(ribbon).toHaveClass(/demo__ribbon-shell--collapsed/);
  await expect(page.locator('.demo__ribbon:not([hidden])')).not.toBeVisible();
  await expect(ribbon).toHaveScreenshot('ribbon-collapsed-tabs-only.png', {
    maxDiffPixels: 80,
    animations: 'disabled',
  });
});

test('@visual ribbon display options menu', async ({ page }) => {
  await mountVisualPage(page, '/?theme=light&locale=en');
  await page.getByRole('button', { name: 'Ribbon Display Options' }).click();

  await expect(page.getByRole('menuitemradio', { name: 'Always show Ribbon' })).toBeVisible();
  await expect(page).toHaveScreenshot('ribbon-display-options-menu.png', {
    maxDiffPixels: 80,
    animations: 'disabled',
  });
});

test('@visual ribbon dropdown — page layout margins', async ({ page }) => {
  await mountVisualPage(page, '/?theme=light&locale=ja');
  await page.getByRole('tab', { name: 'ページ レイアウト', exact: true }).click();
  await page.locator('[data-ribbon-select="marginsPreset"] .demo__rb-dd__btn').click();

  await expect(
    page.locator('[data-ribbon-select="marginsPreset"] .demo__rb-dd__list'),
  ).toBeVisible();
  await expect(page).toHaveScreenshot('ribbon-page-layout-margins-dropdown.png', {
    maxDiffPixels: 80,
    animations: 'disabled',
  });
});

test('@visual ribbon dropdown — font family', async ({ page }) => {
  await mountVisualPage(page, '/?theme=light&locale=ja');
  await page.getByRole('tab', { name: 'ホーム', exact: true }).click();
  await page.locator('[data-ribbon-select="fontFamily"] .demo__rb-dd__btn').click();

  await expect(page.locator('[data-ribbon-select="fontFamily"] .demo__rb-dd__list')).toBeVisible();
  await expect(page).toHaveScreenshot('ribbon-font-family-dropdown.png', {
    maxDiffPixels: 80,
    animations: 'disabled',
  });
});

test('@visual ribbon dropdown — number format', async ({ page }) => {
  await mountVisualPage(page, '/?theme=light&locale=ja');
  await page.getByRole('tab', { name: 'ホーム', exact: true }).click();
  await page.locator('[data-ribbon-select="numberFormat"] .demo__rb-dd__btn').click();

  await expect(
    page.locator('[data-ribbon-select="numberFormat"] .demo__rb-dd__list'),
  ).toBeVisible();
  await expect(page).toHaveScreenshot('ribbon-number-format-dropdown.png', {
    maxDiffPixels: 80,
    animations: 'disabled',
  });
});

// Excel-365 borders dropdown — exercises the three preset sections, the
// "罫線の作成" header, and the draw / grid / erase / submenu rows. Run in
// Japanese to lock the layout consumers care about most.
test('@visual ribbon dropdown — borders (ja)', async ({ page }) => {
  await mountVisualPage(page, '/?theme=light&locale=ja');
  await page.getByRole('tab', { name: 'ホーム', exact: true }).click();
  await page.locator('#btn-borders').click();
  const menu = page.locator('#menu-borders');
  await expect(menu).toBeVisible();
  await expect(page).toHaveScreenshot('ribbon-borders-dropdown.png', {
    maxDiffPixels: 80,
    animations: 'disabled',
  });
});

// Same dropdown in English — locale differences in label width / wrapping
// should not break alignment.
test('@visual ribbon dropdown — borders (en)', async ({ page }) => {
  await mountVisualPage(page, '/?theme=light&locale=en');
  await page.getByRole('tab', { name: 'Home', exact: true }).click();
  await page.locator('#btn-borders').click();
  const menu = page.locator('#menu-borders');
  await expect(menu).toBeVisible();
  await expect(page).toHaveScreenshot('ribbon-borders-dropdown-en.png', {
    maxDiffPixels: 80,
    animations: 'disabled',
  });
});

// Line-style submenu (image #2). Hover the trigger row to reveal the
// patterns and capture the side panel.
test('@visual ribbon dropdown — line style submenu', async ({ page }) => {
  await mountVisualPage(page, '/?theme=light&locale=ja');
  await page.getByRole('tab', { name: 'ホーム', exact: true }).click();
  await page.locator('#btn-borders').click();
  const lineStyleTrigger = page.locator('[data-border-submenu="lineStyle"]');
  await lineStyleTrigger.dispatchEvent('click');
  await expect(lineStyleTrigger).toHaveAttribute('aria-expanded', 'true');
  const submenu = page.locator('.app__submenu--line-style');
  await expect(submenu).toBeVisible();
  await expect(page).toHaveScreenshot('ribbon-borders-line-style.png', {
    maxDiffPixels: 80,
    animations: 'disabled',
  });
});
