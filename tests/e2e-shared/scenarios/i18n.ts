import type { Page } from '@playwright/test';
import { expect } from '@playwright/test';

import { SpreadsheetPage } from '../pages/SpreadsheetPage.js';

/** I01 — `?locale=ja` boots the app in Japanese. We don't have a stable
 *  Japanese-only chrome string we can match in all 3 apps, so we check the
 *  HTML lang or doc title shift after asking the page to re-render in ja.
 *
 *  React and Vue demos set the locale internally after booting. */
export async function runLocaleBootScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await page.goto('/?locale=ja');
  await sp.waitForReady();
  await sp.expectNoStub();
  const jaToggle = page.getByRole('button', { name: 'JA', exact: true });
  if ((await jaToggle.count()) > 0) await jaToggle.first().click();

  await expect(page.getByRole('tab', { name: 'ホーム', exact: true })).toBeVisible();
  await expect(page.getByRole('tab', { name: '挿入', exact: true })).toBeVisible();
  await expect(page.getByRole('button', { name: 'ペースト', exact: true }).first()).toBeVisible();
  const searchBox = page.getByRole('searchbox').first();
  const hasSearchBox = (await searchBox.count()) > 0;
  if (hasSearchBox) {
    await expect(searchBox).toHaveAttribute('aria-label', 'コマンドの検索');
    await searchBox.fill('書式');
    // Scope to the command menu so we don't also match the ribbon button
    // (which now lives in the shared toolbar and shares the same label).
    await expect(
      page.locator('.demo__command-menu').getByRole('button', { name: /セルの書式設定/ }),
    ).toBeVisible();
    await searchBox.evaluate((el: HTMLElement) => el.blur());
    await expect(page.locator('.demo__command-menu')).toHaveCount(0);
  }

  const lang = await page.evaluate(() => document.documentElement.lang);
  expect(lang === 'ja' || lang === '').toBe(true);
  expect(await sp.isCrossOriginIsolated()).toBe(true);

  const enToggle = page.getByRole('button', { name: 'EN', exact: true });
  if ((await enToggle.count()) > 0) {
    await enToggle.first().click();
    await expect(page.getByRole('tab', { name: 'Home', exact: true })).toBeVisible();
    await expect(page.getByRole('tab', { name: 'Insert', exact: true })).toBeVisible();
    await expect(page.getByRole('button', { name: 'Paste', exact: true }).first()).toBeVisible();
    if (hasSearchBox) {
      await expect(searchBox).toHaveAttribute('aria-label', 'Search commands');
      await searchBox.fill('format');
      await expect(page.getByRole('button', { name: /Format Cells/ })).toBeVisible();
      await searchBox.evaluate((el: HTMLElement) => el.blur());
      await expect(page.locator('.demo__command-menu')).toHaveCount(0);
    }
  } else {
    await page.goto('/?locale=en');
    await sp.waitForReady();
    await expect(page.getByRole('tab', { name: 'Home', exact: true })).toBeVisible();
    await expect(page.getByRole('tab', { name: 'Insert', exact: true })).toBeVisible();
    await expect(page.getByRole('button', { name: 'Paste', exact: true }).first()).toBeVisible();
  }
}

/** I02 — `?theme=dark` boots the app in the `ink` core theme.
 *  Observable via the host's `data-fc-theme` attribute. */
export async function runThemeBootScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await page.goto('/?theme=dark');
  await sp.waitForReady();

  const themeAttr = await page.evaluate(() => {
    const host = document.querySelector('.fc-host') as HTMLElement | null;
    return host?.dataset.fcTheme ?? null;
  });
  expect(themeAttr === 'ink' || themeAttr === 'dark').toBe(true);
}
