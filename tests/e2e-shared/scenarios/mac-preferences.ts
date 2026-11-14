import { expect, type Page } from '@playwright/test';

import { UserJourneyPage } from '../pages/UserJourneyPage.js';

/** File > Options opens the Mac Preferences window. Its toolbar carries Back and Forward, Show
 *  All, the shown pane's name and Search; Show All lists the panes, a pane button opens one, and
 *  Back steps back to Show All. */
export async function runMacPreferencesScenario(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount({ platform: 'mac' });
  await page.setViewportSize({ width: 1440, height: 900 });
  await sp.expectNoStub();
  await page.locator('.demo__brand-mark').first().click();
  const backstage = page.locator('.fc-tb__backstage[role="dialog"]').first();
  await expect(backstage).toBeVisible();
  await backstage.getByRole('button', { name: 'Options', exact: true }).first().click();

  const prefs = page.locator('.fc-prefs[role="dialog"]');
  await expect(prefs).toBeVisible();
  const toolbar = prefs.locator('[role="toolbar"]');
  const back = toolbar.getByRole('button', { name: 'Back', exact: true });
  const forward = toolbar.getByRole('button', { name: 'Forward', exact: true });
  const showAll = toolbar.getByRole('button', { name: 'Show All', exact: true });
  const heading = toolbar.locator('.fc-prefs__heading');
  await expect(back).toBeDisabled();
  await expect(forward).toBeDisabled();
  await expect(showAll).toBeVisible();
  await expect(toolbar.getByRole('searchbox', { name: 'Search', exact: true })).toBeVisible();
  // The window opens on its first pane.
  await expect(heading).toHaveText('View');
  await expect(prefs.locator('.fc-prefs__pane[data-pane="view"]')).toBeVisible();

  await showAll.click();
  await expect(heading).toHaveText('');
  await expect(prefs.locator('.fc-prefs__pane:not([hidden])')).toHaveCount(0);
  await expect(back).toBeEnabled();
  await expect(forward).toBeDisabled();

  await prefs.locator('.fc-prefs__pane-button[data-pane="view"]').click();
  await expect(heading).toHaveText('View');
  await expect(prefs).toHaveAttribute('aria-label', 'View');
  await expect(prefs.locator('.fc-prefs__pane[data-pane="view"]')).toBeVisible();
  await expect(back).toBeEnabled();

  await back.click();
  await expect(heading).toHaveText('');
  await expect(prefs).toHaveAttribute('aria-label', 'Preferences');
  await expect(prefs.locator('.fc-prefs__pane-button[data-pane="view"]')).toBeVisible();
  await expect(forward).toBeEnabled();
  await sp.expectNoConsoleErrors();
}

/** The View pane title as the desktop draws it per UI language. */
const VIEW_PANE_TITLE = {
  en: 'View',
  ja: '表示',
  de: 'Ansicht',
  fr: 'Affichage',
  zh: '视图',
} as const;

/** The File backstage's Options entry per UI language. */
const OPTIONS_LABEL = {
  en: 'Options',
  ja: 'オプション',
  de: 'Optionen',
  fr: 'Options',
  zh: '选项',
} as const;

/** The Preferences window opens on View, and its toolbar heading and accessible name carry the
 *  pane title in the UI language. */
export async function runMacPreferencesViewTitleScenario(
  page: Page,
  locale: keyof typeof VIEW_PANE_TITLE,
): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount({ platform: 'mac', locale });
  await page.setViewportSize({ width: 1440, height: 900 });
  await sp.expectNoStub();
  await page.locator('.demo__brand-mark').first().click();
  const backstage = page.locator('.fc-tb__backstage[role="dialog"]').first();
  await expect(backstage).toBeVisible();
  await backstage.getByRole('button', { name: OPTIONS_LABEL[locale], exact: true }).first().click();
  const prefs = page.locator('.fc-prefs[role="dialog"]:not([hidden])');
  await expect(prefs).toBeVisible();
  const title = VIEW_PANE_TITLE[locale];
  await expect(prefs.locator('.fc-prefs__pane[data-pane="view"]')).toBeVisible();
  await expect(prefs.locator('.fc-prefs__heading')).toHaveText(title);
  await expect(prefs.locator('.fc-prefs__header')).toContainText(title);
  await expect(prefs).toHaveAttribute('aria-label', title);
  await sp.expectNoConsoleErrors();
}
