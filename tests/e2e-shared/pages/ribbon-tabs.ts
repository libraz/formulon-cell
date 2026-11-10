import type { Locator, Page } from '@playwright/test';

/**
 * Brings a ribbon tab's panel on screen. A tab that is already selected with
 * its panel showing is left alone: pressing it again collapses the ribbon on
 * the Mac platform.
 */
export async function selectRibbonTab(page: Page, id: string): Promise<void> {
  const tab = page.locator(`[data-ribbon-tab="${id}"]`);
  const panel = page.locator(`[data-ribbon-panel="${id}"]`);
  if ((await tab.getAttribute('aria-selected')) === 'true' && (await panel.isVisible())) return;
  await tab.click();
}

/**
 * Switches the ribbon display mode through the toolbar instance, for hosts
 * where the display-options button is not part of the chrome (Mac).
 */
export async function setRibbonDisplayMode(
  page: Page,
  mode: 'full' | 'singleLine' | 'tabsOnly' | 'autoHide',
): Promise<void> {
  await page.evaluate((next) => {
    const toolbar = (
      window as unknown as { __fcToolbar?: { setDisplayMode(mode: string): void } | null }
    ).__fcToolbar;
    if (!toolbar) throw new Error('toolbar instance is not exposed');
    toolbar.setDisplayMode(next);
  }, mode);
}

/**
 * Opens the popup of the folded ribbon group that holds `button`, when the
 * group has folded into a single named button at the current window width.
 */
export async function revealRibbonCommand(button: Locator): Promise<void> {
  const group = button.locator(
    'xpath=ancestor::*[contains(concat(" ", normalize-space(@class), " "), " fc-tb__ribbon-group--collapsed ")][1]',
  );
  if ((await group.count()) === 0 || (await button.isVisible())) return;
  await group.locator(':scope > .fc-tb__ribbon-group-collapse').click();
}
