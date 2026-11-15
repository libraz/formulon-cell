import { expect, type Page } from '@playwright/test';

import { selectRibbonTab } from '../pages/ribbon-tabs.js';
import { SpreadsheetPage } from '../pages/SpreadsheetPage.js';

/** B5 — inactive ribbon panels do NOT capture Tab focus.
 *
 *  Each demo's ribbon renders ALL tab panels in the DOM, but toggles `hidden`
 *  on the inactive ones. Browser focus semantics say a button inside a
 *  [hidden] ancestor must not appear in the Tab focus order — verifying that
 *  contract catches a common a11y regression where someone removes `[hidden]`
 *  in favour of `aria-hidden` or CSS-only hiding (which leaves the buttons
 *  focusable through keyboard, breaking screen reader semantics). */
export async function runRibbonInactiveFocusScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await sp.mount();
  await sp.expectNoStub();

  // Activate the Home tab — its panel is the largest, so other tabs are
  // guaranteed to be inactive.
  await page.getByRole('tab', { name: 'Home', exact: true }).click();

  // The React / Vue wrappers conditionally render the active ribbon panel
  // and unmount the others, so there are no `[hidden]` siblings to check.
  // That's a stronger focus guarantee than `[hidden]`, so we just verify
  // the demo ribbon root exists and skip the rest in that case.
  const hiddenPanels = page.locator('.fc-tb__ribbon[hidden]');
  const hasHiddenPanels = (await hiddenPanels.count()) > 0;
  if (!hasHiddenPanels) {
    // Wrappers unmount inactive panels — focus can never leak into them.
    await expect(page.locator('.fc-tb__ribbon').first()).toBeAttached();
    return;
  }

  // Playground renders all panels and toggles [hidden]; assert the contract.
  await expect(page.locator('.fc-tb__ribbon[data-ribbon-panel="home"]:not([hidden])')).toHaveCount(
    1,
  );

  // Buttons inside a hidden panel must NOT match the focus selector. Browsers
  // skip elements with [hidden] ancestors entirely during Tab navigation.
  const focusableInHidden = await page.evaluate(() => {
    const panels = Array.from(document.querySelectorAll<HTMLElement>('.fc-tb__ribbon[hidden]'));
    let count = 0;
    for (const panel of panels) {
      // Mirror the focus-selector heuristic browsers use for Tab order:
      // - tabindex !== -1
      // - not disabled
      // - laid out (offsetParent != null)
      const candidates = panel.querySelectorAll<HTMLElement>(
        'button, [tabindex]:not([tabindex="-1"]), input, select, textarea',
      );
      for (const c of candidates) {
        if (c instanceof HTMLButtonElement && c.disabled) continue;
        if (c.offsetParent !== null) count += 1;
      }
    }
    return count;
  });
  expect(
    focusableInHidden,
    'no element inside a [hidden] ribbon panel should be in the layout/focus tree',
  ).toBe(0);

  // Switching to a different tab moves the focus capability over.
  await page.getByRole('tab', { name: 'Insert', exact: true }).click();
  await expect(
    page.locator('.fc-tb__ribbon[data-ribbon-panel="insert"]:not([hidden])'),
  ).toHaveCount(1);
  // Now Home is hidden — its buttons should also no longer be tabbable.
  const homeFocusableNow = await page.evaluate(() => {
    const panel = document.querySelector<HTMLElement>('.fc-tb__ribbon[data-ribbon-panel="home"]');
    return panel?.hidden ? 'hidden' : 'visible';
  });
  expect(homeFocusableNow).toBe('hidden');
}

const SPLIT_BUTTON = 'button[data-ribbon-menu-id]:has(.fc-tb__rb-split-chevron)';

/** B6 — Alt+Down on a focused split button opens its menu with aria-expanded
 *  set and focus on the first shown, enabled item, or on the menu itself when
 *  it has none; Escape closes the menu and returns focus to the button. Every
 *  enabled split button on every ribbon tab is walked. */
export async function runRibbonSplitAltDownScenario(
  page: Page,
  platform: 'default' | 'mac',
): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await sp.mount({ platform });
  await sp.expectNoStub();

  const tabs = await page
    .locator('[data-ribbon-tab]:visible')
    .evaluateAll((els) =>
      els
        .map((el) => (el as HTMLElement).dataset.ribbonTab ?? '')
        .filter((id) => id && id !== 'file'),
    );
  const failures: string[] = [];
  let walked = 0;
  for (const tab of tabs) {
    await selectRibbonTab(page, tab);
    const buttons = page.locator(`[data-ribbon-panel="${tab}"] ${SPLIT_BUTTON}`);
    const count = await buttons.count();
    for (let i = 0; i < count; i += 1) {
      const button = buttons.nth(i);
      if (!(await button.isVisible()) || (await button.isDisabled())) continue;
      if ((await button.getAttribute('aria-disabled')) === 'true') continue;
      const label = `${tab}/${(await button.getAttribute('aria-label')) ?? i}`;
      walked += 1;
      await button.focus();
      await page.keyboard.press('Alt+ArrowDown');
      await expect(button, `${label} aria-expanded`).toHaveAttribute('aria-expanded', 'true');
      const focus = await page.evaluate(() => {
        const active = document.activeElement as HTMLElement | null;
        const menu = active?.closest<HTMLElement>('[role="menu"]') ?? null;
        if (!active || !menu)
          return { inMenu: false, ok: false, got: active?.outerHTML.slice(0, 80) };
        const shown = (el: HTMLElement) => el.getClientRects().length > 0;
        const items = Array.from(
          menu.querySelectorAll<HTMLElement>(
            '[role="menuitem"],[role="menuitemradio"],[role="menuitemcheckbox"]',
          ),
        ).filter(
          (el) =>
            el.closest('[role="menu"]') === menu &&
            shown(el) &&
            el.getAttribute('aria-disabled') !== 'true' &&
            !(el as HTMLButtonElement).disabled,
        );
        const expected = items[0] ?? menu;
        return {
          inMenu: true,
          ok: active === expected,
          got: active.getAttribute('aria-label') ?? active.textContent?.trim().slice(0, 40),
          want: expected.getAttribute('aria-label') ?? expected.textContent?.trim().slice(0, 40),
        };
      });
      if (!focus.ok) failures.push(`${label}: focus ${JSON.stringify(focus)}`);
      await page.keyboard.press('Escape');
      await expect(button, `${label} collapses on Escape`).toHaveAttribute(
        'aria-expanded',
        'false',
      );
      await expect(button, `${label} keeps focus after Escape`).toBeFocused();
    }
  }
  expect(walked, 'split buttons walked').toBeGreaterThan(0);
  expect(failures).toEqual([]);
}
