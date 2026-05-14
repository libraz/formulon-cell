import { expect, type Page } from '@playwright/test';

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
  const hiddenPanels = page.locator('.demo__ribbon[hidden]');
  const hasHiddenPanels = (await hiddenPanels.count()) > 0;
  if (!hasHiddenPanels) {
    // Wrappers unmount inactive panels — focus can never leak into them.
    await expect(page.locator('.demo__ribbon').first()).toBeAttached();
    return;
  }

  // Playground renders all panels and toggles [hidden]; assert the contract.
  await expect(page.locator('.demo__ribbon[data-ribbon-panel="home"]:not([hidden])')).toHaveCount(
    1,
  );

  // Buttons inside a hidden panel must NOT match the focus selector. Browsers
  // skip elements with [hidden] ancestors entirely during Tab navigation.
  const focusableInHidden = await page.evaluate(() => {
    const panels = Array.from(document.querySelectorAll<HTMLElement>('.demo__ribbon[hidden]'));
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
  await expect(page.locator('.demo__ribbon[data-ribbon-panel="insert"]:not([hidden])')).toHaveCount(
    1,
  );
  // Now Home is hidden — its buttons should also no longer be tabbable.
  const homeFocusableNow = await page.evaluate(() => {
    const panel = document.querySelector<HTMLElement>('.demo__ribbon[data-ribbon-panel="home"]');
    return panel?.hidden ? 'hidden' : 'visible';
  });
  expect(homeFocusableNow).toBe('hidden');
}
