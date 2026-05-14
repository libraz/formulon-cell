import { expect, type Page } from '@playwright/test';

import { SpreadsheetPage } from '../pages/SpreadsheetPage.js';

/** B2 — overlays stack on the documented tiers at runtime.
 *
 *  `tests/unit/styles/overlay-z-index.test.ts` already locks down the *source*
 *  CSS tiers. This E2E asserts the rules actually paint at runtime — different
 *  build pipelines (Vite, Vue SFC, React) can re-order layers in subtle ways
 *  that the static CSS read won't catch.
 *
 *  Strategy: for each open-able overlay, capture the resolved `z-index` from
 *  `getComputedStyle`, then assert the cross-tier inequalities (dialog tier
 *  must rank above grid tier; menu must rank above dialog, etc.). */
export async function runOverlayStackingScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await sp.mount();
  await sp.expectNoStub();

  const readZ = async (selector: string): Promise<number> => {
    const z = await page.evaluate((sel) => {
      const el = document.querySelector(sel) as HTMLElement | null;
      if (!el) return null;
      const z = getComputedStyle(el).zIndex;
      const parsed = Number.parseInt(z, 10);
      return Number.isFinite(parsed) ? parsed : null;
    }, selector);
    expect(z, `expected a numeric z-index for ${selector}`).not.toBeNull();
    return z as number;
  };

  // 1. Open the format dialog — dialog tier.
  await sp.focusHost();
  await sp.shortcut('1');
  const dialog = page.locator('[class="fc-fmtdlg"]');
  await expect(dialog).toBeVisible({ timeout: 2_000 });
  const dialogZ = await readZ('[class="fc-fmtdlg"]');

  // 2. Close the dialog before opening another overlay (some overlays compete
  //    for focus). Then open Find — grid tier.
  await page.keyboard.press('Escape');
  await expect(dialog).toBeHidden({ timeout: 2_000 });

  await sp.shortcut('f');
  // Settle paint.
  await page.waitForTimeout(150);
  const findEl = page.locator('.fc-find').first();
  await expect(findEl).toBeAttached({ timeout: 2_000 });
  const findZ = await readZ('.fc-find');

  // Close find — Escape.
  await page.keyboard.press('Escape');

  // 3. Assert the cross-tier ordering. Dialog tier (2,147,483,020) must rank
  //    above grid tier (2,147,483,010). The exact integers come from
  //    core/base.css and are validated against the source by the unit test.
  expect(dialogZ).toBeGreaterThan(findZ);

  // 4. Validate the absolute floors are at least the documented minimums. If
  //    the build accidentally drops or rewrites the `--fc-z-*` custom
  //    properties the resolved values would be near-zero, which is a clear
  //    integration regression we want to catch.
  expect(dialogZ).toBeGreaterThanOrEqual(2_147_483_020);
  expect(findZ).toBeGreaterThanOrEqual(2_147_483_010);
}
