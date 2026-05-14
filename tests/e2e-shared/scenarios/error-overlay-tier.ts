import { expect, type Page } from '@playwright/test';

import { SpreadsheetPage } from '../pages/SpreadsheetPage.js';

/** B9 — the error-info popover lives on the topmost overlay tier.
 *
 *  The error menu (`.fc-errmenu`) is internal — there's no public API to open
 *  it from a test, so we side-step the trigger and validate the runtime CSS
 *  alone: a div bearing the overlay class should resolve to `--fc-z-error`
 *  (≈2,147,483,070), which by spec is above context-menu, dialog, callout,
 *  popover, tooltip and grid tiers.
 *
 *  This catches the bug where someone moves `.fc-errmenu` to a different tier
 *  or drops the `z-index` declaration in a refactor — the error popover would
 *  silently render under whatever other overlay is open at the time. */
export async function runErrorOverlayTierScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await sp.mount();
  await sp.expectNoStub();

  const probe = await page.evaluate(() => {
    const make = (cls: string): HTMLElement => {
      const el = document.createElement('div');
      el.className = cls;
      // The actual selector blocks live inside an @layer cascade; mounting
      // the dummy inside .fc-host inherits the right tokens.
      const host = document.querySelector('.fc-host') as HTMLElement | null;
      (host ?? document.body).appendChild(el);
      return el;
    };
    const read = (el: HTMLElement): number => {
      const z = getComputedStyle(el).zIndex;
      return Number.parseInt(z, 10);
    };
    const err = make('fc-errmenu');
    const menu = make('fc-ctxmenu');
    const popover = make('fc-filter-dropdown');
    const dialog = make('fc-fmtdlg');

    const out = {
      err: read(err),
      menu: read(menu),
      popover: read(popover),
      dialog: read(dialog),
    };
    for (const el of [err, menu, popover, dialog]) el.remove();
    return out;
  });

  // Documented tiers from core/base.css.
  expect(probe.err, 'error tier').toBeGreaterThanOrEqual(2_147_483_070);
  expect(probe.menu, 'menu tier').toBeGreaterThanOrEqual(2_147_483_060);
  expect(probe.popover, 'popover tier').toBeGreaterThanOrEqual(2_147_483_040);
  expect(probe.dialog, 'dialog tier').toBeGreaterThanOrEqual(2_147_483_020);

  // The actual ordering invariants we care about for the error popover:
  expect(probe.err).toBeGreaterThan(probe.menu);
  expect(probe.err).toBeGreaterThan(probe.popover);
  expect(probe.err).toBeGreaterThan(probe.dialog);
}
