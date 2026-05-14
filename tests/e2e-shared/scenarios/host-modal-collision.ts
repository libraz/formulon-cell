import { expect, type Page } from '@playwright/test';

import { SpreadsheetPage } from '../pages/SpreadsheetPage.js';

/** B3 — formulon overlays stay above an aggressive host-side modal.
 *
 *  Real-world apps embed the spreadsheet inside their own UI. The host
 *  frequently pushes a modal/lightbox over the page (think Bootstrap modal,
 *  React-Modal, design-system Dialog) at z-index 1000–9999. The base CSS
 *  comment in `core/base.css` claims "every formulon overlay must rank above
 *  the typical host modal z-index of 1000–9999" — this test makes that claim
 *  observable. */
export async function runHostModalCollisionScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await sp.mount();
  await sp.expectNoStub();

  // Open the format dialog first (the host modal would intercept the click
  // we use to focus the canvas).
  await sp.focusHost();
  await sp.shortcut('1');

  // Inject a fake host modal at z-index 9999 (the upper edge of the
  // "typical host" range). It sits at the centre of the viewport so it
  // overlaps anywhere a formulon dialog could plausibly render.
  await page.evaluate(() => {
    const overlay = document.createElement('div');
    overlay.id = 'fake-host-modal';
    overlay.setAttribute('role', 'dialog');
    overlay.style.cssText = [
      'position: fixed',
      'inset: 0',
      'z-index: 9999',
      'background: rgba(0, 0, 0, 0.25)',
      'pointer-events: auto',
    ].join(';');
    document.body.appendChild(overlay);
  });
  const fmtDialog = page.locator('[class="fc-fmtdlg"]');
  await expect(fmtDialog).toBeVisible({ timeout: 2_000 });

  // Pick a point at the centre of the formulon dialog and ask the browser
  // which element wins. It MUST be the formulon dialog, not the host modal.
  const box = await fmtDialog.boundingBox();
  expect(box, 'dialog must be laid out').not.toBeNull();
  if (!box) throw new Error('dialog not measured');

  const probe = await page.evaluate(
    ({ x, y }) => {
      const el = document.elementFromPoint(x, y) as HTMLElement | null;
      const inFcDialog = el?.closest('[class~="fc-fmtdlg"]') !== null;
      const inHostModal = el?.closest('#fake-host-modal') !== null;
      const z = el ? getComputedStyle(el).zIndex : null;
      return { tag: el?.tagName ?? null, inFcDialog, inHostModal, z };
    },
    { x: Math.round(box.x + box.width / 2), y: Math.round(box.y + box.height / 2) },
  );

  expect(probe.inFcDialog, 'formulon dialog must hit-test above the 9999 host modal').toBe(true);
  expect(probe.inHostModal, 'click must not fall through to the host modal').toBe(false);

  // Clean up the injected modal so subsequent tests in the same worker get a
  // clean DOM.
  await page.keyboard.press('Escape');
  await page.evaluate(() => {
    document.getElementById('fake-host-modal')?.remove();
  });
}
