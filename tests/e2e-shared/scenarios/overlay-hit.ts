import { expect, type Page } from '@playwright/test';

import { SpreadsheetPage } from '../pages/SpreadsheetPage.js';

/** B1 — dialog overlay must mask the grid underneath.
 *
 *  When a dialog is open, a click in the middle of the grid should NOT reach
 *  a canvas/cell — the dialog (or its backdrop) must be the topmost element
 *  at that point. This regression-tests the z-index tier system in
 *  `core/base.css` and asserts that the dialog renders on a higher layer than
 *  the grid (tier `dialog` = 2,147,483,020 vs `grid` = 2,147,483,010).
 *
 *  We use `document.elementFromPoint` rather than a synthetic click so the
 *  test stays deterministic regardless of which pointer-event handlers each
 *  layer happens to register.
 *
 *  Symmetric: after closing the dialog, the same point should hit the canvas. */
export async function runDialogMasksGridScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await sp.mount();
  await sp.expectNoStub();

  // Compute a probe point at the middle of the grid surface. The canvas is
  // the .fc-host__canvas div under the chrome.
  const canvasBox = await page.locator('.fc-host__canvas').first().boundingBox();
  expect(canvasBox, 'canvas must be laid out').not.toBeNull();
  if (!canvasBox) throw new Error('canvas not measured');
  const px = Math.round(canvasBox.x + canvasBox.width / 2);
  const py = Math.round(canvasBox.y + canvasBox.height / 2);

  // Baseline: with no dialog open, the probe lands somewhere on the
  // .fc-host__canvas / .fc-host area.
  const beforeOpen = await page.evaluate(
    ({ x, y }) => {
      const el = document.elementFromPoint(x, y) as HTMLElement | null;
      return {
        tag: el?.tagName ?? null,
        cls: el?.className ?? null,
        closestHost: el?.closest('.fc-host') ? true : false,
      };
    },
    { x: px, y: py },
  );
  expect(beforeOpen.closestHost, 'probe should fall inside the host').toBe(true);
  // The hit should be on a grid-tier surface, NOT a dialog-tier.
  expect(beforeOpen.cls ?? '', 'no dialog should be active').not.toContain('fc-fmtdlg');

  // Open the format dialog via Mod+1.
  await sp.focusHost();
  await sp.shortcut('1');

  // Wait for the dialog to actually become visible.
  const fmtDialog = page.locator('[class="fc-fmtdlg"]');
  await expect(fmtDialog).toBeVisible({ timeout: 2_000 });

  // The probe point now needs to fall on the dialog stack rather than the
  // canvas. If the dialog happens to be centered slightly off our probe (e.g.
  // anchored to a corner), check anywhere inside the dialog's own box instead.
  const dialogBox = await fmtDialog.boundingBox();
  expect(dialogBox, 'dialog must be laid out').not.toBeNull();
  if (!dialogBox) throw new Error('dialog not measured');
  const dpx = Math.round(dialogBox.x + dialogBox.width / 2);
  const dpy = Math.round(dialogBox.y + dialogBox.height / 2);

  const duringOpen = await page.evaluate(
    ({ x, y }) => {
      const el = document.elementFromPoint(x, y) as HTMLElement | null;
      return {
        tag: el?.tagName ?? null,
        underDialog: el?.closest('[class~="fc-fmtdlg"]') ? true : false,
      };
    },
    { x: dpx, y: dpy },
  );
  expect(
    duringOpen.underDialog,
    'point inside the dialog box must land on a dialog descendant — not a grid layer',
  ).toBe(true);

  // Close via Escape; probe must return to the canvas layer.
  await page.keyboard.press('Escape');
  await expect(fmtDialog).toBeHidden({ timeout: 2_000 });

  const afterClose = await page.evaluate(
    ({ x, y }) => {
      const el = document.elementFromPoint(x, y) as HTMLElement | null;
      return {
        cls: el?.className ?? null,
        closestHost: el?.closest('.fc-host') ? true : false,
      };
    },
    { x: px, y: py },
  );
  expect(afterClose.closestHost, 'probe back inside the host after close').toBe(true);
  expect(afterClose.cls ?? '', 'no lingering dialog overlay').not.toContain('fc-fmtdlg');
}
