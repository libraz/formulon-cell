import { expect, type Page } from '@playwright/test';

import { SpreadsheetPage } from '../pages/SpreadsheetPage.js';

/** C1 — formulon overlays stay above the host modal even when an ancestor
 *  creates a new stacking context.
 *
 *  Real-world hosts often apply `transform`, `filter`, `will-change`, or
 *  `isolation: isolate` somewhere up the tree (animation libs, motion
 *  components, design-system panels). Every one of these creates a new
 *  stacking context. Once that happens, a descendant's `z-index: 9_999_999`
 *  is **scoped to that subtree**, and a body-level sibling modal with
 *  `z-index: 9999` can win.
 *
 *  We test the four flavours of context-creating ancestor + a body-level
 *  fake host modal. The formulon dialog must still hit-test on top. If it
 *  doesn't, this scenario fails loudly — the renderer needs to teleport
 *  the overlay to `document.body` (or use the popover top-layer API). */
export async function runStackingContextTrapScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await sp.mount();
  await sp.expectNoStub();

  type Probe = { inFcDialog: boolean; inHostModal: boolean };

  const runOne = async (ancestorStyle: string): Promise<Probe> => {
    // Open the format dialog FIRST. After it's mounted we re-parent the host
    // into a stacking-context-creating wrapper — the dialog must survive
    // that. Doing it the other way around blocks our own focusHost click.
    await sp.focusHost();
    await sp.shortcut('1');
    const dialog = page.locator('[class="fc-fmtdlg"]');
    await expect(dialog).toBeVisible({ timeout: 2_000 });

    // Now inject the wrapper + body-level modal. The wrapper is what creates
    // the new stacking context the dialog needs to escape.
    await page.evaluate((style) => {
      const host = document.querySelector('.fc-host') as HTMLElement | null;
      if (!host) throw new Error('no .fc-host');
      const wrap = document.createElement('div');
      wrap.id = 'fc-stacking-wrap';
      wrap.setAttribute('style', style);
      host.parentNode?.insertBefore(wrap, host);
      wrap.appendChild(host);

      const modal = document.createElement('div');
      modal.id = 'fake-host-modal-2';
      modal.setAttribute(
        'style',
        'position: fixed; inset: 0; z-index: 9999; background: rgba(0,0,0,0.25);',
      );
      document.body.appendChild(modal);
    }, ancestorStyle);

    const box = await dialog.boundingBox();
    if (!box) throw new Error('dialog not measured');
    const probe = await page.evaluate(
      ({ x, y }) => {
        const el = document.elementFromPoint(x, y) as HTMLElement | null;
        return {
          inFcDialog: el?.closest('[class~="fc-fmtdlg"]') !== null,
          inHostModal: el?.closest('#fake-host-modal-2') !== null,
        };
      },
      { x: Math.round(box.x + box.width / 2), y: Math.round(box.y + box.height / 2) },
    );

    // Tear down for the next iteration.
    await page.keyboard.press('Escape');
    await expect(dialog).toBeHidden({ timeout: 2_000 });
    await page.evaluate(() => {
      const wrap = document.getElementById('fc-stacking-wrap');
      const host = document.querySelector('.fc-host');
      if (wrap && host && wrap.parentNode) {
        wrap.parentNode.insertBefore(host, wrap);
        wrap.remove();
      }
      document.getElementById('fake-host-modal-2')?.remove();
    });

    return probe;
  };

  // The four flavours of stacking-context creator. Any one of them is enough
  // to bury a descendant overlay if the implementation isn't aware.
  const triggers = [
    'transform: translateZ(0)',
    'filter: blur(0)',
    'isolation: isolate',
    'will-change: transform',
  ];

  for (const style of triggers) {
    const probe = await runOne(style);
    expect(
      probe.inFcDialog,
      `formulon dialog buried by ancestor style "${style}" — overlay needs to escape the stacking context (use document.body / popover top-layer)`,
    ).toBe(true);
    expect(probe.inHostModal, `click fell through to body modal under "${style}"`).toBe(false);
  }
}
