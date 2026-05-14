import { expect, type Page } from '@playwright/test';

import { SpreadsheetPage } from '../pages/SpreadsheetPage.js';

/** C2 — formulon overlays remain clickable even when a host ancestor sets
 *  `pointer-events: none`.
 *
 *  Some design systems / animation libraries apply `pointer-events: none` on
 *  whole subtrees while a transition runs, or as a global reset. Because
 *  `pointer-events` cascades, every descendant becomes click-inert unless the
 *  descendant re-enables it explicitly. The formulon dialog must defend
 *  itself with `pointer-events: auto` on its overlay root — otherwise users
 *  see the dialog but cannot interact with it.
 *
 *  This test will FAIL until that defensive rule exists. The failure is the
 *  point: it surfaces the bug for the renderer to fix. */
export async function runPointerEventsCascadeScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await sp.mount();
  await sp.expectNoStub();

  // Focus + open the dialog FIRST. The host wrap disables pointer events,
  // which would block our own focusHost click.
  await sp.focusHost();
  await sp.shortcut('1');
  const dialog = page.locator('[class="fc-fmtdlg"]');
  await expect(dialog).toBeVisible({ timeout: 2_000 });

  // Wrap the host in a div that disables pointer events on its entire
  // subtree. Real hosts hit this via transition libraries, modal trap
  // libraries, or accidental global resets.
  await page.evaluate(() => {
    const host = document.querySelector('.fc-host') as HTMLElement | null;
    if (!host) throw new Error('no .fc-host');
    const wrap = document.createElement('div');
    wrap.id = 'fc-cascade-wrap';
    wrap.style.cssText = 'pointer-events: none;';
    host.parentNode?.insertBefore(wrap, host);
    wrap.appendChild(host);
  });

  // Verify the dialog overlay declares pointer-events: auto (or its parent
  // tree somehow re-establishes interactivity). We check the *computed*
  // value, which incorporates the ancestor cascade.
  const computedPe = await page.evaluate(() => {
    const el = document.querySelector('[class="fc-fmtdlg"]') as HTMLElement | null;
    return el ? getComputedStyle(el).pointerEvents : null;
  });
  expect(
    computedPe,
    'formulon dialog overlay must declare pointer-events: auto to survive a host pointer-events:none cascade',
  ).toBe('auto');

  // The cancel button inside the dialog must also remain hit-testable.
  // We look up its element and read computed pointer-events.
  const cancelPe = await page.evaluate(() => {
    // Cancel is the first non-OK button in the dialog footer. Find any
    // button — we just need to know whether *some* button is clickable.
    const btns = Array.from(
      document.querySelectorAll<HTMLButtonElement>('[class="fc-fmtdlg"] button'),
    );
    const first = btns[0];
    return first ? getComputedStyle(first).pointerEvents : null;
  });
  expect(
    cancelPe,
    'buttons inside the formulon dialog must inherit/re-enable pointer-events: auto',
  ).toBe('auto');

  // Diagnostic: confirm the dialog still exists and has layout after the wrap.
  const diag = await page.evaluate(() => {
    const el = document.querySelector('[class="fc-fmtdlg"]') as HTMLElement | null;
    if (!el) return { exists: false } as const;
    const rect = el.getBoundingClientRect();
    return {
      exists: true,
      hidden: el.hidden,
      rect: { x: rect.x, y: rect.y, w: rect.width, h: rect.height },
      parentTag: el.parentElement?.tagName,
      parentClass: el.parentElement?.className,
    } as const;
  });
  // Make the diagnostic part of the failure message so we know what state we're in.
  if (!diag.exists) throw new Error('dialog disappeared after host wrap');

  // Functional probe: simulate a click on the dialog backdrop or a button
  // and assert it lands (elementFromPoint returns a dialog descendant rather
  // than the inert wrapper).
  const box = await dialog.boundingBox();
  if (!box) {
    throw new Error(`dialog not measured; diag=${JSON.stringify(diag)}`);
  }
  const probe = await page.evaluate(
    ({ x, y }) => {
      const el = document.elementFromPoint(x, y) as HTMLElement | null;
      // Track the full ancestor chain + pointer-events at each level so a
      // failure tells us which ancestor is intercepting (or being skipped).
      const chain: { tag: string; cls: string; pe: string }[] = [];
      let cur: HTMLElement | null = el;
      while (cur) {
        chain.push({
          tag: cur.tagName,
          cls: cur.className?.toString?.() ?? '',
          pe: getComputedStyle(cur).pointerEvents,
        });
        cur = cur.parentElement;
      }
      return {
        tag: el?.tagName ?? null,
        cls: el?.className?.toString?.() ?? null,
        inDialog: el?.closest('[class~="fc-fmtdlg"]') !== null,
        pe: el ? getComputedStyle(el).pointerEvents : null,
        chain,
      };
    },
    { x: Math.round(box.x + box.width / 2), y: Math.round(box.y + box.height / 2) },
  );
  expect(
    probe.inDialog,
    `point inside the dialog must hit a dialog descendant (got ${probe.tag} class="${probe.cls}", pointer-events=${probe.pe}, chain=${JSON.stringify(probe.chain)}, diag=${JSON.stringify(diag)}, box=${JSON.stringify(box)})`,
  ).toBe(true);

  // Tear down. (Best-effort: if Escape doesn't fire because pointer-events
  // are gated, we still proceed with the cleanup so subsequent tests in the
  // same worker get a clean DOM.)
  await page.keyboard.press('Escape');
  await page.evaluate(() => {
    const wrap = document.getElementById('fc-cascade-wrap');
    const host = document.querySelector('.fc-host');
    if (wrap && host && wrap.parentNode) {
      wrap.parentNode.insertBefore(host, wrap);
      wrap.remove();
    }
  });
}
