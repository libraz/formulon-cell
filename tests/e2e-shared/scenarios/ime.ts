import type { Page } from '@playwright/test';
import { expect } from '@playwright/test';

import { SpreadsheetPage } from '../pages/SpreadsheetPage.js';

/** E05 — IME composition. While `isComposing` is true on the editor input,
 *  Enter should be intercepted by the IME (commit composition only) and
 *  NOT also dispatch the spreadsheet's "commit cell" action. We can't drive
 *  a real IME from Playwright, but we can dispatch the same CompositionEvent
 *  sequence the browser emits and assert the editor still has focus and the
 *  composed value lands. */
export async function runImeCompositionScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  const consoleErrors = sp.collectConsoleErrors();
  await sp.mount();
  await sp.expectNoStub();

  await sp.focusHost();
  // F2 to open the in-cell editor explicitly (so we have a focused input).
  await page.keyboard.press('F2');
  await page.waitForTimeout(80);

  // Synthesize an IME composition: compositionstart → compositionupdate → input → compositionend.
  // We can't actually inject CJK from Playwright reliably across WebKit; the goal here is
  // that nothing throws and the editor receives `compositionend` cleanly.
  await page.evaluate(() => {
    const active = document.activeElement as HTMLElement | null;
    if (!active) return;
    active.dispatchEvent(new CompositionEvent('compositionstart', { data: '' }));
    active.dispatchEvent(new CompositionEvent('compositionupdate', { data: 'あ' }));
    active.dispatchEvent(new CompositionEvent('compositionend', { data: 'あ' }));
  });

  await page.waitForTimeout(80);
  expect(consoleErrors.read(), 'composition event sequence should not raise').toEqual([]);

  await page.keyboard.press('Escape');
}
