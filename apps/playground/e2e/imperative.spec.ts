import { expect, test } from '@playwright/test';

import { SpreadsheetPage } from '../../../tests/e2e-shared/pages/SpreadsheetPage.js';

/** The playground exposes the mounted instance on `window.__fcInst` for
 *  debugging. That makes it a useful base case for imperative-API behavior
 *  the wrappers also expose — when something diverges between vanilla and
 *  wrapper, this spec is the no-wrapper reference.
 */
test('imperative (playground): instance.i18n.setLocale fires localeChange and updates i18n.locale', async ({
  page,
}) => {
  const sp = new SpreadsheetPage(page);
  await sp.mount();
  await sp.expectNoStub();

  type SpreadsheetGlobal = {
    __fcInst?: {
      i18n: {
        readonly locale: string;
        setLocale(l: string): void;
        subscribe(fn: () => void): () => void;
      };
    };
  };

  // Subscribe before flipping so we record exactly one notification.
  const result = await page.evaluate(() => {
    const w = window as unknown as SpreadsheetGlobal;
    const inst = w.__fcInst;
    if (!inst) return { ok: false as const, reason: 'no __fcInst' };
    const before = inst.i18n.locale;
    let notifications = 0;
    const unsub = inst.i18n.subscribe(() => {
      notifications += 1;
    });
    inst.i18n.setLocale(before === 'ja' ? 'en' : 'ja');
    const after = inst.i18n.locale;
    unsub();
    return { ok: true as const, before, after, notifications };
  });

  expect(result.ok).toBe(true);
  if (!result.ok) return; // narrow for TS; the assertion above already failed
  expect(result.before).not.toBe(result.after);
  expect(result.notifications).toBeGreaterThanOrEqual(1);
});
