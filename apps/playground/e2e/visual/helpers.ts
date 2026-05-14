import { expect, type Page } from '@playwright/test';

export async function mountVisualPage(
  page: Page,
  url = '/',
  opts: { requireRealEngine?: boolean } = {},
): Promise<void> {
  await page.goto(url);
  await page.waitForSelector('.fc-host', { state: 'attached', timeout: 30_000 });
  await page.waitForFunction(
    () => {
      const host = document.querySelector('.fc-host') as HTMLElement | null;
      const state = host?.dataset.fcEngineState;
      return state === 'ready' || state === 'ready-stub';
    },
    { timeout: 30_000 },
  );
  if (opts.requireRealEngine ?? true) {
    const state = await page.locator('.fc-host').first().getAttribute('data-fc-engine-state');
    expect(state, 'engine fell back to stub — check COOP/COEP headers').toBe('ready');
  }
  await page.waitForLoadState('networkidle');
  await page.evaluate(() => document.fonts?.ready);
}
