import { expect, type Page } from '@playwright/test';

/** Whether the page may read the system clipboard. WebKit rejects `navigator.clipboard.readText()`
 *  with `NotAllowedError` and ignores the `clipboard-read` permission Playwright grants. */
export const canReadSystemClipboard = (page: Page): boolean =>
  page.context().browser()?.browserType().name() !== 'webkit';

/** Waits until the system clipboard holds `text`, so a paste never races the async copy. Where the
 *  clipboard cannot be read, paste falls back to the workbook's own copy, so there is no race to
 *  wait out and nothing to check. */
export const expectSystemClipboardText = async (page: Page, text: string): Promise<void> => {
  if (!canReadSystemClipboard(page)) return;
  await expect.poll(() => page.evaluate(() => navigator.clipboard.readText())).toBe(text);
};
