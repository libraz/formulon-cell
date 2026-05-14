import type { BrowserContext, Page } from '@playwright/test';
import { expect } from '@playwright/test';

import { SpreadsheetPage } from '../pages/SpreadsheetPage.js';

/** C03 — paste an externally-authored HTML table into the grid. WebKit's
 *  clipboard permissions model is restrictive; this scenario runs on
 *  Chromium only — callers must `test.skip(browserName === 'webkit')`
 *  before invoking. */
export async function runExternalHtmlPasteScenario(
  page: Page,
  context: BrowserContext,
): Promise<void> {
  const sp = new SpreadsheetPage(page);
  const consoleErrors = sp.collectConsoleErrors();
  await sp.mount();
  await sp.expectNoStub();

  await context.grantPermissions(['clipboard-read', 'clipboard-write']);

  // Author a tiny spreadsheet-style HTML payload on the system clipboard.
  await page.evaluate(async () => {
    const html =
      '<table><tr><td>ext1</td><td>ext2</td></tr><tr><td>ext3</td><td>ext4</td></tr></table>';
    const item = new ClipboardItem({
      'text/html': new Blob([html], { type: 'text/html' }),
      'text/plain': new Blob(['ext1\text2\next3\text4'], { type: 'text/plain' }),
    });
    await navigator.clipboard.write([item]);
  });

  await sp.focusHost();
  await sp.shortcut('v');
  await page.waitForTimeout(150);

  // Verify the first pasted value reached the active cell via the formula bar.
  expect(await sp.formulaBarValue()).toBe('ext1');
  expect(consoleErrors.read()).toEqual([]);
}
