import AxeBuilder from '@axe-core/playwright';
import type { Page, TestInfo } from '@playwright/test';
import { expect } from '@playwright/test';

import { SpreadsheetPage } from '../pages/SpreadsheetPage.js';

/** A02 — open the format dialog (Mod+1) and run axe against the focused state.
 *  We assert zero new violations introduced by the dialog overlay. */
export async function runFormatDialogA11yScenario(page: Page, testInfo: TestInfo): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await sp.mount();
  await sp.expectNoStub();

  await sp.focusHost();
  await sp.shortcut('1');
  // Ensure the dialog is up.
  await expect(page.locator('[class="fc-fmtdlg"]')).toBeVisible({ timeout: 2000 });

  const results = await new AxeBuilder({ page })
    .withTags(['wcag2a', 'wcag2aa', 'wcag21a', 'wcag21aa', 'wcag22aa'])
    .include('[class="fc-fmtdlg"]')
    .analyze();

  if (results.violations.length > 0) {
    const lines: string[] = ['Format dialog a11y violations:', ''];
    for (const v of results.violations) {
      lines.push(`- [${v.impact}] ${v.id}: ${v.help}`, `  ${v.helpUrl}`);
      for (const node of v.nodes) lines.push(`  · ${node.target.join(' ')}`);
    }
    await testInfo.attach('format-dialog-a11y.txt', {
      body: lines.join('\n'),
      contentType: 'text/plain',
    });
  }
  expect(results.violations).toEqual([]);
}
