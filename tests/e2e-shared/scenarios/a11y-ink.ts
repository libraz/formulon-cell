import AxeBuilder from '@axe-core/playwright';
import type { Page, TestInfo } from '@playwright/test';
import { expect } from '@playwright/test';

import { SpreadsheetPage } from '../pages/SpreadsheetPage.js';

/** A04 — full-page axe audit in the ink (dark) theme. The acceptance bar is
 *  zero color-contrast violations: spreadsheet chrome must stay legible
 *  under both themes. Other rule families are also checked so we get the
 *  same baseline coverage A01 gives in paper. */
export async function runInkA11yScenario(page: Page, testInfo: TestInfo): Promise<void> {
  const sp = new SpreadsheetPage(page);
  // Boot directly into ink so the test doesn't depend on a theme toggle.
  await page.goto('/?theme=dark');
  await sp.waitForReady();
  await sp.expectNoStub();

  const results = await new AxeBuilder({ page })
    .withTags(['wcag2a', 'wcag2aa', 'wcag21a', 'wcag21aa', 'wcag22aa'])
    .analyze();

  if (results.violations.length > 0) {
    const lines: string[] = ['Ink-theme a11y violations:', ''];
    for (const v of results.violations) {
      lines.push(`- [${v.impact}] ${v.id}: ${v.help}`, `  ${v.helpUrl}`);
      for (const n of v.nodes) lines.push(`  · ${n.target.join(' ')}`);
    }
    await testInfo.attach('ink-a11y.txt', {
      body: lines.join('\n'),
      contentType: 'text/plain',
    });
  }
  expect(results.violations).toEqual([]);
}
