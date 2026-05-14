import AxeBuilder from '@axe-core/playwright';
import type { Page, TestInfo } from '@playwright/test';
import { expect } from '@playwright/test';

import { SpreadsheetPage } from '../pages/SpreadsheetPage.js';

/** A01 — WCAG 2.2 AA scan of the mounted app. The canvas grid is not analyzed
 *  (axe can't see pixels); per-cell a11y comes from the ARIA mirror layer. */
export async function runA11yBaselineScenario(page: Page, testInfo: TestInfo): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await sp.mount();
  await sp.expectNoStub();

  const results = await new AxeBuilder({ page })
    .withTags(['wcag2a', 'wcag2aa', 'wcag21a', 'wcag21aa', 'wcag22aa'])
    .analyze();

  if (results.violations.length > 0) {
    const lines: string[] = ['A11y violations:', ''];
    for (const v of results.violations) {
      lines.push(`- [${v.impact}] ${v.id}: ${v.help}`, `  ${v.helpUrl}`);
      for (const node of v.nodes) {
        lines.push(`  · ${node.target.join(' ')}`);
      }
    }
    await testInfo.attach('a11y-violations.txt', {
      body: lines.join('\n'),
      contentType: 'text/plain',
    });
  }

  expect(results.violations).toEqual([]);
}
