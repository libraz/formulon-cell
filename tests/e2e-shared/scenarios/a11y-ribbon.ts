import AxeBuilder from '@axe-core/playwright';
import type { Page, TestInfo } from '@playwright/test';
import { expect } from '@playwright/test';

import { SpreadsheetPage } from '../pages/SpreadsheetPage.js';

/** A03 — run axe scoped to the toolbar / ribbon region (where it exists).
 *  Each demo app renders its own ribbon with `role="toolbar"` somewhere on
 *  the page; the scenario picks the first such region and audits it. */
export async function runRibbonA11yScenario(page: Page, testInfo: TestInfo): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await sp.mount();
  await sp.expectNoStub();

  const exists = (await page.locator('[role="toolbar"]').count()) > 0;
  if (!exists) {
    // No ribbon/toolbar mounted in this app's chrome — nothing to audit.
    test_skipReason(testInfo, 'no [role="toolbar"] found in this demo');
    return;
  }

  const results = await new AxeBuilder({ page })
    .withTags(['wcag2a', 'wcag2aa', 'wcag21a', 'wcag21aa', 'wcag22aa'])
    .include('[role="toolbar"]')
    .analyze();

  if (results.violations.length > 0) {
    const lines: string[] = ['Ribbon a11y violations:', ''];
    for (const v of results.violations) {
      lines.push(`- [${v.impact}] ${v.id}: ${v.help}`, `  ${v.helpUrl}`);
      for (const n of v.nodes) lines.push(`  · ${n.target.join(' ')}`);
    }
    await testInfo.attach('ribbon-a11y.txt', {
      body: lines.join('\n'),
      contentType: 'text/plain',
    });
  }
  expect(results.violations).toEqual([]);
}

function test_skipReason(info: TestInfo, reason: string): void {
  info.annotations.push({ type: 'skip-reason', description: reason });
}
