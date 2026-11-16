import AxeBuilder from '@axe-core/playwright';
import type { Page, TestInfo } from '@playwright/test';
import { expect } from '@playwright/test';

import { SpreadsheetPage } from '../pages/SpreadsheetPage.js';

/** Runs axe over the toolbar regions. Every demo app renders its chrome with
 *  `role="toolbar"`, so a page without one fails rather than auditing nothing. */
export async function runRibbonA11yScenario(page: Page, testInfo: TestInfo): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await sp.mount();
  await sp.expectNoStub();

  await expect(
    page.locator('[role="toolbar"]').first(),
    'the demo renders a [role="toolbar"] region',
  ).toBeAttached();

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
