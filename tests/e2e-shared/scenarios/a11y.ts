import AxeBuilder from '@axe-core/playwright';
import type { Page, TestInfo } from '@playwright/test';
import { expect } from '@playwright/test';

import { selectRibbonTab } from '../pages/ribbon-tabs.js';
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

/** WCAG 2.2 AA scan of the whole app with each main ribbon tab shown in turn, so a control that
 *  only renders on a non-Home tab is audited too. */
export async function runA11yEveryTabScenario(
  page: Page,
  testInfo: TestInfo,
  platform: 'default' | 'mac',
): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await sp.mount({ platform });
  await page.setViewportSize({ width: 1440, height: 900 });
  await sp.expectNoStub();
  const tabs = await page
    .locator('[data-ribbon-tab]:visible')
    .evaluateAll((els) =>
      els
        .map((el) => (el as HTMLElement).dataset.ribbonTab ?? '')
        .filter((id) => id !== '' && id !== 'file'),
    );
  expect(tabs.length, 'main ribbon tabs').toBeGreaterThan(5);

  const lines: string[] = [];
  for (const tab of tabs) {
    await selectRibbonTab(page, tab);
    await expect(page.locator(`[data-ribbon-panel="${tab}"]`)).toBeVisible();
    const results = await new AxeBuilder({ page })
      .withTags(['wcag2a', 'wcag2aa', 'wcag21a', 'wcag21aa', 'wcag22aa'])
      .analyze();
    for (const v of results.violations) {
      for (const node of v.nodes)
        lines.push(`${tab}: [${v.impact}] ${v.id} ${node.target.join(' ')}`);
    }
  }
  if (lines.length > 0) {
    await testInfo.attach('a11y-every-tab.txt', {
      body: lines.join('\n'),
      contentType: 'text/plain',
    });
  }
  expect(lines).toEqual([]);
}
