import AxeBuilder from '@axe-core/playwright';
import { expect, test } from '@playwright/test';

/**
 * WCAG 2.2 AA audit for the playground. The acceptance bar (per
 * docs/plans/09-roadmap.md MS-E) is **zero AA violations** in the default
 * (light/paper) theme. We restrict the rule set to WCAG 2.0/2.1/2.2 A and AA
 * tags so we don't fail on best-practice or experimental rules.
 *
 * The grid host is a `<canvas>` — axe does not analyze pixels, so cell-level
 * a11y comes from the ARIA mirror layer. Keep this spec focused on chrome
 * (header/title/menus/inputs); per-cell checks belong elsewhere.
 */
test.describe('accessibility (axe-core)', () => {
  test('home page has no WCAG 2.2 AA violations', async ({ page }) => {
    await page.goto('/');
    // Wait for the spreadsheet to mount — chrome is wired up after the
    // engine reports ready.
    await page.waitForSelector('.fc-host', { state: 'attached' });

    const results = await new AxeBuilder({ page })
      .withTags(['wcag2a', 'wcag2aa', 'wcag21a', 'wcag21aa', 'wcag22aa'])
      .analyze();

    if (results.violations.length > 0) {
      // Render a readable summary in the report instead of a raw blob.
      const lines = ['A11y violations:', ''];
      for (const v of results.violations) {
        lines.push(`- [${v.impact}] ${v.id}: ${v.help}`, `  ${v.helpUrl}`);
        for (const node of v.nodes) {
          lines.push(`  · ${node.target.join(' ')}`);
        }
      }
      await test.info().attach('a11y-violations.txt', {
        body: lines.join('\n'),
        contentType: 'text/plain',
      });
    }

    expect(results.violations).toEqual([]);
  });
});
