import { test } from '@playwright/test';
import {
  runMacAutoFilterByColorScenario,
  runMacSortByColorScenario,
} from '../../../tests/e2e-shared/scenarios/autofilter-by-color.js';
import { runMacAutoFilterDateScenario } from '../../../tests/e2e-shared/scenarios/autofilter-date-conditions.js';
import { runAutoFilterDateTreeScenario } from '../../../tests/e2e-shared/scenarios/autofilter-date-tree.js';
import { runMacAutoFilterOrPairScenario } from '../../../tests/e2e-shared/scenarios/autofilter-two-conditions.js';

for (const locale of ['en', 'de', 'ja'] as const) {
  test(`Mac date column filter reads a ${locale} Before date and applies a period at once`, async ({
    page,
  }) => {
    await runMacAutoFilterDateScenario(page, locale);
  });
}

test('Mac column filter joins a second comparison with Or', async ({ page }) => {
  await runMacAutoFilterOrPairScenario(page);
});

for (const platform of ['default', 'mac'] as const) {
  test(`date column filter groups its checklist by year and month (${platform} chrome)`, async ({
    page,
  }) => {
    await runAutoFilterDateTreeScenario(page, platform);
  });
}

test('Mac column filter By color keeps only the rows with the picked fill', async ({ page }) => {
  await runMacAutoFilterByColorScenario(page);
});

test('Mac column sort By color brings the rows with the picked fill to the top', async ({
  page,
}) => {
  await runMacSortByColorScenario(page);
});
