import { test } from '@playwright/test';
import { runMacAutoFilterDateScenario } from '../../../tests/e2e-shared/scenarios/autofilter-date-conditions.js';
import { runAutoFilterDateTreeScenario } from '../../../tests/e2e-shared/scenarios/autofilter-date-tree.js';
import { runMacAutoFilterOrPairScenario } from '../../../tests/e2e-shared/scenarios/autofilter-two-conditions.js';
import {
  runFilterShortcutScenario,
  runProtectedSheetFilterAndBackgroundScenario,
  runProtectedSheetFilterButtonScenario,
} from '../../../tests/e2e-shared/scenarios/filter-protection.js';

test('Ctrl+Shift+L toggles AutoFilter over the current region', async ({ page }) => {
  await runFilterShortcutScenario(page);
});

test('a protected sheet disables Background and refuses Ctrl+Shift+L', async ({ page }) => {
  await runProtectedSheetFilterAndBackgroundScenario(page);
});

test('a protected sheet leaves a filter button inert without Use AutoFilter', async ({ page }) => {
  await runProtectedSheetFilterButtonScenario(page);
});

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
