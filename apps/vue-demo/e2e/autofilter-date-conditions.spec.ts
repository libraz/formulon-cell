import { test } from '@playwright/test';

import { runMacAutoFilterDateScenario } from '../../../tests/e2e-shared/scenarios/autofilter-date-conditions.js';

for (const locale of ['en', 'de', 'ja'] as const) {
  test(`Mac date column filter reads a ${locale} Before date and applies a period at once`, async ({
    page,
  }) => {
    await runMacAutoFilterDateScenario(page, locale);
  });
}
