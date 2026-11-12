import { test } from '@playwright/test';

import { runSortOnCellColorScenario } from '../../../tests/e2e-shared/scenarios/sort-on-color.js';

for (const platform of ['default', 'mac'] as const) {
  test(`Sort dialog Sort On Cell Color brings the filled row to the top (${platform} chrome)`, async ({
    page,
  }) => {
    await runSortOnCellColorScenario(page, platform);
  });
}
