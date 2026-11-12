import { test } from '@playwright/test';

import { runConditionalCustomFormatScenario } from '../../../tests/e2e-shared/scenarios/cf-custom-format.js';

for (const platform of ['default', 'mac'] as const) {
  test(`conditional format Custom Format opens Format Cells in the desktop form (${platform} chrome)`, async ({
    page,
  }) => {
    await runConditionalCustomFormatScenario(page, platform);
  });
}
