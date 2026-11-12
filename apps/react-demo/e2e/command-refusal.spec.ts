import { test } from '@playwright/test';

import { runCreateNamesTooLargeScenario } from '../../../tests/e2e-shared/scenarios/create-names-too-large.js';

for (const platform of ['default', 'mac'] as const) {
  for (const locale of ['en', 'ja'] as const) {
    test(`Create from Selection over two whole rows says the selection is too large in ${locale} (${platform} chrome)`, async ({
      page,
    }) => {
      await runCreateNamesTooLargeScenario(page, locale, platform);
    });
  }
}
