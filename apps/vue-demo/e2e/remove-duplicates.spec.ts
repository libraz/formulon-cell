import { runRemoveDuplicatesResultScenario } from '../../../tests/e2e-shared/scenarios/remove-duplicates-result.js';
import { test } from '../../../tests/e2e-shared/test.js';

for (const platform of ['default', 'mac'] as const) {
  for (const locale of ['en', 'ja'] as const) {
    test(`Remove Duplicates reports removed and remaining values in ${locale} (${platform} chrome)`, async ({
      page,
    }) => {
      await runRemoveDuplicatesResultScenario(page, locale, platform);
    });
  }
}
