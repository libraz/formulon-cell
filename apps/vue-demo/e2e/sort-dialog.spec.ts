import { runSortCustomListScenario } from '../../../tests/e2e-shared/scenarios/sort-custom-list.js';
import { runSortOnCellColorScenario } from '../../../tests/e2e-shared/scenarios/sort-on-color.js';
import { test } from '../../../tests/e2e-shared/test.js';

for (const platform of ['default', 'mac'] as const) {
  test(`Sort dialog Sort On Cell Color brings the filled row to the top (${platform} chrome)`, async ({
    page,
  }) => {
    await runSortOnCellColorScenario(page, platform);
  });
}

for (const platform of ['default', 'mac'] as const) {
  test(`Sort dialog Custom List orders the rows by a typed list (${platform} chrome)`, async ({
    page,
  }) => {
    await runSortCustomListScenario(page, platform);
  });
}
