import { test } from '@playwright/test';

import { runSaveXlsxScenario } from '../../../tests/e2e-shared/scenarios/save-xlsx.js';

test('D01 (playground): instance.workbook.save() produces a valid xlsx (ZIP PK header)', async ({
  page,
}) => {
  await runSaveXlsxScenario(page);
});
