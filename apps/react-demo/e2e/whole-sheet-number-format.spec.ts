import { test } from '@playwright/test';
import { runWholeSheetNumberFormat } from '../../../tests/e2e-shared/scenarios/whole-sheet-number-format.js';

for (const platform of ['default', 'mac'] as const) {
  test(`whole-sheet number format preserves text and shows mixed font: ${platform}`, async ({
    page,
  }) => {
    await runWholeSheetNumberFormat(page, platform);
  });
}
