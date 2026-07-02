import { test } from '@playwright/test';

import { runPrintPdfSmokeScenario } from '../../../tests/e2e-shared/scenarios/print-pdf.js';

test('Print PDF smoke (react-demo): preview HTML can render to a browser PDF', async ({
  browserName,
  page,
}) => {
  test.skip(browserName !== 'chromium', 'page.pdf() is Chromium-only.');
  await runPrintPdfSmokeScenario(page);
});
