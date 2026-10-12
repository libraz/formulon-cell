import { test } from '@playwright/test';
import { runWrappedText } from '../../../tests/e2e-shared/scenarios/wrapped-text.js';

test('real canvas wrapping and AutoFit agree across seven locales and zoom', async ({ page }) => {
  test.setTimeout(120_000);
  await runWrappedText(page);
});
