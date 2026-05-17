import { test } from '@playwright/test';

import {
  runCopyPasteScenario,
  runCutPasteScenario,
  runRibbonPasteUndoScenario,
} from '../../../tests/e2e-shared/scenarios/clipboard.js';

test('C01 (vue-demo): Mod+C/V round-trips a cell value', async ({ page }) => {
  await runCopyPasteScenario(page);
});

test('C02 (vue-demo): Mod+X clears the source after paste', async ({ page }) => {
  await runCutPasteScenario(page);
});

// Phase 1.5: ribbon Paste now routes through `instance.clipboard.runShortcut`
// so the Vue wrapper's default hook (which used to silent-fail via
// document.execCommand) actually writes into the sheet.
test('C05 (vue-demo): ribbon Paste click writes the clipboard text into the selection', async ({
  page,
}) => {
  await runRibbonPasteUndoScenario(page);
});
