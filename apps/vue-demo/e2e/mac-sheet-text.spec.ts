import {
  runMacBottomGapScenario,
  runMacHeaderLetterSizeScenario,
} from '../../../tests/e2e-shared/scenarios/mac-sheet-text.js';
import { test } from '../../../tests/e2e-shared/test.js';

test('Mac: column-header letters take the cell text size at 100% and 150%', async ({ page }) => {
  await runMacHeaderLetterSizeScenario(page);
});

test('Mac: bottom-aligned Calibri and Aptos Narrow keep a 10 pt bottom gap', async ({ page }) => {
  await runMacBottomGapScenario(page);
});
