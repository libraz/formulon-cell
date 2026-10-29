import { test } from '@playwright/test';
import {
  runSelectEscapeLayers,
  runSelectHomeEndAndDeparture,
} from '../../../tests/e2e-shared/scenarios/custom-select-keyboard.js';

test('Excel: selector Home and End commit the highlighted option and Tab closes it', async ({
  page,
}) => {
  await runSelectHomeEndAndDeparture(page);
});
test('Excel: Escape closes the selector before its parent dialog', async ({ page }) => {
  await runSelectEscapeLayers(page);
});
