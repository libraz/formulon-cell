import { test } from '@playwright/test';
import { runNestedNameDialogKeyboard } from '../../../tests/e2e-shared/scenarios/nested-dialog-keyboard.js';

test('Excel: nested name dialogs keep keyboard focus and dismiss one layer', async ({ page }) => {
  await runNestedNameDialogKeyboard(page);
});
