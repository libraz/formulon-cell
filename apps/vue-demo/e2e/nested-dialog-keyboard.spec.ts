import { test } from '@playwright/test';
import {
  runNameDeleteFocus,
  runNestedNameDialogKeyboard,
} from '../../../tests/e2e-shared/scenarios/nested-dialog-keyboard.js';

test('Excel: nested name dialogs keep keyboard focus and dismiss one layer', async ({ page }) => {
  await runNestedNameDialogKeyboard(page);
});

test('Excel: name deletion restores focus and undo restores the name', async ({ page }) => {
  await runNameDeleteFocus(page);
});
