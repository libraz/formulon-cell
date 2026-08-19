import { test } from '@playwright/test';

import {
  runCopyPasteScenario,
  runCutPasteScenario,
  runRibbonPasteUndoScenario,
} from '../../../tests/e2e-shared/scenarios/clipboard.js';

// Paste routes through `navigator.clipboard.readText()` because the host is a
// non-editable div that never receives a native paste event. WebKit rejects
// that read with `NotAllowedError` and ignores the `clipboard-read` permission
// Playwright grants (cf. the permission override in `playwright.shared.ts`), so
// the paste half of these round-trips cannot run there.
test.skip(({ browserName }) => browserName === 'webkit', 'WebKit denies clipboard reads');

test('C01 (react-demo): Mod+C/V round-trips a cell value', async ({ page }) => {
  await runCopyPasteScenario(page);
});

test('C02 (react-demo): Mod+X clears the source after paste', async ({ page }) => {
  await runCutPasteScenario(page);
});

// Phase 1.5: ribbon Paste now routes through `instance.clipboard.runShortcut`
// so the React wrapper's default hook (which used to silent-fail via
// document.execCommand) actually writes into the sheet.
test('C05 (react-demo): ribbon Paste click writes the clipboard text into the selection', async ({
  page,
}) => {
  await runRibbonPasteUndoScenario(page);
});
