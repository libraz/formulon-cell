import {
  runCopyPasteScenario,
  runCutPasteScenario,
  runHtmlFontFamilyEscapeScenario,
  runRibbonPasteUndoScenario,
} from '../../../tests/e2e-shared/scenarios/clipboard.js';
import { SYSTEM_CLIPBOARD } from '../../../tests/e2e-shared/tags.js';
import { test } from '../../../tests/e2e-shared/test.js';

// Paste routes through `navigator.clipboard.readText()` because the host is a
// non-editable div that never receives a native paste event. WebKit rejects
// that read with `NotAllowedError` and ignores the `clipboard-read` permission
// Playwright grants, so paste there falls back to the workbook's own copy: an
// in-workbook round-trip runs, a paste of external clipboard text cannot.

test('C01 (react-demo): Mod+C/V round-trips a cell value', { tag: SYSTEM_CLIPBOARD }, async ({
  page,
}) => {
  await runCopyPasteScenario(page);
});

test('C02 (react-demo): Mod+X clears the source after paste', { tag: SYSTEM_CLIPBOARD }, async ({
  page,
}) => {
  await runCutPasteScenario(page);
});

// Phase 1.5: ribbon Paste now routes through `instance.clipboard.runShortcut`
// so the React wrapper's default hook (which used to silent-fail via
// document.execCommand) actually writes into the sheet.
test('C05 (react-demo): ribbon Paste click writes the clipboard text into the selection', {
  tag: SYSTEM_CLIPBOARD,
}, async ({ page, browserName }) => {
  test.skip(browserName === 'webkit', 'WebKit denies the clipboard read an external paste needs');
  await runRibbonPasteUndoScenario(page);
});

test('C06 (react-demo): a copied font family keeps an ampersand and a quote through the CSS parser', async ({
  page,
}) => {
  await runHtmlFontFamilyEscapeScenario(page);
});
