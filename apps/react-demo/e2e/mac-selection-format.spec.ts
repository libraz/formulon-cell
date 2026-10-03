import { test } from '@playwright/test';
import {
  runMacSelectionAccentStyles,
  runMacSelectionBordersWithHole,
  runMacSelectionCellStyleWithHole,
  runMacSelectionClearAllWithHole,
  runMacSelectionClearFormatsWithHole,
  runMacSelectionContextClearWithHole,
  runMacSelectionFormatDialogWithHole,
  runMacSelectionFormatRepeat,
  runMacSelectionFormatWithHole,
} from '../../../tests/e2e-shared/scenarios/mac-selection-format.js';

test('Mac: Japanese Accent cell styles expose all chips and repeat with one undo', async ({
  page,
}) => {
  await runMacSelectionAccentStyles(page);
});

test('Mac: formatting a mixed selection preserves holes with one undo and redo', async ({
  page,
}) => {
  await runMacSelectionFormatWithHole(page);
});

test('Mac: Clear Formats covers additional areas and preserves deselected cells', async ({
  page,
}) => {
  await runMacSelectionClearFormatsWithHole(page);
});

test('Mac: F4 formats the current selection union and undoes it in one step', async ({ page }) => {
  await runMacSelectionFormatRepeat(page);
});

test('Mac: borders cover additional areas and preserve holes with one undo and F4', async ({
  page,
}) => {
  await runMacSelectionBordersWithHole(page);
});

test('Mac: Clear All preserves holes and restores all selected values and formats with one undo', async ({
  page,
}) => {
  await runMacSelectionClearAllWithHole(page);
});

test('Mac: Format Cells preserves mixed fields and holes, with one undo and F4 over all areas', async ({
  page,
}) => {
  await runMacSelectionFormatDialogWithHole(page);
});

test('Mac: Cell Styles replace font and fill across the selection union, with one undo and F4', async ({
  page,
}) => {
  await runMacSelectionCellStyleWithHole(page);
});

test('Mac: context-menu Clear preserves holes and clears all areas with one undo and redo', async ({
  page,
}) => {
  await runMacSelectionContextClearWithHole(page);
});
