import { test } from '@playwright/test';
import {
  runMacHeaderSelectionToggle,
  runMacMergedSelectionToggle,
  runMacSelectionClearWithHole,
  runMacSelectionDrag,
  runMacSelectionInputWithHole,
  runMacSelectionToggle,
} from '../../../tests/e2e-shared/scenarios/mac-multi-selection.js';

test('Mac: Command click deselects and reselects cells without changing their values', async ({
  page,
}) => {
  await runMacSelectionToggle(page);
});
test('Mac: Command drag adds and removes ranges from a stable starting selection', async ({
  page,
}) => {
  await runMacSelectionDrag(page);
});
test('Mac: Command click on a merged body toggles the entire merge', async ({ page }) => {
  await runMacMergedSelectionToggle(page);
});

test('Mac: Command Return writes only selected cells and undoes the batch', async ({ page }) => {
  await runMacSelectionInputWithHole(page);
});

test('Mac: Command header gestures toggle full bands and restore a shrinking drag', async ({
  page,
}) => {
  await runMacHeaderSelectionToggle(page);
});

test('Mac: Delete and Clear Contents preserve deselected holes with one undo and redo', async ({
  page,
}) => {
  await runMacSelectionClearWithHole(page);
});
