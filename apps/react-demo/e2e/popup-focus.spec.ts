import { test } from '@playwright/test';
import {
  runPopupFocusScenario,
  runPopupViewportScenario,
  runSubmenuHighlightScenario,
} from '../../../tests/e2e-shared/scenarios/popup-focus.js';

test('popup keyboard focus stays with the active surface (react-demo)', async ({ page }) => {
  await runPopupFocusScenario(page);
});

test('right-click menus fit a short viewport and scroll to focused items (react-demo)', async ({
  page,
}) => {
  await runPopupViewportScenario(page);
});

test('submenu transitions keep one highlighted root item (react-demo)', async ({ page }) => {
  await runSubmenuHighlightScenario(page);
});
