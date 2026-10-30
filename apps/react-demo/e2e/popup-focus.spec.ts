import { test } from '@playwright/test';
import {
  runMacSubmenuLayoutScenario,
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

test('cell and object context menus keep one highlighted item (react-demo)', async ({ page }) => {
  await runSubmenuHighlightScenario(page);
});

test('Mac submenus fit their labels without inheriting the root width (react-demo)', async ({
  page,
}) => {
  await runMacSubmenuLayoutScenario(page);
});
