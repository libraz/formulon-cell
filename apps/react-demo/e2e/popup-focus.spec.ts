import { test } from '@playwright/test';
import {
  runPopupFocusScenario,
  runPopupViewportScenario,
} from '../../../tests/e2e-shared/scenarios/popup-focus.js';

test('popup keyboard focus stays with the active surface (react-demo)', async ({ page }) => {
  await runPopupFocusScenario(page);
});

test('right-click menus fit a short viewport and scroll to focused items (react-demo)', async ({
  page,
}) => {
  await runPopupViewportScenario(page);
});
