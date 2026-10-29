import { test } from '@playwright/test';
import {
  runDirectionFocusFeedbackScenario,
  runMiniToolbarFeedbackScenario,
} from '../../../tests/e2e-shared/scenarios/visual-feedback.js';

test('direction choices show keyboard focus in every theme (react-demo)', async ({ page }) => {
  await runDirectionFocusFeedbackScenario(page);
});

test('mini toolbar shows current formatting and keyboard focus (react-demo)', async ({ page }) => {
  await runMiniToolbarFeedbackScenario(page);
});
