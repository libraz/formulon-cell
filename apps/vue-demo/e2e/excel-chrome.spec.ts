import { test } from '@playwright/test';

import {
  runExcelChromeBackstageSearchScenario,
  runExcelChromeConditionalFormattingMenuScenario,
  runExcelChromeCreateTableDialogScenario,
  runExcelChromeHomeDenseRibbonScenario,
  runExcelChromeTableStyleGalleryScenario,
} from '../../../tests/e2e-shared/scenarios/excel-chrome.js';

test('Excel chrome (vue-demo): File backstage, F6 landmarks, and Alt+Q search are wired', async ({
  page,
}) => {
  await runExcelChromeBackstageSearchScenario(page);
});

test('Excel chrome (vue-demo): Format as Table gallery is scrollable and menu-backed', async ({
  page,
}) => {
  await runExcelChromeTableStyleGalleryScenario(page);
});

test('Excel chrome (vue-demo): Insert Table primary click opens Create Table dialog', async ({
  page,
}) => {
  await runExcelChromeCreateTableDialogScenario(page);
});

test('Excel chrome (vue-demo): Conditional Formatting opens Excel-style menu', async ({ page }) => {
  await runExcelChromeConditionalFormattingMenuScenario(page);
});

test('Excel chrome (vue-demo): Home dense ribbon groups do not overflow', async ({ page }) => {
  await runExcelChromeHomeDenseRibbonScenario(page);
});
