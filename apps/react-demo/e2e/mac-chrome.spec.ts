import { test } from '@playwright/test';

import {
  runMacChromeScenario,
  runMacRibbonDisplayPlacementScenario,
  runMacRibbonDisplayScenario,
  runMacRibbonLayoutScenario,
} from '../../../tests/e2e-shared/scenarios/mac-chrome.js';

test('Mac Excel chrome (react-demo): compact tabs, geometry, and Home backstage entry', async ({
  page,
}) => {
  await runMacChromeScenario(page);
});

test('Mac ribbon labels and scroll endpoints (react-demo)', async ({ page }) => {
  test.setTimeout(120_000);
  await runMacRibbonLayoutScenario(page);
});

test('Mac ribbon display modes and tab rail (react-demo)', async ({ page }) => {
  test.setTimeout(120_000);
  await runMacRibbonDisplayScenario(page);
});

test('Mac ribbon display-menu placement in narrow transformed hosts (react-demo)', async ({
  page,
}) => {
  test.setTimeout(120_000);
  await runMacRibbonDisplayPlacementScenario(page);
});
