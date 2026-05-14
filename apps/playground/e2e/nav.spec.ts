import { test } from '@playwright/test';
import { runArrowSequenceScenario } from '../../../tests/e2e-shared/scenarios/arrow-sequence.js';
import {
  runArrowAndTabNavScenario,
  runF2EscapeScenario,
} from '../../../tests/e2e-shared/scenarios/nav.js';

test('E03 (playground): Tab / Shift+Tab / Arrow navigate the grid', async ({ page }) => {
  await runArrowAndTabNavScenario(page);
});

test('E04 (playground): F2 → ESC discards the in-progress edit', async ({ page }) => {
  await runF2EscapeScenario(page);
});

test('Arrow sequence (playground): Arrow keys + Enter walk the grid', async ({ page }) => {
  await runArrowSequenceScenario(page);
});
