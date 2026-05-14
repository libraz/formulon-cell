import { test } from '@playwright/test';

import { runPointerEventsCascadeScenario } from '../../../tests/e2e-shared/scenarios/pointer-events-cascade.js';

test('C2 (vue-demo): formulon overlays defend against host pointer-events:none', async ({
  page,
}) => {
  await runPointerEventsCascadeScenario(page);
});
