import { runPointerEventsCascadeScenario } from '../../../tests/e2e-shared/scenarios/pointer-events-cascade.js';
import { test } from '../../../tests/e2e-shared/test.js';

test('C2 (react-demo): formulon overlays defend against host pointer-events:none', async ({
  page,
}) => {
  await runPointerEventsCascadeScenario(page);
});
