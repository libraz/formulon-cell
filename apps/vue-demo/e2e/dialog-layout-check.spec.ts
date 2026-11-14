import { runDialogLayoutCheckScenario } from '../../../tests/e2e-shared/scenarios/dialog-layout-check.js';
import { test } from '../../../tests/e2e-shared/test.js';

// The check's own fixture draws layout faults on purpose and asserts on them itself.
test.use({ dialogLayoutFaults: 'expected' });

test('dialog layout check passes self-clipped text and reports overlapping captions', async ({
  page,
}) => {
  await runDialogLayoutCheckScenario(page);
});
