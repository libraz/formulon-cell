import {
  FULL_CHECK_MATRIX,
  runFullCheckMatrixRow,
} from '../../../tests/e2e-shared/scenarios/full-check-matrix.js';
import { SYSTEM_CLIPBOARD } from '../../../tests/e2e-shared/tags.js';
import { test } from '../../../tests/e2e-shared/test.js';

test.describe('full locale/runtime pairwise matrix (react-demo)', () => {
  test.describe.configure({ mode: 'parallel' });
  for (const row of FULL_CHECK_MATRIX) {
    test(`${row.locale}/${row.platform}/${row.theme}/${row.width}/${row.flow}`, {
      tag: row.flow === 'clipboard' ? [SYSTEM_CLIPBOARD] : [],
    }, async ({ page }) => {
      test.setTimeout(180_000);
      await runFullCheckMatrixRow(page, row);
    });
  }
});
