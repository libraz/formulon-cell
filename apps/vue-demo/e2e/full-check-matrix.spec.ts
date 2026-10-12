import { test } from '@playwright/test';

import {
  FULL_CHECK_MATRIX,
  runFullCheckMatrixRow,
} from '../../../tests/e2e-shared/scenarios/full-check-matrix.js';

test.describe('full locale/runtime pairwise matrix (vue-demo)', () => {
  test.describe.configure({ mode: 'parallel' });
  for (const row of FULL_CHECK_MATRIX) {
    test(`${row.locale}/${row.platform}/${row.theme}/${row.width}/${row.flow}`, async ({
      page,
    }) => {
      test.setTimeout(180_000);
      await runFullCheckMatrixRow(page, row);
    });
  }
});
