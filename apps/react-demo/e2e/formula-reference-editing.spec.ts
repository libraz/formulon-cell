import { test } from '@playwright/test';
import {
  type FormulaEditEntry,
  runFormulaReferenceEditing,
} from '../../../tests/e2e-shared/scenarios/formula-reference-editing.js';

for (const platform of ['default', 'mac'] as const) {
  for (const entry of ['f2', 'double-click', 'formula-bar'] satisfies FormulaEditEntry[]) {
    test(`react-demo: ${platform} ${entry} preserves formulas during reference clicks and range drags`, async ({
      page,
    }) => {
      await runFormulaReferenceEditing(page, entry, platform);
    });
  }
}
