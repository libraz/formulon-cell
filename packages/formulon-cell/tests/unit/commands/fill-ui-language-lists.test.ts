import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { spreadsheetProfileForLocale } from '../../../src/mount/locale-profile.js';
import { LaneAllocator, type MatrixCase, runMatrixCase } from './fill-matrix-harness.js';

const fixture = JSON.parse(
  readFileSync(join(__dirname, '../../fixtures/excel-autofill-ui-language.json'), 'utf8'),
) as { helperIndex: number; languages: Record<string, MatrixCase[]> };

// Captured from Excel launched per UI language (-AppleLanguages), which supersedes the
// locale-switch capture in excel-autofill.json for the weekday and month lists.
// The desktop application's UI language decides which weekday and month lists extend; the
// host's counterpart is the workbook's spreadsheet profile locale.
describe('default fill of weekday and month names per UI language', () => {
  for (const [locale, cases] of Object.entries(fixture.languages)) {
    describe(locale, () => {
      let wb: WorkbookHandle;
      beforeAll(async () => {
        wb = await WorkbookHandle.createDefault();
        const profile = spreadsheetProfileForLocale(locale, 'mac');
        if (profile) wb.setSpreadsheetProfileId(profile);
      });
      afterAll(() => wb.dispose());

      const lanes = new LaneAllocator();
      for (const c of cases) {
        const lane = lanes.next(c.dir);
        it(c.id, () => {
          const diffs = runMatrixCase(wb, c, lane, fixture.helperIndex, {}, locale);
          expect(diffs, diffs.join('\n')).toEqual([]);
        });
      }
    });
  }
});
