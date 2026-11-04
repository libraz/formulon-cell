import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import type { FillOptions } from '../../../src/commands/fill.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { spreadsheetProfileForLocale } from '../../../src/mount/locale-profile.js';
import {
  LaneAllocator,
  type MatrixCase,
  type MatrixFile,
  runMatrixCase,
} from './fill-matrix-harness.js';

interface EdgeCase extends MatrixCase {
  typ: number | null;
}

const matrix = JSON.parse(
  readFileSync(join(__dirname, '../../fixtures/excel-autofill-edge-matrix.json'), 'utf8'),
) as MatrixFile<EdgeCase>;

const LOCALE = 'ja-JP';
// A plain fill-handle drag equals the default fill type, so `null` carries no options.
const OPTIONS: Record<number, FillOptions> = {
  1: { copyOnly: true },
  2: { series: true },
  5: { dateUnit: 'days' },
  6: { dateUnit: 'weekdays' },
  7: { dateUnit: 'months' },
  8: { dateUnit: 'years' },
  10: { growth: true },
};

describe('fillRange matches the measured AutoFill edge matrix (ja-JP)', () => {
  let wb: WorkbookHandle;
  beforeAll(async () => {
    wb = await WorkbookHandle.createDefault();
    const profile = spreadsheetProfileForLocale(LOCALE, 'mac');
    if (profile) wb.setSpreadsheetProfileId(profile);
  });
  afterAll(() => wb.dispose());

  const lanes = new LaneAllocator();
  matrix.cases.forEach((c, index) => {
    const lane = lanes.next(c.dir);
    it(`${index} ${c.id} ${c.dir} typ=${c.typ}`, () => {
      const options = c.typ === null ? {} : (OPTIONS[c.typ] ?? {});
      const diffs = runMatrixCase(wb, c, lane, matrix.helperIndex, options, LOCALE);
      expect(diffs, diffs.join('\n')).toEqual([]);
    });
  });
});
