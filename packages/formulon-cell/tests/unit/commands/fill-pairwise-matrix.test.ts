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

interface PairwiseCase extends MatrixCase {
  ft: string;
}

const matrix = JSON.parse(
  readFileSync(join(__dirname, '../../fixtures/excel-autofill-pairwise.json'), 'utf8'),
) as MatrixFile<PairwiseCase>;

const LOCALE = 'ja-JP';
const OPTIONS: Record<string, FillOptions> = {
  default: {},
  series: { series: true },
  copy: { copyOnly: true },
  days: { dateUnit: 'days' },
  weekdays: { dateUnit: 'weekdays' },
  months: { dateUnit: 'months' },
  years: { dateUnit: 'years' },
};

describe('fillRange matches the measured pairwise AutoFill matrix (ja-JP)', () => {
  let wb: WorkbookHandle;
  beforeAll(async () => {
    wb = await WorkbookHandle.createDefault();
    const profile = spreadsheetProfileForLocale(LOCALE, 'mac');
    if (profile) wb.setSpreadsheetProfileId(profile);
  });
  afterAll(() => wb.dispose());

  const lanes = new LaneAllocator();
  for (const c of matrix.cases) {
    const lane = lanes.next(c.dir);
    it(c.id, () => {
      const diffs = runMatrixCase(wb, c, lane, matrix.helperIndex, OPTIONS[c.ft] ?? {}, LOCALE);
      expect(diffs, diffs.join('\n')).toEqual([]);
    });
  }
});
