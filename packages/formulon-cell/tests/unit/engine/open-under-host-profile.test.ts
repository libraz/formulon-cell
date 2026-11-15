import { describe, expect, it } from 'vitest';
import { Workbook, workbookHandleOf } from '../../../src/api/workbook.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { syncLocaleProfile } from '../../../src/mount/locale-profile.js';

const A1 = { sheet: 0, row: 0, col: 0 };
const B1 = { sheet: 0, row: 0, col: 1 };
const A2 = { sheet: 0, row: 1, col: 0 };
const B2 = { sheet: 0, row: 1, col: 1 };

/** Manual, calcOnSave off: B1 = A1*2 is current at 2; B2 = A2*2 is stale at 2 over A2 = 5. */
const mixedBytes = async (): Promise<Uint8Array> => {
  const wb = await WorkbookHandle.createDefault();
  try {
    wb.setNumber(A1, 1);
    wb.setFormula(B1, '=A1*2');
    wb.setNumber(A2, 1);
    wb.setFormula(B2, '=A2*2');
    wb.setCalcMode(1);
    wb.setCalcSetting('calcOnSave', false);
    wb.setNumber(A2, 5);
    return wb.save();
  } finally {
    wb.dispose();
  }
};

describe('opening a manual book under the host profile', () => {
  it.each([
    ['mac', 'ja'],
    ['default', 'ja'],
    ['mac', 'en'],
  ] as const)('keeps only the file stale state on %s / %s', async (platform, locale) => {
    const workbook = await Workbook.load(await mixedBytes(), { platform, locale });
    try {
      const handle = workbookHandleOf(workbook);
      syncLocaleProfile(handle, locale, platform);
      expect(handle.isCellStale(B1)).toBe(false);
      expect(handle.getValue(B1)).toEqual({ kind: 'number', value: 2 });
      expect(handle.isCellStale(B2)).toBe(true);
      expect(handle.getValue(B2)).toEqual({ kind: 'number', value: 2 });
    } finally {
      workbook.dispose();
    }
  });
});
