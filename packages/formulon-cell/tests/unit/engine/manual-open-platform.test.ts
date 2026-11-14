import { describe, expect, it } from 'vitest';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { syncLocaleProfile } from '../../../src/mount/locale-profile.js';

const A2 = { sheet: 0, row: 1, col: 0 };
const B2 = { sheet: 0, row: 1, col: 1 };

/** A book saved in `mode` with B2 = A2*2 computed at A2 = 1, then A2 set to 5. */
const savedBytes = async (mode: 0 | 1): Promise<Uint8Array> => {
  const wb = await WorkbookHandle.createDefault();
  try {
    wb.setNumber(A2, 1);
    wb.setFormula(B2, '=A2*2');
    wb.setCalcMode(1);
    wb.setCalcSetting('calcOnSave', false);
    wb.setNumber(A2, 5);
    if (mode === 0) wb.setCalcMode(0);
    return wb.save();
  } finally {
    wb.dispose();
  }
};

describe('opening a manual-calculation book under another spreadsheet profile', () => {
  it.each([
    ['the Mac platform', (wb: WorkbookHandle) => (wb.geometryPlatform = 'mac')],
    ['a Japanese locale profile', (wb: WorkbookHandle) => syncLocaleProfile(wb, 'ja', 'default')],
    ['the Mac Japanese locale profile', (wb: WorkbookHandle) => syncLocaleProfile(wb, 'ja', 'mac')],
  ])('keeps its stale values on %s', async (_, open) => {
    const wb = await WorkbookHandle.loadBytes(await savedBytes(1));
    try {
      open(wb);
      expect(wb.isCellStale(B2)).toBe(true);
      expect(wb.getValue(B2)).toEqual({ kind: 'number', value: 2 });
    } finally {
      wb.dispose();
    }
  });

  it('still recalculates an automatic book when the platform changes', async () => {
    const wb = await WorkbookHandle.loadBytes(await savedBytes(0));
    try {
      wb.geometryPlatform = 'mac';
      expect(wb.getValue(B2)).toEqual({ kind: 'number', value: 10 });
    } finally {
      wb.dispose();
    }
  });
});
