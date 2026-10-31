// @vitest-environment node
import { describe, expect, it } from 'vitest';
import { MAX_COL, MAX_ROW } from '../../src/engine/address.js';
import { patchXfIndexFromBase } from '../../src/engine/cell-xf-writeback.js';
import { WorkbookHandle } from '../../src/engine/workbook-handle.js';

describe('updated native range-style API', () => {
  it('rejects invalid remaps and snapshots without a partial native style change', async () => {
    const workbook = await WorkbookHandle.createDefault();
    const range = { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 2 };
    try {
      expect(workbook.capabilities.rangeStylePatch).toBe(true);
      expect(workbook.capabilities.rangeStyleSnapshot).toBe(true);
      const xf = patchXfIndexFromBase(workbook, 0, { fontSize: 32 });
      expect(xf).toBeGreaterThan(0);
      const before = workbook.snapshotRangeStyle(range);
      expect(before).not.toBeNull();
      expect(workbook.remapRangeStyle(range, [0, 0], [xf, xf])).toBe(false);
      expect(workbook.snapshotRangeStyle(range)).toEqual(before);
      expect(workbook.restoreRangeStyle(0, Uint8Array.of(1, 2, 3))).toBe(false);
      expect(workbook.snapshotRangeStyle(range)).toEqual(before);
    } finally {
      workbook.dispose();
    }
  });
  for (const locale of ['en', 'ja', 'de', 'fr', 'zh', 'ko', 'th'] as const) {
    it(`patches and restores column intersections without changing contents (${locale})`, async () => {
      const workbook = await WorkbookHandle.createDefault({ locale });
      const range = { sheet: 0, r0: 0, c0: 1, r1: MAX_ROW, c1: 1 };
      const addr = { sheet: 0, row: 2, col: 1 };
      const read = (row: number, col: number) => {
        const effective = workbook.getEffectiveStyle(0, row, col);
        const xf = effective && workbook.getCellXf(effective.xfIndex);
        const font = xf && workbook.getFontRecord(xf.fontIndex);
        return { size: font?.size, color: font?.colorArgb, xf, source: effective?.source };
      };
      try {
        expect(workbook.capabilities.rangeStylePatch).toBe(true);
        expect(workbook.capabilities.rangeStyleSnapshot).toBe(true);
        const rowXf = patchXfIndexFromBase(workbook, 0, {
          fontSize: 24,
          color: '#ff0000',
          numFmt: { kind: 'fixed', decimals: 2 },
        });
        expect(workbook.setRowStyleRange(0, 2, 3, rowXf)).toBe(true);
        workbook.setNumber(addr, 12.5);
        const before = read(2, 1);
        const beforeBlank = read(3, 1);
        const beforeStyle = workbook.snapshotRangeStyle(range);
        expect(beforeStyle).not.toBeNull();
        const bases = workbook.getRangeStyleBases(range);
        expect(bases).not.toBeNull();
        if (!bases || !beforeStyle) throw new Error('Missing native style capture');
        const patched = bases.map((base) => patchXfIndexFromBase(workbook, base, { fontSize: 32 }));
        expect(patched.every((xf) => xf >= 0)).toBe(true);
        expect(workbook.remapRangeStyle(range, bases, patched)).toBe(true);
        expect(read(2, 1)).toMatchObject({ size: 24, color: 0xffff0000 });
        expect(read(3, 1)).toMatchObject({ size: 24, color: 0xffff0000 });
        expect(read(2, 0)).toMatchObject({ size: 18, color: 0xffff0000 });
        expect(read(MAX_ROW, 1).size).toBe(24);
        expect(read(MAX_ROW, MAX_COL).size).toBe(11);
        expect(read(2, 1).xf?.numFmtId).toBe(before.xf?.numFmtId);
        expect(workbook.getNumFmtCode(read(2, 1).xf?.numFmtId ?? 0)).toBe('0.00');
        expect(workbook.getValue(addr)).toEqual({ kind: 'number', value: 12.5 });
        expect(workbook.restoreRangeStyle(0, beforeStyle)).toBe(true);
        expect(read(2, 1)).toEqual(before);
        expect(read(3, 1)).toEqual(beforeBlank);
        expect(workbook.getValue(addr)).toEqual({ kind: 'number', value: 12.5 });
        expect(workbook.clearRowStyleRange(0, 2, 3)).toBe(true);
        expect(read(3, 0).size).toBe(11);
      } finally {
        workbook.dispose();
      }
    });
  }
});
