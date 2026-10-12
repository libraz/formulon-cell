import { describe, expect, it } from 'vitest';
import { WorkbookHandle } from '../../src/engine/workbook-handle.js';
import { WorkbookHandleStylesMethods } from '../../src/engine/workbook-handle-styles.js';

describe('clear direct native XF', () => {
  it('returns false for an older native host without a clear method', () => {
    const oldHost = {
      assertAlive() {},
      wb: {},
    } as unknown as WorkbookHandleStylesMethods;
    expect(WorkbookHandleStylesMethods.prototype.clearCellXfIndex.call(oldHost, 0, 2, 1)).toBe(
      false,
    );
  });

  it.each(['en', 'ja', 'de', 'fr', 'zh', 'ko', 'th'])(
    '%s restores blank-cell inheritance and follows a later row style',
    async (locale) => {
      const workbook = await WorkbookHandle.createDefault({ locale });
      const addr = { sheet: 0, row: 2, col: 1 };
      try {
        expect(workbook.isStub).toBe(false);
        expect(workbook.capabilities.cellFormattingClear).toBe(true);
        const normal = workbook.getCellXf(0);
        const font = normal && workbook.getFontRecord(normal.fontIndex);
        if (!normal || !font) throw new Error('Missing native Normal font');
        const xfForSize = (size: number) => {
          const fontIndex = workbook.addFontRecord({ ...font, size });
          expect(fontIndex).toBeGreaterThanOrEqual(0);
          const xf = workbook.addXfRecord({ ...normal, fontIndex, applyFont: true });
          expect(xf).toBeGreaterThan(0);
          return xf;
        };
        const rowXf = xfForSize(18);
        const directXf = xfForSize(30);
        expect(workbook.setRowStyle(0, 2, rowXf)).toBe(true);
        expect(workbook.setCellXfIndex(0, 2, 1, directXf)).toBe(true);
        expect(workbook.getEffectiveStyle(0, 2, 1)).toMatchObject({
          source: 0,
          xfIndex: directXf,
        });
        expect(workbook.clearCellXfIndex(0, 2, 1)).toBe(true);
        expect(workbook.getEffectiveStyle(0, 2, 1)).toMatchObject({ source: 1, xfIndex: rowXf });
        const laterRowXf = xfForSize(36);
        expect(workbook.setRowStyle(0, 2, laterRowXf)).toBe(true);
        expect(workbook.getEffectiveStyle(0, 2, 1)).toMatchObject({
          source: 1,
          xfIndex: laterRowXf,
        });
        expect(workbook.cellFormula(addr)).toBeNull();
        expect(workbook.getValue(addr)).toEqual({ kind: 'blank' });
      } finally {
        workbook.dispose();
      }
    },
  );

  it('preserves a formula and its result when clearing the direct XF', async () => {
    const workbook = await WorkbookHandle.createDefault({ locale: 'ja' });
    const addr = { sheet: 0, row: 2, col: 1 };
    try {
      workbook.setFormula(addr, '=1+2');
      workbook.recalc();
      expect(workbook.setCellXfIndex(0, 2, 1, 0)).toBe(true);
      expect(workbook.clearCellXfIndex(0, 2, 1)).toBe(true);
      expect(workbook.cellFormula(addr)).toBe('=1+2');
      expect(workbook.getValue(addr)).toEqual({ kind: 'number', value: 3 });
      // Nonblank inheritance is a separate native defect, documented in
      // formulon/backup/fb-clear-cell-xf-nonblank-inheritance-2026-10-12.md.
    } finally {
      workbook.dispose();
    }
  });
});
