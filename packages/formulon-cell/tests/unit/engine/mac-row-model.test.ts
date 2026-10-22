// @vitest-environment node

import { describe, expect, it } from 'vitest';
import { macDefaultRowPt } from '../../../src/engine/mac-row-model.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';

const PX_PER_PT = 96 / 72;

describe('macDefaultRowPt', () => {
  it.each([
    ['Calibri', 11, 15],
    ['Calibri', 18, 24],
    ['游ゴシック', 11, 18],
    ['游ゴシック', 12, 20],
    ['Yu Gothic', 11, 17],
    ['ＭＳ Ｐゴシック', 11, 14],
    ['MS PGothic', 11, 14],
    ['Meiryo UI', 11, 16],
    ['Aptos Narrow', 12, 16],
    ['DengXian', 12, 16],
    ['Malgun Gothic', 12, 18],
    ['新細明體', 12, 15],
    ['Tahoma', 12, 15],
  ] as const)('gives %s %d pt a %d pt row', (name, size, row) => {
    expect(macDefaultRowPt(name, size)).toBe(row);
  });

  it('returns null outside the table', () => {
    expect(macDefaultRowPt('Comic Sans MS', 11)).toBeNull();
    expect(macDefaultRowPt('Calibri', 13)).toBeNull();
  });
});

describe('getSheetDefaultLayoutPx on a loaded book', () => {
  it('recomputes the default row from the Normal font on Mac only', async () => {
    const created = await WorkbookHandle.createDefault({ locale: 'ja' });
    const loaded = await WorkbookHandle.loadBytes(created.save());
    created.dispose();
    try {
      // The Windows book states 18.75 pt; Mac Excel shows 游ゴシック 11 rows at 18 pt.
      loaded.geometryPlatform = 'default';
      expect(loaded.getSheetDefaultLayoutPx(0)?.defaultRowHeight).toBeCloseTo(18.75 * PX_PER_PT, 8);
      loaded.geometryPlatform = 'mac';
      expect(loaded.getSheetDefaultLayoutPx(0)?.defaultRowHeight).toBeCloseTo(18 * PX_PER_PT, 8);
    } finally {
      loaded.dispose();
    }
  });

  it('keeps the stated height for an untabulated Normal font on Mac', async () => {
    const wb = await WorkbookHandle.createDefault({ locale: 'en' });
    try {
      expect(wb.setWorkbookDefaultFont({ name: 'Comic Sans MS', size: 11 })).toBe(true);
      wb.geometryPlatform = 'mac';
      expect(wb.getSheetDefaultLayoutPx(0)?.defaultRowHeight).toBeCloseTo(15 * PX_PER_PT, 8);
    } finally {
      wb.dispose();
    }
  });
});
