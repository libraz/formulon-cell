import { describe, expect, it } from 'vitest';

import { macDigitAdvance, roundDigitAdvance } from '../../../src/engine/mac-digit-metrics.js';
import { macDefaultRowPt } from '../../../src/engine/mac-row-model.js';

describe('Mac font aliases', () => {
  it('gives Yu Gothic its own entries: 游ゴシック digits, its own row height', () => {
    expect(macDefaultRowPt('Yu Gothic', 11)).not.toBe(macDefaultRowPt('游ゴシック', 11));
    expect(macDigitAdvance('Yu Gothic', 10)).toBe(macDigitAdvance('游ゴシック', 10));
    // Measured whole-point digit widths at 10, 11 and 12 pt.
    expect([10, 11, 12].map((pt) => roundDigitAdvance(macDigitAdvance('Yu Gothic', pt)))).toEqual([
      6, 6, 7,
    ]);
  });

  it('resolves the shared aliases in both tables', () => {
    expect(macDigitAdvance('MS PGothic', 10)).toBe(macDigitAdvance('ＭＳ Ｐゴシック', 10));
    expect(macDigitAdvance('メイリオ', 10)).toBe(macDigitAdvance('Meiryo', 10));
    expect(macDigitAdvance('DengXian', 10)).toBe(macDigitAdvance('等线', 10));
    expect(macDefaultRowPt('MS PGothic', 11)).toBe(macDefaultRowPt('ＭＳ Ｐゴシック', 11));
    expect(macDefaultRowPt('Malgun Gothic', 11)).toBe(macDefaultRowPt('맑은 고딕', 11));
  });
});
