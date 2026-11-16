import { describe, expect, it } from 'vitest';

import { macDigitAdvance } from '../../../src/engine/mac-digit-metrics.js';
import { macDefaultRowPt } from '../../../src/engine/mac-row-model.js';

describe('Mac font aliases', () => {
  it('keeps Yu Gothic apart from 游ゴシック in the digit metrics, as in the row model', () => {
    expect(macDefaultRowPt('Yu Gothic', 11)).not.toBe(macDefaultRowPt('游ゴシック', 11));
    // Yu Gothic has no digit-metric entry of its own, so it uses the unknown-family em.
    expect(macDigitAdvance('Yu Gothic', 10)).toBeCloseTo(5.5, 5);
    expect(macDigitAdvance('游ゴシック', 10)).toBeCloseTo(5.6, 5);
  });

  it('resolves the shared aliases in both tables', () => {
    expect(macDigitAdvance('MS PGothic', 10)).toBe(macDigitAdvance('ＭＳ Ｐゴシック', 10));
    expect(macDigitAdvance('メイリオ', 10)).toBe(macDigitAdvance('Meiryo', 10));
    expect(macDigitAdvance('DengXian', 10)).toBe(macDigitAdvance('等线', 10));
    expect(macDefaultRowPt('MS PGothic', 11)).toBe(macDefaultRowPt('ＭＳ Ｐゴシック', 11));
    expect(macDefaultRowPt('Malgun Gothic', 11)).toBe(macDefaultRowPt('맑은 고딕', 11));
  });
});
