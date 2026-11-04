import { describe, expect, it } from 'vitest';
import { textBaselineY } from '../../../../src/render/painters/text.js';

// Bottom-aligned baselines measured in Mac Excel for 11 pt text in a 24 pt row
// at 50-200% zoom (cells-zoom-text-bottom-second-font facts), as CSS px from the
// row's top edge. A cell's paint bounds run 1 CSS px past the row pitch. Sizes
// are the painted whole-px sizes at each zoom (11 pt is 8 px at 100%).
const ZOOMS = [
  { zoom: 50, size: 4, rowH: 12 },
  { zoom: 75, size: 6, rowH: 18 },
  { zoom: 90, size: 7, rowH: 22 },
  { zoom: 100, size: 8, rowH: 24 },
  { zoom: 110, size: 9, rowH: 26 },
  { zoom: 125, size: 10, rowH: 30 },
  { zoom: 150, size: 12, rowH: 36 },
  { zoom: 175, size: 14, rowH: 42 },
  { zoom: 200, size: 17, rowH: 48 },
];
const CALIBRI_BASELINE = [11, 16, 19, 21, 23, 26, 31, 36, 41];
// MS PGothic at 50% is an invalid capture; 175/200% land on half CSS px.
const MSP_BASELINE = [null, 17, 20, 22, 24, 27, 32, 37.5, 42.5];

const baseline = (family: string, size: number, rowH: number): number =>
  textBaselineY(
    { x: 0, y: 0, w: 100, h: rowH + 1 },
    { ascent: size * 0.9, descent: size * 0.2 },
    'bottom',
    1,
    false,
    {
      family,
      size,
    },
  );

describe('bottom baseline of fonts measured in Excel', () => {
  it('places Calibri on the measured baseline at every zoom', () => {
    ZOOMS.forEach(({ size, rowH }, k) => {
      expect(baseline('Calibri, sans-serif', size, rowH)).toBe(CALIBRI_BASELINE[k]);
    });
  });

  it('places MS PGothic within half a CSS px of the measured baseline', () => {
    ZOOMS.forEach(({ size, rowH }, k) => {
      const want = MSP_BASELINE[k];
      if (want == null) return;
      for (const family of ['"ＭＳ Ｐゴシック", sans-serif', 'MS PGothic']) {
        expect(Math.abs(baseline(family, size, rowH) - want)).toBeLessThanOrEqual(0.5);
      }
    });
  });

  it('keeps the descent formula for a font without measurements', () => {
    expect(baseline('Arial', 17, 48)).toBe(
      textBaselineY({ x: 0, y: 0, w: 100, h: 49 }, { ascent: 15.3, descent: 3.4 }, 'bottom', 1),
    );
  });
});
