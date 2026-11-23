import { describe, expect, it } from 'vitest';
import { fontPointsToCssPx } from '../../../../src/engine/font-units.js';
import { textBaselineY } from '../../../../src/render/painters/text.js';

// MS PGothic and Calibri bottom-aligned baselines measured for 11 pt text in a 24 pt row
// at 50-200% zoom, as CSS px from the row's top edge. A cell's paint bounds run
// 1 CSS px past the row pitch. Sizes are the painted whole-px sizes at each zoom.
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
// Calibri baselines from the same capture (Mac scale, 0.75 x zoom CSS px per pt).
const CALIBRI_BASELINE = [11, 16, 19, 21, 23, 26, 31, 36, 41];
// MS PGothic at 50% is an invalid capture; 175/200% land on half CSS px.
const MSP_BASELINE = [null, 17, 20, 22, 24, 27, 32, 37.5, 42.5];

const baseline = (family: string, size: number, rowH: number, scale?: number): number =>
  textBaselineY(
    { x: 0, y: 0, w: 100, h: rowH + 1 },
    { ascent: size * 0.9, descent: size * 0.2 },
    'bottom',
    1,
    false,
    {
      family,
      size,
      scale,
    },
  );

describe('bottom baseline of fonts measured in Excel', () => {
  it.each(['Calibri, sans-serif', '"Aptos Narrow", Calibri', 'aptos narrow'])(
    'keeps %s 4 pt above the bottom edge at every size and row height',
    (family) => {
      for (const size of [8, 11, 14, 20, 36]) {
        for (const rowH of [24, 48, 80]) {
          // Mac scale: 4 pt is 4 screen px.
          expect(baseline(family, size, rowH, 0.75)).toBe(Math.round(rowH + 1 - 4));
          expect(baseline(family, size, rowH, 1)).toBe(Math.round(rowH + 1 - fontPointsToCssPx(4)));
        }
      }
    },
  );

  it('places Calibri on the measured baseline at every zoom', () => {
    ZOOMS.forEach(({ zoom, size, rowH }, k) => {
      expect(baseline('Calibri', size, rowH, 0.75 * (zoom / 100))).toBe(CALIBRI_BASELINE[k]);
    });
  });

  it('scales the 4 pt bottom gap with the render scale', () => {
    expect(baseline('Calibri', 17, 72, 1.5)).toBe(Math.round(73 - fontPointsToCssPx(4) * 1.5));
    expect(baseline('Aptos Narrow', 6, 18, 0.375)).toBe(Math.round(19 - 2));
  });

  it('leaves room for 11 pt text in the default 20 px row', () => {
    // 15 pt row at scale 1: the baseline sits below the cap height of the glyphs.
    expect(baseline('Calibri', fontPointsToCssPx(11), 20, 1)).toBe(16);
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
