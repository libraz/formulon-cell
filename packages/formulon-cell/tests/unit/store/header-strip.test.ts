import { describe, expect, it } from 'vitest';
import { macDigitAdvance } from '../../../src/engine/mac-digit-metrics.js';
import { macDefaultRowPt } from '../../../src/engine/mac-row-model.js';
import { headerStripSize } from '../../../src/store/header-strip.js';

const PX_PER_PT = 96 / 72;

/** Mac strips for a Normal font at 100%, in device px at DPR 2 like the
 *  measurements: [row-number strip with 3-digit labels, with 4-digit labels,
 *  column-letter strip]. */
const strips = (family: string, sizePt: number): [number, number, number] => {
  const layout = {
    headerColWidth: 23,
    headerRowHeight: 22,
    headerDigitAdvance: macDigitAdvance(family, sizePt),
    defaultRowHeight: (macDefaultRowPt(family, sizePt) ?? 0) * PX_PER_PT,
    screenScale: 0.75,
  };
  const three = headerStripSize(layout, { rowStart: 0, rowCount: 100 });
  const four = headerStripSize(layout, { rowStart: 994, rowCount: 100 });
  return [three.headerColWidth * 2, four.headerColWidth * 2, three.headerRowHeight * 2];
};

describe('headerStripSize on the Mac profile', () => {
  it.each([
    ['Courier New', 9, [44, 50, 44]],
    ['Courier New', 10, [46, 58, 44]],
    ['Courier New', 14, [58, 74, 44]],
    ['Courier New', 16, [74, 94, 44]],
    ['Courier New', 18, [80, 102, 48]],
    ['Courier New', 20, [86, 110, 54]],
    ['Courier New', 24, [102, 130, 64]],
    ['Courier New', 28, [124, 158, 74]],
    ['MS PGothic', 9, [44, 50, 44]],
    ['MS PGothic', 11, [46, 58, 44]],
    ['MS PGothic', 14, [52, 66, 44]],
    ['MS PGothic', 16, [58, 74, 44]],
    ['MS PGothic', 18, [68, 86, 44]],
    ['MS PGothic', 24, [86, 110, 56]],
    ['MS PGothic', 28, [102, 130, 66]],
  ] as const)('sizes the strips for %s %d pt as measured', (family, size, expected) => {
    expect(strips(family, size)).toEqual(expected);
  });
});

/** Row-number strip width (device px) at `zoom` for a label of `digits`. */
const stripAt = (family: string, sizePt: number, zoom: number, digits: number): number[] => {
  const layout = {
    headerColWidth: 23,
    headerRowHeight: 22,
    headerDigitAdvance: macDigitAdvance(family, sizePt),
    defaultRowHeight: (macDefaultRowPt(family, sizePt) ?? 0) * PX_PER_PT,
    screenScale: 0.75,
  };
  const size = headerStripSize(layout, {
    rowStart: 10 ** (digits - 1) - 1,
    rowCount: 1,
    zoom,
  });
  return [size.headerColWidth * 2, size.headerRowHeight * 2];
};

describe('headerStripSize at a small Normal font and a low zoom', () => {
  // Small Normal font: 3 / 4 / 5 / 6-digit strips at 100%.
  it.each([
    ['Courier New', 6, [44, 44, 46, 54]],
    ['Courier New', 7, [44, 44, 46, 54]],
    ['Courier New', 8, [44, 50, 60, 70]],
    ['Arial', 7, [44, 44, 46, 54]],
    ['Arial', 8, [44, 44, 46, 54]],
  ] as const)('sizes the row strip for %s %d pt as measured', (family, size, widths) => {
    expect([3, 4, 5, 6].map((d) => stripAt(family, size, 1, d)[0])).toEqual(widths);
  });

  // Low zoom (MS PGothic 11): [width for 3..6 digits, column strip height].
  it.each([
    [1, [46, 58, 70, 82], 44],
    [0.5, [24, 30, 36, 42], 22],
    [0.4, [18, 22, 26, 30], 18],
    [0.33, [18, 22, 26, 30], 16],
    [0.25, [18, 22, 26, 30], 16],
  ] as const)('sizes the strips at %d zoom as measured', (zoom, widths, height) => {
    expect([3, 4, 5, 6].map((d) => stripAt('MS PGothic', 11, zoom, d)[0])).toEqual(widths);
    expect(stripAt('MS PGothic', 11, zoom, 3)[1]).toBe(height);
  });
});
