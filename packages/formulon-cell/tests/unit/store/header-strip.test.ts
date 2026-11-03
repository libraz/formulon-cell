import { describe, expect, it } from 'vitest';
import { macDigitAdvance } from '../../../src/engine/mac-column-model.js';
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
