import { describe, expect, it } from 'vitest';
import { planAxisSelection } from '../../../src/commands/axis-selection-plan.js';
import { MAX_COL, MAX_ROW } from '../../../src/engine/address.js';

describe('compact axis selection geometry', () => {
  it('combines disjoint bands, multiple holes and an ordinary cell within one budget', () => {
    const selected = [
      { sheet: 0, r0: 0, r1: 3, c0: 1, c1: 1 },
      { sheet: 0, r0: 5, r1: 6, c0: 1, c1: 1 },
      { sheet: 0, r0: 8, r1: MAX_ROW, c0: 1, c1: 1 },
      { sheet: 0, r0: 0, r1: MAX_ROW, c0: 3, c1: 3 },
      { sheet: 0, r0: 9, r1: 9, c0: 5, c1: 5 },
    ];
    expect(planAxisSelection(selected, 3)).toEqual({
      axisBands: [
        { axis: 'column', first: 1, last: 1 },
        { axis: 'column', first: 3, last: 3 },
      ],
      ordinaryRanges: [selected[4]],
      exceptionRanges: [
        { sheet: 0, r0: 4, r1: 4, c0: 1, c1: 1 },
        { sheet: 0, r0: 7, r1: 7, c0: 1, c1: 1 },
      ],
    });
    expect(planAxisSelection(selected, 2)).toBeNull();
  });

  it('uses a column overlay for almost all rows rather than a million row calls', () => {
    const plan = planAxisSelection([{ sheet: 0, r0: 0, r1: MAX_ROW - 1, c0: 0, c1: MAX_COL }]);
    expect(plan).toEqual({
      axisBands: [{ axis: 'column', first: 0, last: MAX_COL }],
      ordinaryRanges: [],
      exceptionRanges: [{ sheet: 0, r0: MAX_ROW, r1: MAX_ROW, c0: 0, c1: MAX_COL }],
    });
  });
  it('compensates excluded edge cells as well as interior holes', () => {
    expect(planAxisSelection([{ sheet: 0, r0: 1, r1: MAX_ROW - 1, c0: 1, c1: 1 }])).toEqual({
      axisBands: [{ axis: 'column', first: 1, last: 1 }],
      ordinaryRanges: [],
      exceptionRanges: [
        { sheet: 0, r0: 0, r1: 0, c0: 1, c1: 1 },
        { sheet: 0, r0: MAX_ROW, r1: MAX_ROW, c0: 1, c1: 1 },
      ],
    });
  });
  it('recognizes a column with a removed middle cell', () => {
    expect(
      planAxisSelection([
        { sheet: 0, r0: 0, r1: 3, c0: 1, c1: 1 },
        { sheet: 0, r0: 5, r1: MAX_ROW, c0: 1, c1: 1 },
      ]),
    ).toEqual({
      axisBands: [{ axis: 'column', first: 1, last: 1 }],
      ordinaryRanges: [],
      exceptionRanges: [{ sheet: 0, r0: 4, r1: 4, c0: 1, c1: 1 }],
    });
  });

  it('recognizes eight complete rows with one excluded cell', () => {
    const plan = planAxisSelection([
      { sheet: 0, r0: 2, r1: 9, c0: 0, c1: 0 },
      { sheet: 0, r0: 2, r1: 9, c0: 2, c1: MAX_COL },
      { sheet: 0, r0: 2, r1: 3, c0: 1, c1: 1 },
      { sheet: 0, r0: 5, r1: 9, c0: 1, c1: 1 },
    ]);
    expect(plan).toEqual({
      axisBands: [{ axis: 'row', first: 2, last: 9 }],
      ordinaryRanges: [],
      exceptionRanges: [{ sheet: 0, r0: 4, r1: 4, c0: 1, c1: 1 }],
    });
  });

  it('recognizes a mixed row/column union already fragmented by selection subtraction', () => {
    const plan = planAxisSelection([
      { sheet: 0, r0: 0, r1: MAX_ROW, c0: 1, c1: 1 },
      { sheet: 0, r0: 2, r1: 2, c0: 0, c1: 0 },
      { sheet: 0, r0: 2, r1: 2, c0: 2, c1: MAX_COL },
    ]);
    expect(plan).toEqual({
      axisBands: [
        { axis: 'column', first: 1, last: 1 },
        { axis: 'row', first: 2, last: 2 },
      ],
      ordinaryRanges: [],
      exceptionRanges: [],
    });
  });

  it('keeps small rectangles ordinary and leaves their holes unselected', () => {
    const ranges = [
      { sheet: 0, r0: 0, r1: 0, c0: 0, c1: 2 },
      { sheet: 0, r0: 2, r1: 2, c0: 0, c1: 2 },
      { sheet: 0, r0: 1, r1: 1, c0: 0, c1: 0 },
      { sheet: 0, r0: 1, r1: 1, c0: 2, c1: 2 },
    ];
    expect(planAxisSelection(ranges)).toEqual({
      axisBands: [],
      ordinaryRanges: ranges,
      exceptionRanges: [],
    });
  });

  it('rejects a complement over budget without enumerating a million cells', () => {
    expect(
      planAxisSelection([
        { sheet: 0, r0: 0, r1: 0, c0: 1, c1: 1 },
        { sheet: 0, r0: 100002, r1: MAX_ROW, c0: 1, c1: 1 },
      ]),
    ).toBeNull();
  });

  it('chooses column coverage for a complete worksheet', () => {
    expect(planAxisSelection([{ sheet: 0, r0: 0, r1: MAX_ROW, c0: 0, c1: MAX_COL }])).toEqual({
      axisBands: [{ axis: 'column', first: 0, last: MAX_COL }],
      ordinaryRanges: [],
      exceptionRanges: [],
    });
  });
});
