import { describe, expect, it } from 'vitest';
import { segmentsFor } from '../../../src/charts/series-geometry.js';
import type { CategoryChartData } from '../../../src/charts/types.js';

const SERIES = 40;
const CATEGORIES = 4;

const build = (counter: { reads: number }): CategoryChartData => ({
  domain: 'category',
  categories: Array.from({ length: CATEGORIES }, (_, i) => `c${i}`),
  series: Array.from({ length: SERIES }, (_, s) => {
    const values = Array.from({ length: CATEGORIES }, (_, c) =>
      (s + c) % 3 === 0 ? -(s + 1) : s + 1,
    );
    return {
      id: `s${s}`,
      values: new Proxy(values, {
        get(target, key, receiver) {
          if (typeof key === 'string' && /^\d+$/.test(key)) counter.reads += 1;
          return Reflect.get(target, key, receiver);
        },
      }),
    };
  }),
});

describe('stacked segments', () => {
  it('reads each value a bounded number of times regardless of series count', () => {
    for (const grouping of ['stacked', 'percent-stacked'] as const) {
      const counter = { reads: 0 };
      segmentsFor(build(counter), grouping);
      expect(counter.reads).toBeLessThan(SERIES * CATEGORIES * 6);
    }
  });

  it('accumulates positive and negative runs separately', () => {
    const data: CategoryChartData = {
      domain: 'category',
      categories: ['a'],
      series: [
        { id: '1', values: [2] },
        { id: '2', values: [-3] },
        { id: '3', values: [4] },
        { id: '4', values: [null] },
        { id: '5', values: [-1] },
      ],
    };
    const seg = segmentsFor(data, 'stacked').map((s) => s[0]);
    expect(seg.map((s) => s && [s.start, s.end])).toEqual([
      [0, 2],
      [0, -3],
      [2, 6],
      null,
      [-3, -4],
    ]);
    const pct = segmentsFor(data, 'percent-stacked').map((s) => s[0]);
    expect(pct[2]?.start).toBeCloseTo(2 / 6);
    expect(pct[2]?.end).toBeCloseTo(1);
    expect(pct[4]?.end).toBeCloseTo(-1);
  });
});
