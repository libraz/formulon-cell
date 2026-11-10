import { describe, expect, it } from 'vitest';
import { segmentsFor } from '../../../src/charts/series-geometry.js';
import type { CategoryChartData } from '../../../src/charts/types.js';

const SERIES = 40;
const CATEGORIES = 4;

const build = (counter: { reads: number }): CategoryChartData => ({
  domain: 'category',
  categories: Array.from({ length: CATEGORIES }, (_, i) => ({ label: `c${i}` })),
  series: Array.from({ length: SERIES }, (_, s) => {
    const values = Array.from({ length: CATEGORIES }, (_, c) =>
      (s + c) % 3 === 0 ? -(s + 1) : s + 1,
    );
    return {
      id: `s${s}`,
      points: new Proxy(
        values.map((value) => ({ value })),
        {
          get(target, key, receiver) {
            if (typeof key === 'string' && /^\d+$/.test(key)) counter.reads += 1;
            return Reflect.get(target, key, receiver);
          },
        },
      ),
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
      categories: [{ label: 'a' }],
      series: [
        { id: '1', points: [{ value: 2 }] },
        { id: '2', points: [{ value: -3 }] },
        { id: '3', points: [{ value: 4 }] },
        { id: '4', points: [{ value: null }] },
        { id: '5', points: [{ value: -1 }] },
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
