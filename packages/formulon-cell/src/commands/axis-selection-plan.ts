import { MAX_COL, MAX_ROW } from '../engine/address.js';
import type { AxisFormatBand } from '../engine/axis-format-sync.js';
import type { Range } from '../engine/types.js';
import { subtractRange } from '../store/selection-geometry.js';

export interface AxisSelectionPlan {
  readonly axisBands: readonly AxisFormatBand[];
  readonly ordinaryRanges: readonly Range[];
  /** Unselected cells whose original appearance must compensate an axis write. */
  readonly exceptionRanges: readonly Range[];
}

const union = (ranges: readonly Range[]): Range[] => {
  const result: Range[] = [];
  for (const range of ranges) {
    let pieces = [{ ...range }];
    for (const previous of result)
      pieces = pieces.flatMap((piece) => subtractRange(piece, previous));
    result.push(...pieces);
  }
  return result;
};

const difference = (ranges: readonly Range[], covers: readonly Range[]): Range[] => {
  let remaining = ranges.map((range) => ({ ...range }));
  for (const cover of covers) remaining = remaining.flatMap((range) => subtractRange(range, cover));
  return remaining;
};

const area = (ranges: readonly Range[], budget: number): number => {
  let count = 0;
  for (const range of ranges) {
    count += (range.r1 - range.r0 + 1) * (range.c1 - range.c0 + 1);
    if (count > budget) return count;
  }
  return count;
};

const mergeBands = (bands: readonly AxisFormatBand[]): AxisFormatBand[] => {
  const result: AxisFormatBand[] = [];
  for (const axis of ['column', 'row'] as const) {
    const sorted = bands.filter((band) => band.axis === axis).sort((a, b) => a.first - b.first);
    for (const band of sorted) {
      const previous = result.at(-1);
      if (previous?.axis === axis && band.first <= previous.last + 1) {
        result[result.length - 1] = { ...previous, last: Math.max(previous.last, band.last) };
      } else result.push({ ...band });
    }
  }
  // A complete column span also covers every row, without a million row writes.
  if (result.some((band) => band.axis === 'column' && band.first === 0 && band.last === MAX_COL)) {
    return result.filter((band) => band.axis === 'column');
  }
  return result;
};

interface SweptBands {
  exact: AxisFormatBand[];
  near: AxisFormatBand[];
}

const sweep = (ranges: readonly Range[], axis: 'row' | 'column', budget: number): SweptBands => {
  const boundaries = new Set<number>();
  const lastAcross = axis === 'column' ? MAX_ROW : MAX_COL;
  for (const range of ranges) {
    boundaries.add(axis === 'column' ? range.c0 : range.r0);
    boundaries.add((axis === 'column' ? range.c1 : range.r1) + 1);
  }
  const ordered = [...boundaries].sort((a, b) => a - b);
  const result: SweptBands = { exact: [], near: [] };
  for (let index = 0; index + 1 < ordered.length; index++) {
    const first = ordered[index] as number;
    const last = (ordered[index + 1] as number) - 1;
    const intervals = ranges
      .filter((range) =>
        axis === 'column'
          ? range.c0 <= first && range.c1 >= last
          : range.r0 <= first && range.r1 >= last,
      )
      .map((range) =>
        axis === 'column' ? ([range.r0, range.r1] as const) : ([range.c0, range.c1] as const),
      )
      .sort((a, b) => a[0] - b[0]);
    if (intervals.length === 0) continue;
    let cursor = 0;
    let holes = 0;
    for (const [start, end] of intervals) {
      if (start > cursor) holes += start - cursor;
      cursor = Math.max(cursor, end + 1);
    }
    holes += lastAcross + 1 - cursor;
    const band = { axis, first, last };
    if (holes === 0) result.exact.push(band);
    else if (holes * (last - first + 1) <= budget) result.near.push(band);
  }
  return result;
};

/** Find compact row/column coverage and a bounded set of ordinary/excluded cells. */
export function planAxisSelection(
  ranges: readonly Range[],
  budget = 100_000,
): AxisSelectionPlan | null {
  const sheet = ranges[0]?.sheet;
  if (sheet === undefined || !Number.isInteger(budget) || budget < 0) return null;
  if (
    !ranges.every(
      (range) =>
        range.sheet === sheet &&
        Number.isInteger(sheet) &&
        sheet >= 0 &&
        [range.r0, range.r1, range.c0, range.c1].every(Number.isInteger) &&
        range.r0 >= 0 &&
        range.c0 >= 0 &&
        range.r1 <= MAX_ROW &&
        range.c1 <= MAX_COL &&
        range.r0 <= range.r1 &&
        range.c0 <= range.c1,
    )
  )
    return null;
  const selected = union(ranges);
  const columns = sweep(selected, 'column', budget);
  const rows = sweep(selected, 'row', budget);
  const exact = [...columns.exact, ...rows.exact];
  const candidates = [
    exact,
    [...exact, ...columns.near],
    [...exact, ...rows.near],
    [...exact, ...columns.near, ...rows.near],
  ];
  let best: AxisSelectionPlan | null = null;
  let smallest = Infinity;
  for (const candidate of candidates) {
    const axisBands = mergeBands(candidate);
    const rowWrites = axisBands.reduce(
      (count, band) => count + (band.axis === 'row' ? band.last - band.first + 1 : 0),
      0,
    );
    if (rowWrites > budget) continue;
    const coverage = union(
      axisBands.map((band) =>
        band.axis === 'column'
          ? { sheet, r0: 0, r1: MAX_ROW, c0: band.first, c1: band.last }
          : { sheet, r0: band.first, r1: band.last, c0: 0, c1: MAX_COL },
      ),
    );
    const ordinaryRanges = difference(selected, coverage);
    const exceptionRanges = difference(coverage, selected);
    const cost = area(ordinaryRanges, budget) + area(exceptionRanges, budget);
    if (cost > budget || cost >= smallest) continue;
    smallest = cost;
    best = { axisBands, ordinaryRanges, exceptionRanges };
  }
  return best;
}
