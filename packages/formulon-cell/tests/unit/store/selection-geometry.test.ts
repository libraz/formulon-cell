import { describe, expect, it } from 'vitest';

import type { Range } from '../../../src/engine/types.js';
import {
  applySelectionRectangle,
  rangeArea,
  rangeContainsAddr,
  rangeContainsRange,
  rangesIntersect,
  sameRange,
  selectionContainsAddr,
  selectionCoversRange,
  subtractRange,
} from '../../../src/store/selection-geometry.js';
import type { SelectionSlice } from '../../../src/store/types.js';

const range = (r0: number, c0: number, r1 = r0, c1 = c0, sheet = 0): Range => ({
  sheet,
  r0,
  c0,
  r1,
  c1,
});

const selection = (primary: Range, extras: Range[] = []): SelectionSlice => ({
  active: { sheet: primary.sheet, row: primary.r0, col: primary.c0 },
  anchor: { sheet: primary.sheet, row: primary.r0, col: primary.c0 },
  range: primary,
  extraRanges: extras,
});

describe('store/selection-geometry', () => {
  it('subtracts a center rectangle into four disjoint pieces', () => {
    expect(subtractRange(range(0, 0, 4, 4), range(2, 2))).toEqual([
      range(0, 0, 1, 4),
      range(3, 0, 4, 4),
      range(2, 0, 2, 1),
      range(2, 3, 2, 4),
    ]);
  });

  it('subtracts edges and corners without empty fragments', () => {
    expect(subtractRange(range(0, 0, 2, 2), range(0, 1, 0, 2))).toEqual([
      range(1, 0, 2, 2),
      range(0, 0),
    ]);
    expect(subtractRange(range(0, 0, 2, 2), range(0, 0))).toEqual([
      range(1, 0, 2, 2),
      range(0, 1, 0, 2),
    ]);
  });

  it('returns a cloned range for no overlap and an empty list for full coverage', () => {
    const original = range(1, 2, 4, 5);
    const untouched = subtractRange(original, range(8, 8));
    expect(untouched).toEqual([original]);
    expect(untouched[0]).not.toBe(original);
    expect(subtractRange(original, range(0, 0, 8, 8))).toEqual([]);
    expect(subtractRange(original, range(0, 0, 8, 8, 1))).toEqual([original]);
  });

  it('handles sheet-sized rectangles with coordinate partitions only', () => {
    expect(subtractRange(range(0, 0, 1_048_575, 16_383), range(500, 500))).toEqual([
      range(0, 0, 499, 16_383),
      range(501, 0, 1_048_575, 16_383),
      range(500, 0, 500, 499),
      range(500, 501, 500, 16_383),
    ]);
  });

  it('adds a rectangle as primary and removes overlap from prior ranges', () => {
    const base = selection(range(0, 0, 2, 2), [range(4, 4, 5, 5)]);
    const next = applySelectionRectangle(
      base,
      range(2, 2, 4, 4),
      'add',
      {
        sheet: 0,
        row: 2,
        col: 2,
      },
      {
        sheet: 0,
        row: 4,
        col: 4,
      },
    );
    if (!next) throw new Error('Expected add gesture to retain a selected range');
    expect(next?.range).toEqual(range(2, 2, 4, 4));
    expect(next?.extraRanges).toEqual([
      range(5, 4, 5, 5),
      range(4, 5),
      range(0, 0, 1, 2),
      range(2, 0, 2, 1),
    ]);
    expect(selectionContainsAddr(next, { sheet: 0, row: 4, col: 4 })).toBe(true);
    expect(selectionContainsAddr(next, { sheet: 0, row: 4, col: 5 })).toBe(true);
    expect(selectionContainsAddr(next, { sheet: 0, row: 2, col: 2 })).toBe(true);
  });

  it('promotes the surviving active fragment and keeps its anchor only there', () => {
    const base = selection(range(0, 0, 4, 4));
    base.active = { sheet: 0, row: 2, col: 0 };
    base.anchor = { sheet: 0, row: 0, col: 0 };
    const next = applySelectionRectangle(
      base,
      range(1, 0, 3, 4),
      'subtract',
      {
        sheet: 0,
        row: 1,
        col: 0,
      },
      {
        sheet: 0,
        row: 3,
        col: 4,
      },
    );
    expect(next?.range).toEqual(range(0, 0, 0, 4));
    expect(next?.active).toEqual({ sheet: 0, row: 0, col: 0 });
    expect(next?.anchor).toEqual(next?.active);
    expect(next?.extraRanges).toEqual([range(4, 0, 4, 4)]);
  });

  it('keeps an active address inside a surviving fragment and prevents empty selections', () => {
    const base = selection(range(0, 0, 2, 2));
    base.active = { sheet: 0, row: 2, col: 2 };
    base.anchor = { sheet: 0, row: 2, col: 2 };
    const next = applySelectionRectangle(
      base,
      range(0, 0, 1, 2),
      'subtract',
      {
        sheet: 0,
        row: 0,
        col: 0,
      },
      {
        sheet: 0,
        row: 1,
        col: 2,
      },
    );
    expect(next?.range).toEqual(range(2, 0, 2, 2));
    expect(next?.active).toEqual(base.active);
    expect(
      applySelectionRectangle(
        selection(range(0, 0)),
        range(0, 0),
        'subtract',
        {
          sheet: 0,
          row: 0,
          col: 0,
        },
        {
          sheet: 0,
          row: 0,
          col: 0,
        },
      ),
    ).toBeNull();
  });

  it('tests address membership and coverage across the union of ranges', () => {
    const selected = selection(range(0, 0, 0, 1), [range(0, 2, 0, 3)]);
    expect(rangeContainsAddr(range(0, 0, 1, 1), { sheet: 0, row: 1, col: 1 })).toBe(true);
    expect(selectionContainsAddr(selected, { sheet: 0, row: 0, col: 3 })).toBe(true);
    expect(selectionCoversRange(selected, range(0, 0, 0, 3))).toBe(true);
    expect(selectionCoversRange(selected, range(0, 0, 1, 3))).toBe(false);
  });

  describe('range primitives', () => {
    it('intersects only ranges on the same sheet that share a cell', () => {
      expect(rangesIntersect(range(0, 0, 2, 2), range(2, 2, 4, 4))).toBe(true);
      expect(rangesIntersect(range(0, 0, 2, 2), range(3, 0, 4, 2))).toBe(false);
      expect(rangesIntersect(range(0, 0, 2, 2), range(0, 0, 2, 2, 1))).toBe(false);
    });

    it('contains a range only when every edge lies inside, on the same sheet', () => {
      expect(rangeContainsRange(range(0, 0, 4, 4), range(1, 1, 4, 4))).toBe(true);
      expect(rangeContainsRange(range(0, 0, 4, 4), range(1, 1, 5, 4))).toBe(false);
      expect(rangeContainsRange(range(0, 0, 4, 4), range(1, 1, 2, 2, 1))).toBe(false);
    });

    it('compares every coordinate and the sheet for equality', () => {
      expect(sameRange(range(1, 2, 3, 4), range(1, 2, 3, 4))).toBe(true);
      expect(sameRange(range(1, 2, 3, 4), range(1, 2, 3, 5))).toBe(false);
      expect(sameRange(range(1, 2, 3, 4), range(1, 2, 3, 4, 1))).toBe(false);
    });

    it('counts cells and treats an inverted axis as empty', () => {
      expect(rangeArea(range(0, 0))).toBe(1);
      expect(rangeArea(range(0, 0, 2, 3))).toBe(12);
      expect(rangeArea(range(2, 0, 1, 3))).toBe(0);
      expect(rangeArea(range(5, 5, 0, 0))).toBe(0);
    });
  });
});
