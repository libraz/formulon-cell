import { cloneAddr, MAX_COL, MAX_ROW } from '../engine/address.js';
import type { Addr, Range } from '../engine/types.js';
import type { SelectionSlice } from './types.js';

export type SelectionGestureMode = 'add' | 'subtract';

/** Whether `range` is a non-inverted, in-grid range on a valid sheet index. */
export const isValidRange = (range: Range): boolean =>
  Number.isInteger(range.sheet) &&
  Number.isInteger(range.r0) &&
  Number.isInteger(range.c0) &&
  Number.isInteger(range.r1) &&
  Number.isInteger(range.c1) &&
  range.sheet >= 0 &&
  range.r0 >= 0 &&
  range.c0 >= 0 &&
  range.r0 <= range.r1 &&
  range.c0 <= range.c1 &&
  range.r1 <= MAX_ROW &&
  range.c1 <= MAX_COL;

/** A well-formed, in-bounds range on `sheet`. */
export const isValidSheetRange = (range: Range, sheet: number): boolean =>
  range.sheet === sheet && isValidRange(range);

export const cloneRange = (range: Range): Range => ({
  sheet: range.sheet,
  r0: range.r0,
  c0: range.c0,
  r1: range.r1,
  c1: range.c1,
});

/** The one-cell range covering `addr`. */
export const rangeOfAddr = (addr: Addr): Range => ({
  sheet: addr.sheet,
  r0: addr.row,
  c0: addr.col,
  r1: addr.row,
  c1: addr.col,
});

/** `range` with its corners ordered, so `r0 <= r1` and `c0 <= c1`. */
export const normalizeRange = (range: Range): Range => ({
  sheet: range.sheet,
  r0: Math.min(range.r0, range.r1),
  c0: Math.min(range.c0, range.c1),
  r1: Math.max(range.r0, range.r1),
  c1: Math.max(range.c0, range.c1),
});

/** True when `range` covers one cell. */
export const isSingleCell = (range: Pick<Range, 'r0' | 'c0' | 'r1' | 'c1'>): boolean =>
  range.r0 === range.r1 && range.c0 === range.c1;

/** The primary range of a selection followed by its added ranges. */
export const selectionRangeList = (
  selection: Pick<SelectionSlice, 'range' | 'extraRanges'>,
): Range[] => [selection.range, ...(selection.extraRanges ?? [])];

/** True when `range` spans every column, i.e. it selects whole rows. */
export const isWholeRowRange = (range: Pick<Range, 'c0' | 'c1'>): boolean =>
  range.c0 === 0 && range.c1 >= MAX_COL;

/** True when `range` spans every row, i.e. it selects whole columns. */
export const isWholeColumnRange = (range: Pick<Range, 'r0' | 'r1'>): boolean =>
  range.r0 === 0 && range.r1 >= MAX_ROW;

export const sameRange = (a: Range, b: Range): boolean =>
  a.sheet === b.sheet && a.r0 === b.r0 && a.c0 === b.c0 && a.r1 === b.r1 && a.c1 === b.c1;

export const sameRangeList = (a: readonly Range[], b: readonly Range[]): boolean =>
  a.length === b.length && a.every((range, i) => sameRange(range, b[i] as Range));

export const rangeContainsAddr = (range: Range, addr: Addr): boolean =>
  range.sheet === addr.sheet &&
  addr.row >= range.r0 &&
  addr.row <= range.r1 &&
  addr.col >= range.c0 &&
  addr.col <= range.c1;

export const rangeContainsRange = (outer: Range, inner: Range): boolean =>
  outer.sheet === inner.sheet &&
  outer.r0 <= inner.r0 &&
  outer.c0 <= inner.c0 &&
  outer.r1 >= inner.r1 &&
  outer.c1 >= inner.c1;

export function rangesIntersect(a: Range, b: Range): boolean {
  return a.sheet === b.sheet && !(a.r1 < b.r0 || a.r0 > b.r1 || a.c1 < b.c0 || a.c0 > b.c1);
}

/** Cell count of `range`; 0 when either axis is inverted. */
export const rangeArea = (range: Range): number =>
  range.r1 < range.r0 || range.c1 < range.c0
    ? 0
    : (range.r1 - range.r0 + 1) * (range.c1 - range.c0 + 1);

export const selectionContainsAddr = (
  selection: Pick<SelectionSlice, 'range' | 'extraRanges'>,
  addr: Addr,
): boolean =>
  rangeContainsAddr(selection.range, addr) ||
  (selection.extraRanges ?? []).some((range) => rangeContainsAddr(range, addr));

const intersection = (a: Range, b: Range): Range | null => {
  if (a.sheet !== b.sheet || a.r1 < b.r0 || a.r0 > b.r1 || a.c1 < b.c0 || a.c0 > b.c1) {
    return null;
  }
  return {
    sheet: a.sheet,
    r0: Math.max(a.r0, b.r0),
    c0: Math.max(a.c0, b.c0),
    r1: Math.min(a.r1, b.r1),
    c1: Math.min(a.c1, b.c1),
  };
};

/** Return disjoint pieces of `range` not covered by `cover`, without visiting cells. */
export const subtractRange = (range: Range, cover: Range): Range[] => {
  const overlap = intersection(range, cover);
  if (!overlap) return [cloneRange(range)];

  const pieces: Range[] = [];
  if (range.r0 < overlap.r0) pieces.push({ ...range, r1: overlap.r0 - 1 });
  if (overlap.r1 < range.r1) pieces.push({ ...range, r0: overlap.r1 + 1 });
  if (range.c0 < overlap.c0) {
    pieces.push({ ...range, r0: overlap.r0, r1: overlap.r1, c1: overlap.c0 - 1 });
  }
  if (overlap.c1 < range.c1) {
    pieces.push({ ...range, r0: overlap.r0, r1: overlap.r1, c0: overlap.c1 + 1 });
  }
  return pieces;
};

/** Test range coverage by progressively removing each selected rectangle. */
export const selectionCoversRange = (
  selection: Pick<SelectionSlice, 'range' | 'extraRanges'>,
  target: Range,
): boolean => {
  let uncovered = [cloneRange(target)];
  for (const selected of selectionRangeList(selection)) {
    uncovered = uncovered.flatMap((piece) => subtractRange(piece, selected));
    if (uncovered.length === 0) return true;
  }
  return false;
};

interface SelectionRangePart {
  range: Range;
  source: number;
  fragment: number;
}

const nearestAddrInRange = (addr: Addr, range: Range): Addr => ({
  sheet: range.sheet,
  row: Math.max(range.r0, Math.min(range.r1, addr.row)),
  col: Math.max(range.c0, Math.min(range.c1, addr.col)),
});

const distanceToAddr = (from: Addr, to: Addr): number =>
  Math.abs(from.row - to.row) + Math.abs(from.col - to.col);

/** Apply a rectangular add/subtract gesture to a frozen base selection. */
export const applySelectionRectangle = (
  base: SelectionSlice,
  gesture: Range,
  mode: SelectionGestureMode,
  anchor: Addr,
): SelectionSlice | null => {
  const ranges = selectionRangeList(base);

  if (mode === 'add') {
    const extras = [...(base.extraRanges ?? []), base.range].flatMap((range) =>
      subtractRange(range, gesture),
    );
    return {
      active: cloneAddr(anchor),
      anchor: cloneAddr(anchor),
      range: cloneRange(gesture),
      extraRanges: extras,
    };
  }

  const survivors: SelectionRangePart[] = [];
  ranges.forEach((range, source) => {
    subtractRange(range, gesture).forEach((fragment, index) => {
      survivors.push({ range: fragment, source, fragment: index });
    });
  });
  if (survivors.length === 0) return null;

  let primaryIndex = survivors.findIndex(({ range }) => rangeContainsAddr(range, base.active));
  let active = cloneAddr(base.active);
  if (primaryIndex < 0) {
    let bestIndex = -1;
    let bestAddr: Addr | null = null;
    let bestSameSheet = false;
    let bestDistance = Number.POSITIVE_INFINITY;
    survivors.forEach(({ range }, index) => {
      const candidate = nearestAddrInRange(base.active, range);
      const sameSheet = candidate.sheet === base.active.sheet;
      const distance = distanceToAddr(base.active, candidate);
      if (
        bestIndex < 0 ||
        (sameSheet && !bestSameSheet) ||
        (sameSheet === bestSameSheet && distance < bestDistance)
      ) {
        bestIndex = index;
        bestAddr = candidate;
        bestSameSheet = sameSheet;
        bestDistance = distance;
      }
    });
    if (bestIndex < 0 || !bestAddr) return null;
    primaryIndex = bestIndex;
    active = cloneAddr(bestAddr);
  }

  // The source and fragment order is the stable tie-break when multiple legacy
  // ranges overlap and the same active coordinate survives in more than one.
  survivors.sort((a, b) => a.source - b.source || a.fragment - b.fragment);
  primaryIndex = survivors.findIndex(({ range }) => rangeContainsAddr(range, active));
  if (primaryIndex < 0) return null;
  const primary = survivors[primaryIndex]?.range;
  if (!primary) return null;
  const keptAnchor = rangeContainsAddr(primary, base.anchor) ? base.anchor : active;

  return {
    active: cloneAddr(active),
    anchor: cloneAddr(keptAnchor),
    range: cloneRange(primary),
    extraRanges: survivors
      .filter((_, index) => index !== primaryIndex)
      .map(({ range }) => cloneRange(range)),
  };
};

export const fullSheetRange = (sheet: number): Range => ({
  sheet,
  r0: 0,
  c0: 0,
  r1: MAX_ROW,
  c1: MAX_COL,
});

export const cloneSelection = (selection: SelectionSlice): SelectionSlice => ({
  active: { ...selection.active },
  anchor: { ...selection.anchor },
  range: { ...selection.range },
  extraRanges: (selection.extraRanges ?? []).map((range) => ({ ...range })),
});

export const axisSelectionRange = (
  bounds: Range,
  axis: 'row' | 'col',
  start: number,
  tip: number,
): Range =>
  axis === 'row'
    ? {
        sheet: bounds.sheet,
        r0: Math.min(start, tip),
        c0: bounds.c0,
        r1: Math.max(start, tip),
        c1: bounds.c1,
      }
    : {
        sheet: bounds.sheet,
        r0: bounds.r0,
        c0: Math.min(start, tip),
        r1: bounds.r1,
        c1: Math.max(start, tip),
      };
