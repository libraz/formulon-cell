import { formatA1Range } from '../engine/address.js';
import { parseRangeRef } from '../engine/range-resolver.js';
import type { Range } from '../engine/types.js';
import { splitSourceQualifier } from './sparkline.js';

type Block = Pick<Range, 'r0' | 'c0' | 'r1' | 'c1'>;

const cellCount = (range: Block): number => (range.r1 - range.r0 + 1) * (range.c1 - range.c0 + 1);

/** The data range of each location cell, in row-major order: the rows of `data` for a location
 *  that is not one row, its columns for a single-row location. Null when neither dimension of
 *  `data` matches the number of location cells. A single location cell takes all of `data`. */
export function splitDataForLocation(data: Range, location: Block): Range[] | null {
  const count = cellCount(location);
  if (count === 1) return [{ ...data }];
  const rows = data.r1 - data.r0 + 1;
  const cols = data.c1 - data.c0 + 1;
  const byColumn = rows !== count || (location.r0 === location.r1 && cols === count);
  if (byColumn ? cols !== count : rows !== count) return null;
  return Array.from({ length: count }, (_, i) =>
    byColumn
      ? { ...data, c0: data.c0 + i, c1: data.c0 + i }
      : { ...data, r0: data.r0 + i, r1: data.r0 + i },
  );
}

const parseBody = (body: string): Block | null => {
  const parsed = parseRangeRef(body);
  return parsed && parsed.sheetName === null ? parsed : null;
};

/** Re-split the block the sources of a group tile along the other axis, as Switch Row/Column
 *  does: sources that are consecutive single rows of one block become its columns, and
 *  consecutive single columns become its rows. Null when the sources do not tile one block or
 *  the source count differs from the other dimension of the block. */
export function switchSparklineSources(sources: readonly string[]): string[] | null {
  if (sources.length < 2) return null;
  const qualifier = splitSourceQualifier(sources[0] ?? '').qualifier;
  const blocks: Block[] = [];
  for (const source of sources) {
    const split = splitSourceQualifier(source);
    const block = split.qualifier === qualifier ? parseBody(split.body) : null;
    if (!block) return null;
    blocks.push(block);
  }
  const first = blocks[0] as Block;
  const count = blocks.length;
  const tiles = (axis: 'row' | 'col'): boolean =>
    blocks.every((block, i) =>
      axis === 'row'
        ? block.r0 === block.r1 &&
          block.r0 === first.r0 + i &&
          block.c0 === first.c0 &&
          block.c1 === first.c1
        : block.c0 === block.c1 &&
          block.c0 === first.c0 + i &&
          block.r0 === first.r0 &&
          block.r1 === first.r1,
    );
  const format = (range: Block): string => `${qualifier}${formatA1Range(range)}`;
  if (tiles('row')) {
    if (first.c1 - first.c0 + 1 !== count) return null;
    return blocks.map((_, i) =>
      format({ r0: first.r0, r1: first.r0 + count - 1, c0: first.c0 + i, c1: first.c0 + i }),
    );
  }
  if (tiles('col')) {
    if (first.r1 - first.r0 + 1 !== count) return null;
    return blocks.map((_, i) =>
      format({ c0: first.c0, c1: first.c0 + count - 1, r0: first.r0 + i, r1: first.r0 + i }),
    );
  }
  return null;
}

/** The data range a group reads as one: the bounding block of its sources when they share a
 *  sheet qualifier and all parse, else the first source. */
export function unionSparklineSources(sources: readonly string[]): string {
  const first = sources[0] ?? '';
  const qualifier = splitSourceQualifier(first).qualifier;
  let union: Block | null = null;
  for (const source of sources) {
    const split = splitSourceQualifier(source);
    const block = split.qualifier === qualifier ? parseBody(split.body) : null;
    if (!block) return first;
    union = union
      ? {
          r0: Math.min(union.r0, block.r0),
          c0: Math.min(union.c0, block.c0),
          r1: Math.max(union.r1, block.r1),
          c1: Math.max(union.c1, block.c1),
        }
      : block;
  }
  return union ? `${qualifier}${formatA1Range(union)}` : first;
}
