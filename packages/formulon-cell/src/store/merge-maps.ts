/** Mutations of the two merge lookup maps (`byAnchor`, `byCell`) shared by every
 *  merge writer. Kept free of runtime store imports so command modules can use it
 *  without entering the store's module cycle. */
import { addrKey } from '../engine/address.js';
import type { Range } from '../engine/types.js';
import { rangesIntersect } from './selection-geometry.js';

/** Register `range` in both merge lookup maps (anchor → range, covered cell → anchor key). */
export function addMergeToMaps(
  byAnchor: Map<string, Range>,
  byCell: Map<string, string>,
  range: Range,
): void {
  const ak = addrKey({ sheet: range.sheet, row: range.r0, col: range.c0 });
  byAnchor.set(ak, range);
  for (let row = range.r0; row <= range.r1; row += 1) {
    for (let col = range.c0; col <= range.c1; col += 1) {
      if (row === range.r0 && col === range.c0) continue;
      byCell.set(addrKey({ sheet: range.sheet, row, col }), ak);
    }
  }
}

/** Drop every merge intersecting `range` from both merge lookup maps. */
export function removeIntersectingMerges(
  byAnchor: Map<string, Range>,
  byCell: Map<string, string>,
  range: Range,
): void {
  for (const [anchorKey, merge] of byAnchor) {
    if (!rangesIntersect(merge, range)) continue;
    byAnchor.delete(anchorKey);
    for (let row = merge.r0; row <= merge.r1; row += 1) {
      for (let col = merge.c0; col <= merge.c1; col += 1) {
        byCell.delete(addrKey({ sheet: merge.sheet, row, col }));
      }
    }
  }
}
