import { addrKey, rangeKey } from '../../engine/address.js';
import type { Addr, Range } from '../../engine/types.js';
import { borderSideSignature } from '../../format/border-signature.js';
import type { CellBorderSide, CellFormat, State } from '../../store/store.js';

export const mergeRangeAt = (state: State, addr: Addr): Range | null => {
  const key = addrKey(addr);
  const anchorKey = state.merges.byCell.get(key) ?? key;
  return state.merges.byAnchor.get(anchorKey) ?? null;
};

interface DiagonalSummary {
  coordinates: string;
  borders: CellFormat['borders'];
}

const diagonalSummaries = new WeakMap<
  ReadonlyMap<string, CellFormat>,
  WeakMap<Range, DiagonalSummary>
>();

/** Excel stores a uniform diagonal border on every cell covered by a merge,
 *  but renders that border once across the merged surface. Mixed diagonal
 *  metadata is discarded by Excel at merge time; suppress it here as well so
 *  stale per-cell data cannot produce crossing lines inside a merge. */
export const mergedDiagonalBorders = (
  formats: ReadonlyMap<string, CellFormat>,
  merge: Range,
): CellFormat['borders'] => {
  let byMerge = diagonalSummaries.get(formats);
  if (!byMerge) {
    byMerge = new WeakMap();
    diagonalSummaries.set(formats, byMerge);
  }
  const coordinates = rangeKey(merge);
  const cached = byMerge.get(merge);
  if (cached?.coordinates === coordinates) return cached.borders;

  let down: CellBorderSide | undefined;
  let downSignature: string | null = null;
  let up: CellBorderSide | undefined;
  let upSignature: string | null = null;
  for (let row = merge.r0; row <= merge.r1; row += 1) {
    for (let col = merge.c0; col <= merge.c1; col += 1) {
      const borders = formats.get(addrKey({ sheet: merge.sheet, row, col }))?.borders;
      const nextDown = borders?.diagonalDown;
      const nextUp = borders?.diagonalUp;
      const nextDownSignature = borderSideSignature(nextDown);
      const nextUpSignature = borderSideSignature(nextUp);
      if (downSignature === null) {
        down = nextDown;
        downSignature = nextDownSignature;
      } else if (downSignature !== 'mixed' && downSignature !== nextDownSignature) {
        downSignature = 'mixed';
      }
      if (upSignature === null) {
        up = nextUp;
        upSignature = nextUpSignature;
      } else if (upSignature !== 'mixed' && upSignature !== nextUpSignature) {
        upSignature = 'mixed';
      }
    }
  }
  const result: NonNullable<CellFormat['borders']> = {};
  if (downSignature !== null && downSignature !== 'mixed' && downSignature !== 'absent') {
    result.diagonalDown = down;
  }
  if (upSignature !== null && upSignature !== 'mixed' && upSignature !== 'absent') {
    result.diagonalUp = up;
  }
  const borders = Object.keys(result).length > 0 ? result : undefined;
  byMerge.set(merge, { coordinates, borders });
  return borders;
};
