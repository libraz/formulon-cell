import type { LayoutSlice } from './types.js';

type RowVisibilityLayout = Pick<LayoutSlice, 'hiddenRows' | 'defaultRowsHidden' | 'shownRows'>;

/** Membership test over hidden row or column indices. */
export interface HiddenIndices {
  has(index: number): boolean;
}

/** Whether `row` is hidden: explicitly, or by a sheet that hides rows by default. */
export const isRowHidden = (layout: RowVisibilityLayout, row: number): boolean =>
  layout.hiddenRows.has(row) || (layout.defaultRowsHidden === true && !layout.shownRows?.has(row));

/** Every hidden row of `layout`, including the ones hidden by default. */
export const hiddenRowsOf = (layout: RowVisibilityLayout): HiddenIndices =>
  layout.defaultRowsHidden === true
    ? { has: (row) => isRowHidden(layout, row) }
    : layout.hiddenRows;

/** `lo..hi` split into runs of visible rows. */
export function visibleRowSegments(
  layout: RowVisibilityLayout,
  lo: number,
  hi: number,
): [number, number][] {
  const out: [number, number][] = [];
  const candidates =
    layout.defaultRowsHidden === true
      ? [...(layout.shownRows ?? [])].filter((r) => r >= lo && r <= hi).sort((a, b) => a - b)
      : null;
  if (!candidates) {
    const skipped = [...layout.hiddenRows].filter((i) => i >= lo && i <= hi).sort((a, b) => a - b);
    let start = lo;
    for (const i of skipped) {
      if (i > start) out.push([start, i - 1]);
      start = i + 1;
    }
    if (start <= hi) out.push([start, hi]);
    return out;
  }
  for (const r of candidates) {
    if (layout.hiddenRows.has(r)) continue;
    const last = out.at(-1);
    if (last && last[1] === r - 1) last[1] = r;
    else out.push([r, r]);
  }
  return out;
}
