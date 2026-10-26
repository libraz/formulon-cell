import type { LayoutSlice, OutlineSummary } from './types.js';

/** The outline summary position of a sheet that states none. */
export const DEFAULT_OUTLINE_SUMMARY: Readonly<OutlineSummary> = { below: true, right: true };

/** Outline summary position of `sheet`, falling back to the defaults. */
export const outlineSummaryFor = (
  layout: Pick<LayoutSlice, 'outlineSummary'>,
  sheet: number,
): Readonly<OutlineSummary> => layout.outlineSummary?.get(sheet) ?? DEFAULT_OUTLINE_SUMMARY;

/** Width of one bracket slot in CSS pixels. The gutter widens by this amount
 *  per outline level so each level has its own visual track. */
export const OUTLINE_GUTTER_PER_LEVEL = 14;

/** Gutter width an axis's outline levels call for. */
export const outlineGutterWidth = (levels: ReadonlyMap<number, number>): number => {
  let max = 0;
  for (const v of levels.values()) if (v > max) max = v;
  return max * OUTLINE_GUTTER_PER_LEVEL;
};
