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

/** Mac lane width in CSS px: Excel for Mac draws 24 device px per level at 2x. */
export const MAC_OUTLINE_LANE = 12;

/** Mac gutter margin past the lanes, CSS px (6 device px at 2x). */
const MAC_OUTLINE_MARGIN = 3;

/** Gutter width an axis's outline levels call for. On the Mac point model
 *  (screen scale other than 1) Excel reserves one lane more than the deepest
 *  level, for the detail dots and the level buttons, plus a small margin. */
export const outlineGutterWidth = (
  levels: ReadonlyMap<number, number>,
  layout: Pick<LayoutSlice, 'screenScale'>,
): number => {
  let max = 0;
  for (const v of levels.values()) if (v > max) max = v;
  if (max === 0) return 0;
  return (layout.screenScale ?? 1) !== 1
    ? (max + 1) * MAC_OUTLINE_LANE + MAC_OUTLINE_MARGIN
    : max * OUTLINE_GUTTER_PER_LEVEL;
};
