import type { LayoutSlice, OutlineSummary } from './types.js';

/** The outline summary position of a sheet that states none. */
export const DEFAULT_OUTLINE_SUMMARY: Readonly<OutlineSummary> = { below: true, right: true };

/** Outline summary position of `sheet`, falling back to the defaults. */
export const outlineSummaryFor = (
  layout: Pick<LayoutSlice, 'outlineSummary'>,
  sheet: number,
): Readonly<OutlineSummary> => layout.outlineSummary?.get(sheet) ?? DEFAULT_OUTLINE_SUMMARY;
