import type { LayoutSlice } from './types.js';

/** First row of the frozen band. Rows above it are unreachable while frozen. */
export function frozenRowStart(layout: LayoutSlice): number {
  return layout.freezeRows > 0 ? (layout.freezeOriginRow ?? 0) : 0;
}

/** One past the last frozen row; the scrollable body starts here or later. */
export function frozenRowEnd(layout: LayoutSlice): number {
  return layout.freezeRows > 0 ? frozenRowStart(layout) + layout.freezeRows : 0;
}

/** First column of the frozen band. Columns left of it are unreachable while frozen. */
export function frozenColStart(layout: LayoutSlice): number {
  return layout.freezeCols > 0 ? (layout.freezeOriginCol ?? 0) : 0;
}

/** One past the last frozen column; the scrollable body starts here or later. */
export function frozenColEnd(layout: LayoutSlice): number {
  return layout.freezeCols > 0 ? frozenColStart(layout) + layout.freezeCols : 0;
}
