import { isRowHidden } from './row-visibility.js';
import type { LayoutSlice, State, UiSlice } from './types.js';

/** OOXML stores a split in twips of screen space: 20 per point, 15 per CSS px. */
const TWIPS_PER_CSS_PX = 15;

type SplitAxis = 'row' | 'col';
type SplitLayout = Pick<
  LayoutSlice,
  | 'rowHeights'
  | 'colWidths'
  | 'defaultRowHeight'
  | 'defaultColWidth'
  | 'hiddenRows'
  | 'defaultRowsHidden'
  | 'shownRows'
  | 'hiddenCols'
  | 'headerRowHeight'
  | 'headerColWidth'
  | 'outlineRowGutter'
  | 'outlineColGutter'
  | 'screenScale'
>;
type SplitUi = Pick<UiSlice, 'showHeaders' | 'showOutlineSymbols'>;

const sizeAt = (layout: SplitLayout, axis: SplitAxis, i: number): number =>
  axis === 'row'
    ? isRowHidden(layout, i)
      ? 0
      : (layout.rowHeights.get(i) ?? layout.defaultRowHeight)
    : layout.hiddenCols.has(i)
      ? 0
      : (layout.colWidths.get(i) ?? layout.defaultColWidth);

/** The header strip and outline gutter a split's position is measured across,
 *  in canonical px. They draw in screen px and do not zoom, but the file
 *  measures them zoomed, as the desktop spreadsheet draws them. */
const headerExtent = (layout: SplitLayout, ui: SplitUi, axis: SplitAxis): number => {
  const header =
    ui.showHeaders === false ? 0 : axis === 'row' ? layout.headerRowHeight : layout.headerColWidth;
  const gutter =
    ui.showOutlineSymbols === false
      ? 0
      : axis === 'row'
        ? layout.outlineColGutter
        : layout.outlineRowGutter;
  return (header + gutter) / (layout.screenScale ?? 1);
};

/** Twips position of a split whose top/left pane shows `count` rows or columns from `origin`. */
export function splitTwips(
  layout: SplitLayout,
  ui: SplitUi,
  axis: SplitAxis,
  origin: number,
  count: number,
  zoom: number,
): number {
  if (count <= 0) return 0;
  let px = headerExtent(layout, ui, axis);
  for (let i = origin; i < origin + count; i += 1) px += sizeAt(layout, axis, i);
  return Math.round(px * zoom * TWIPS_PER_CSS_PX);
}

/** Rows or columns from `origin` a saved split covers, to the nearest boundary.
 *  0 when the split sits within the headers. */
export function splitCount(
  layout: SplitLayout,
  ui: SplitUi,
  axis: SplitAxis,
  origin: number,
  twips: number,
  zoom: number,
  limit: number,
): number {
  if (!(twips > 0) || !(zoom > 0)) return 0;
  const target = twips / TWIPS_PER_CSS_PX / zoom - headerExtent(layout, ui, axis);
  let px = 0;
  let count = 0;
  while (origin + count < limit) {
    const size = sizeAt(layout, axis, origin + count);
    if (px + size / 2 > target) break;
    px += size;
    count += 1;
  }
  return count;
}

/** Where the active sheet's split sits, in twips; null unless the band is split panes. */
export function splitPosition(
  state: Pick<State, 'layout' | 'ui' | 'viewport'>,
): { x: number; y: number } | null {
  const { layout, ui, viewport } = state;
  if (layout.splitPanes !== true) return null;
  return {
    x: splitTwips(layout, ui, 'col', layout.freezeOriginCol ?? 0, layout.freezeCols, viewport.zoom),
    y: splitTwips(layout, ui, 'row', layout.freezeOriginRow ?? 0, layout.freezeRows, viewport.zoom),
  };
}

/** Where a top/left pane of `count` rows or columns may start: inside the
 *  sheet and ending at or before `bodyStart`, the first index of the pane past it. */
export const clampSplitPaneOrigin = (count: number, bodyStart: number, origin: number): number =>
  count > 0 ? Math.min(Math.max(0, bodyStart - count), Math.max(0, origin)) : origin;
