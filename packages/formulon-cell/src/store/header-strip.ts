import { MAC_COLUMN_PADDING_PT, roundDigitAdvance } from '../engine/mac-column-model.js';
import type { LayoutSlice, ViewportSlice } from './types.js';

/** Fewest digits the row-number strip is sized for. */
const MIN_ROW_HEADER_DIGITS = 3;

type HeaderLayout = Pick<
  LayoutSlice,
  'headerColWidth' | 'headerRowHeight' | 'headerDigitAdvance' | 'defaultRowHeight' | 'screenScale'
>;

/** Header strip sizes in screen px for the rows in view at the view's zoom.
 *  With a Normal-font digit advance (Mac), the column-letter strip is the
 *  larger of the chrome height and one default row, scaled by the zoom and
 *  rounded up; the row-number strip is a column as many zoomed digits wide as
 *  the last visible row number (at least three) plus the column padding, and
 *  never narrower than the column strip is tall (Excel 75-200%). Otherwise
 *  the fixed chrome sizes stand. */
export function headerStripSize(
  layout: HeaderLayout,
  viewport: Pick<ViewportSlice, 'rowStart' | 'rowCount'> & { zoom?: number },
): Pick<LayoutSlice, 'headerColWidth' | 'headerRowHeight'> {
  const advance = layout.headerDigitAdvance;
  if (advance === undefined) {
    return { headerColWidth: layout.headerColWidth, headerRowHeight: layout.headerRowHeight };
  }
  const zoom = viewport.zoom !== undefined && viewport.zoom > 0 ? viewport.zoom : 1;
  const baseHeight = Math.max(
    layout.headerRowHeight,
    layout.defaultRowHeight * (layout.screenScale ?? 1),
  );
  // Guard the ceiling against float noise in a product like 22 x 1.1.
  const headerRowHeight = Math.ceil(baseHeight * zoom - 1e-9);
  const lastRowNumber = viewport.rowStart + viewport.rowCount;
  const digits = Math.max(MIN_ROW_HEADER_DIGITS, String(lastRowNumber).length);
  const digitWidth = roundDigitAdvance(advance * zoom);
  return {
    headerColWidth: Math.max(headerRowHeight, MAC_COLUMN_PADDING_PT + digits * digitWidth),
    headerRowHeight,
  };
}
