/** Row height ceiling: 409.5 pt in canonical CSS px. */
export const MAX_ROW_HEIGHT_PX = (409.5 * 96) / 72;

/** Quarter points per canonical CSS px: 4 per pt at 0.75 pt per px. */
const QUARTER_POINTS_PER_PX = 3;

/** Snap a row height in canonical CSS px to a quarter point, the unit the
 *  desktop spreadsheet stores; painting snaps each row to screen px apart. */
export const snapRowHeightPx = (px: number): number =>
  Math.round(px * QUARTER_POINTS_PER_PX) / QUARTER_POINTS_PER_PX;
