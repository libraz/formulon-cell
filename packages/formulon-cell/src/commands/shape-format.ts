import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type {
  SessionIllustration,
  SessionShapeLineDash,
  SpreadsheetStore,
} from '../store/store.js';
import type { History } from './history.js';
import {
  type SessionIllustrationPatch,
  updateSessionIllustration,
} from './session-illustration.js';

/** The fill of a shape that has not been given one: accent 1 of the default theme. */
export const DEFAULT_SHAPE_COLOR = '#156082';

/** The outline of a filled shape that has not been given one, as the spreadsheet draws it. */
export const DEFAULT_SHAPE_OUTLINE_COLOR = '#042433';

/** The outline weight of a new shape, in points. */
export const DEFAULT_SHAPE_LINE_WIDTH = 1.5;

/** The heaviest outline the spreadsheet accepts, in points. */
export const MAX_SHAPE_LINE_WIDTH = 1584;

/** The Weight entries of Shape Outline, in points. */
export const SHAPE_LINE_WEIGHTS: readonly number[] = [0.25, 0.5, 0.75, 1, 1.5, 2.25, 3, 4.5, 6];

/** The Dashes entries of Shape Outline, in menu order. */
export const SHAPE_LINE_DASHES: readonly SessionShapeLineDash[] = [
  'solid',
  'roundDot',
  'squareDot',
  'dash',
  'dashDot',
  'longDash',
  'longDashDot',
  'longDashDotDot',
];

/** Each dash style's dash and gap lengths, in multiples of the line width. */
const DASH_PATTERNS: Record<SessionShapeLineDash, readonly number[]> = {
  solid: [],
  roundDot: [0, 2],
  squareDot: [1, 1],
  dash: [4, 3],
  dashDot: [4, 3, 1, 3],
  longDash: [8, 3],
  longDashDot: [8, 3, 1, 3],
  longDashDotDot: [8, 3, 1, 3, 1, 3],
};

/** The SVG stroke of `dash` at `width` pixels: its dash array (none for solid)
 *  and its line cap. A round dot is a zero-length dash with a round cap. */
export const shapeDashStroke = (
  dash: SessionShapeLineDash,
  width: number,
): { dasharray: string; linecap: 'round' | 'butt' } => ({
  dasharray: DASH_PATTERNS[dash].map((step) => step * width).join(' ') || 'none',
  linecap: dash === 'roundDot' ? 'round' : 'butt',
});

/** Lines and arrows are all outline; they have no fill. */
export const shapeHasFill = (item: SessionIllustration): boolean =>
  item.kind === 'shape' && item.shape !== 'line' && item.shape !== 'arrow';

/** The corner radius of a `w` x `h` rectangle, in pixels: a rounded
 *  rectangle's corner is a sixth of its shorter side. */
export const shapeCornerRadius = (item: SessionIllustration, w: number, h: number): number =>
  item.radius ?? (item.shape === 'rounded-rectangle' ? Math.min(w, h) / 6 : 0);

/** The fill colour of a shape, or null for No Fill. */
export const shapeFillColor = (item: SessionIllustration): string | null =>
  shapeHasFill(item) && item.opacity !== 0 ? (item.color ?? DEFAULT_SHAPE_COLOR) : null;

/** The outline colour of a shape, or null for No Outline. A line that has not
 *  been given an outline colour is drawn in its own colour. */
export const shapeOutlineColor = (item: SessionIllustration): string | null => {
  if (item.lineWidth === 0) return null;
  if (item.lineColor) return item.lineColor;
  return shapeHasFill(item) ? DEFAULT_SHAPE_OUTLINE_COLOR : (item.color ?? DEFAULT_SHAPE_COLOR);
};

/** The outline weight of a shape in points, or null for No Outline. */
export const shapeOutlineWeight = (item: SessionIllustration): number | null =>
  item.lineWidth === 0 ? null : (item.lineWidth ?? DEFAULT_SHAPE_LINE_WIDTH);

/** The dash style of a shape's outline. */
export const shapeOutlineDash = (item: SessionIllustration): SessionShapeLineDash =>
  item.lineDash ?? 'solid';

/** The patch that gives the shape `color` as its fill, or No Fill when null. */
export const shapeFillPatch = (
  _item: SessionIllustration,
  color: string | null,
): SessionIllustrationPatch => (color === null ? { opacity: 0 } : { color, opacity: 1 });

/** A weight that turns an outline that was off back on, at the default weight. */
const visibleWeight = (item: SessionIllustration): number | undefined =>
  item.lineWidth === 0 ? undefined : item.lineWidth;

/** The patch that gives the shape `color` as its outline, or No Outline when
 *  null. An outline that was off comes back at the default weight. */
export const shapeOutlinePatch = (
  item: SessionIllustration,
  color: string | null,
): SessionIllustrationPatch =>
  color === null ? { lineWidth: 0 } : { lineColor: color, lineWidth: visibleWeight(item) };

/** The patch that gives the outline `weight` points; it turns the outline on. */
export const shapeOutlineWeightPatch = (
  _item: SessionIllustration,
  weight: number,
): SessionIllustrationPatch => ({ lineWidth: weight });

/** The patch that gives the outline `dash`; it turns the outline on. */
export const shapeOutlineDashPatch = (
  item: SessionIllustration,
  dash: SessionShapeLineDash,
): SessionIllustrationPatch => ({
  lineDash: dash === 'solid' ? undefined : dash,
  lineWidth: visibleWeight(item),
});

/** Whether the shape's style can be changed: a shape drawn only as its
 *  bounding box has none of its own. */
export const shapeStyleEditable = (item: SessionIllustration): boolean =>
  item.kind === 'shape' && !item.nativeLocked;

/** Whether the outline dash can be set: the file stores no dash style, so a
 *  shape saved with the workbook has none. */
export const shapeDashEditable = (item: SessionIllustration): boolean =>
  shapeStyleEditable(item) && item.nativeObjectId === undefined;

const applyShapePatch = (
  store: SpreadsheetStore,
  item: SessionIllustration,
  patch: SessionIllustrationPatch,
  commandId: string,
  history: History | null,
  wb?: WorkbookHandle,
): SessionIllustration | null =>
  shapeStyleEditable(item)
    ? updateSessionIllustration(store, item.id, patch, history, wb, {
        origin: 'ribbon',
        commandId,
      })
    : null;

/** Shape Fill: one undo step. */
export function applySessionShapeFill(
  store: SpreadsheetStore,
  item: SessionIllustration,
  color: string | null,
  history: History | null = null,
  wb?: WorkbookHandle,
): SessionIllustration | null {
  if (!shapeHasFill(item)) return null;
  return applyShapePatch(store, item, shapeFillPatch(item, color), 'shapeFill', history, wb);
}

/** Shape Outline colour: one undo step. */
export function applySessionShapeOutline(
  store: SpreadsheetStore,
  item: SessionIllustration,
  color: string | null,
  history: History | null = null,
  wb?: WorkbookHandle,
): SessionIllustration | null {
  return applyShapePatch(store, item, shapeOutlinePatch(item, color), 'shapeOutline', history, wb);
}

/** Shape Outline > Weight: one undo step. */
export function applySessionShapeOutlineWeight(
  store: SpreadsheetStore,
  item: SessionIllustration,
  weight: number,
  history: History | null = null,
  wb?: WorkbookHandle,
): SessionIllustration | null {
  if (!(weight > 0)) return null;
  const patch = shapeOutlineWeightPatch(item, weight);
  return applyShapePatch(store, item, patch, 'shapeOutline', history, wb);
}

/** Shape Outline > Dashes: one undo step. */
export function applySessionShapeOutlineDash(
  store: SpreadsheetStore,
  item: SessionIllustration,
  dash: SessionShapeLineDash,
  history: History | null = null,
  wb?: WorkbookHandle,
): SessionIllustration | null {
  if (!shapeDashEditable(item)) return null;
  const patch = shapeOutlineDashPatch(item, dash);
  return applyShapePatch(store, item, patch, 'shapeOutline', history, wb);
}
