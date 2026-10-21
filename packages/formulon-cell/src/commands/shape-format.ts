import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { SessionIllustration, SpreadsheetStore } from '../store/store.js';
import type { History } from './history.js';
import {
  type SessionIllustrationPatch,
  updateSessionIllustration,
} from './session-illustration.js';

/** The colour of a shape that has not been given one. */
export const DEFAULT_SHAPE_COLOR = '#0f6cbd';

/** Lines and arrows are all outline; they have no fill. */
export const shapeHasFill = (item: SessionIllustration): boolean =>
  item.kind === 'shape' && item.shape !== 'line' && item.shape !== 'arrow';

/** The fill colour of a shape, or null for No Fill. */
export const shapeFillColor = (item: SessionIllustration): string | null =>
  shapeHasFill(item) && item.opacity !== 0 ? (item.color ?? DEFAULT_SHAPE_COLOR) : null;

/** The outline colour of a shape, or null for No Outline. The outline takes
 *  the shape's colour until it is given one of its own. */
export const shapeOutlineColor = (item: SessionIllustration): string | null =>
  item.lineWidth === 0 ? null : (item.lineColor ?? item.color ?? DEFAULT_SHAPE_COLOR);

/** The patch that gives the shape `color` as its fill, or No Fill when null.
 *  The outline keeps its colour. */
export const shapeFillPatch = (
  item: SessionIllustration,
  color: string | null,
): SessionIllustrationPatch => {
  const outline = { lineColor: item.lineColor ?? item.color ?? DEFAULT_SHAPE_COLOR };
  return color === null ? { ...outline, opacity: 0 } : { ...outline, color, opacity: 1 };
};

/** The patch that gives the shape `color` as its outline, or No Outline when
 *  null. An outline that was off comes back at the default weight. */
export const shapeOutlinePatch = (
  item: SessionIllustration,
  color: string | null,
): SessionIllustrationPatch =>
  color === null
    ? { lineWidth: 0 }
    : { lineColor: color, lineWidth: item.lineWidth === 0 ? undefined : item.lineWidth };

const applyShapePatch = (
  store: SpreadsheetStore,
  item: SessionIllustration,
  patch: SessionIllustrationPatch,
  commandId: string,
  history: History | null,
  wb?: WorkbookHandle,
): SessionIllustration | null =>
  updateSessionIllustration(store, item.id, patch, history, wb, { origin: 'ribbon', commandId });

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

/** Shape Outline: one undo step. */
export function applySessionShapeOutline(
  store: SpreadsheetStore,
  item: SessionIllustration,
  color: string | null,
  history: History | null = null,
  wb?: WorkbookHandle,
): SessionIllustration | null {
  if (item.kind !== 'shape') return null;
  return applyShapePatch(store, item, shapeOutlinePatch(item, color), 'shapeOutline', history, wb);
}
