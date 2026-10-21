import type { ImagePlacement } from '@libraz/formulon';
import {
  bodyBandOrigin,
  cellRectUnclamped,
  gridOriginX,
  gridOriginY,
  layoutForView,
} from '../render/geometry.js';
import { renderScale } from '../store/render-scale.js';
import type { SessionIllustration, SpreadsheetStore } from '../store/store.js';
import type { EngineDrawingAnchor } from './types.js';
import type { EngineDrawingObject } from './workbook-handle-annotations.js';

/** `AnchorKind.Absolute`: the engine cannot re-anchor such an object. */
export const ANCHOR_KIND_ABSOLUTE = 2;
/** `AnchorKind.TwoCell`: the size follows the `to` marker. */
const ANCHOR_KIND_TWO_CELL = 1;

export const EMU_PER_PX = 9525;

/** How far a duplicated chart, picture or shape sits from its original. */
export const DUPLICATE_OFFSET_PX = 24;

/** Canonical → screen px factor of the current view; 1 without a store. */
export const viewportZoom = (store: SpreadsheetStore | undefined): number => {
  if (!store) return 1;
  const { layout, viewport } = store.getState();
  return renderScale(layout, viewport);
};

/** The anchor fields of a drawing object, as the store keeps them. */
export const drawingAnchorOf = (object: EngineDrawingObject): EngineDrawingAnchor => ({
  anchorKind: object.anchorKind,
  editAs: object.editAs,
  fromRow: object.fromRow,
  fromCol: object.fromCol,
  fromRowOff: object.fromRowOff,
  fromColOff: object.fromColOff,
  toRow: object.toRow,
  toCol: object.toCol,
  toRowOff: object.toRowOff,
  toColOff: object.toColOff,
  cx: object.cx,
  cy: object.cy,
});

const dimensionsForAnchor = (
  store: SpreadsheetStore | undefined,
  object: EngineDrawingAnchor,
): { widthEmu: number; heightEmu: number } => {
  let widthEmu = object.cx;
  let heightEmu = object.cy;
  if (object.anchorKind !== ANCHOR_KIND_TWO_CELL || (widthEmu > 0 && heightEmu > 0)) {
    return { widthEmu: Math.max(1, widthEmu), heightEmu: Math.max(1, heightEmu) };
  }
  try {
    if (!store) throw new Error('drawing dimensions need a store layout');
    const state = store.getState();
    const layout = layoutForView(state);
    const zoom = viewportZoom(store);
    const from = cellRectUnclamped(layout, state.viewport, object.fromRow, object.fromCol);
    const to = cellRectUnclamped(layout, state.viewport, object.toRow, object.toCol);
    if (widthEmu <= 0) {
      widthEmu = Math.abs(
        ((to.x - from.x) / zoom) * EMU_PER_PX + object.toColOff - object.fromColOff,
      );
    }
    if (heightEmu <= 0) {
      heightEmu = Math.abs(
        ((to.y - from.y) / zoom) * EMU_PER_PX + object.toRowOff - object.fromRowOff,
      );
    }
  } catch {
    // Keep the native dimensions below when the current view has no usable
    // layout (for example while a sheet is being replaced).
  }
  return { widthEmu: Math.max(1, widthEmu), heightEmu: Math.max(1, heightEmu) };
};

/** `anchor` as the placement `setDrawingAnchor` takes. A two-cell anchor that
 *  reports no size is measured on the current layout. */
export const placementForAnchor = (
  store: SpreadsheetStore | undefined,
  anchor: EngineDrawingAnchor,
): ImagePlacement => {
  const dimensions = dimensionsForAnchor(store, anchor);
  return {
    anchorKind: anchor.anchorKind,
    editAs: anchor.editAs,
    row: anchor.fromRow,
    col: anchor.fromCol,
    rowOffEmu: anchor.fromRowOff,
    colOffEmu: anchor.fromColOff,
    widthEmu: dimensions.widthEmu,
    heightEmu: dimensions.heightEmu,
  };
};

type ViewLayoutOf = ReturnType<typeof layoutForView>;
type ViewportOf = ReturnType<SpreadsheetStore['getState']>['viewport'];

/** The column whose grid pixels hold `x`, clamped to the sheet. */
const colAtX = (layout: ViewLayoutOf, viewport: ViewportOf, x: number, zoom: number): number => {
  const origin = cellRectUnclamped(layout, viewport, 0, 0);
  let col = Math.max(0, Math.floor((x - origin.x) / Math.max(1, layout.defaultColWidth * zoom)));
  col = Math.min(16_383, col);
  for (let i = 0; i < 16_384; i += 1) {
    const rect = cellRectUnclamped(layout, viewport, 0, col);
    if (x < rect.x) col = Math.max(0, col - 1);
    else if (x >= rect.x + rect.w) col = Math.min(16_383, col + 1);
    else break;
  }
  return col;
};

/** The row whose grid pixels hold `y`, clamped to the sheet. */
const rowAtY = (layout: ViewLayoutOf, viewport: ViewportOf, y: number, zoom: number): number => {
  const origin = cellRectUnclamped(layout, viewport, 0, 0);
  let row = Math.max(0, Math.floor((y - origin.y) / Math.max(1, layout.defaultRowHeight * zoom)));
  row = Math.min(1_048_575, row);
  for (let i = 0; i < 1_048_576; i += 1) {
    const rect = cellRectUnclamped(layout, viewport, row, 0);
    if (y < rect.y) row = Math.max(0, row - 1);
    else if (y >= rect.y + rect.h) row = Math.min(1_048_575, row + 1);
    else break;
  }
  return row;
};

/** The gridline nearest `value` on `axis`, for objects dragged with Alt. */
export const nearestGridline = (
  store: SpreadsheetStore,
  axis: 'x' | 'y',
  value: number,
): number => {
  const state = store.getState();
  const layout = layoutForView(state);
  const zoom = viewportZoom(store);
  const rect =
    axis === 'x'
      ? cellRectUnclamped(layout, state.viewport, 0, colAtX(layout, state.viewport, value, zoom))
      : cellRectUnclamped(layout, state.viewport, rowAtY(layout, state.viewport, value, zoom), 0);
  const start = axis === 'x' ? rect.x : rect.y;
  const end = start + (axis === 'x' ? rect.w : rect.h);
  return value - start <= end - value ? start : end;
};

/** `anchor` moved and resized to grid pixels; omitted values keep the anchor's. */
export const placementForPixels = (
  store: SpreadsheetStore | undefined,
  anchor: EngineDrawingAnchor,
  patch: Partial<Pick<SessionIllustration, 'x' | 'y' | 'w' | 'h'>>,
): ImagePlacement => {
  const original = placementForAnchor(store, anchor);
  const originalWidth = original.widthEmu ?? 1;
  const originalHeight = original.heightEmu ?? 1;
  const zoom = viewportZoom(store);
  const nextWidth = Math.max(
    1,
    Math.round((patch.w === undefined ? originalWidth / EMU_PER_PX : patch.w / zoom) * EMU_PER_PX),
  );
  const nextHeight = Math.max(
    1,
    Math.round((patch.h === undefined ? originalHeight / EMU_PER_PX : patch.h / zoom) * EMU_PER_PX),
  );
  if (!store) {
    return {
      ...original,
      widthEmu: nextWidth,
      heightEmu: nextHeight,
    };
  }
  const state = store.getState();
  try {
    const layout = layoutForView(state);
    let col = anchor.fromCol;
    let row = anchor.fromRow;
    let colOffEmu = anchor.fromColOff;
    let rowOffEmu = anchor.fromRowOff;
    if (patch.x !== undefined) {
      const targetX = patch.x;
      col = colAtX(layout, state.viewport, targetX, zoom);
      const rect = cellRectUnclamped(layout, state.viewport, 0, col);
      colOffEmu = Math.max(0, Math.round(((targetX - rect.x) / zoom) * EMU_PER_PX));
    }
    if (patch.y !== undefined) {
      const targetY = patch.y;
      row = rowAtY(layout, state.viewport, targetY, zoom);
      const rect = cellRectUnclamped(layout, state.viewport, row, 0);
      rowOffEmu = Math.max(0, Math.round(((targetY - rect.y) / zoom) * EMU_PER_PX));
    }
    return {
      ...original,
      row,
      col,
      rowOffEmu,
      colOffEmu,
      widthEmu: nextWidth,
      heightEmu: nextHeight,
    };
  } catch {
    return {
      ...original,
      widthEmu: nextWidth,
      heightEmu: nextHeight,
    };
  }
};

/** Grid pixels of the object at `anchor` under the current view. An object
 *  whose anchor has scrolled off keeps its true, possibly negative, offset. */
export const sessionPosition = (
  store: SpreadsheetStore,
  object: EngineDrawingAnchor,
): { x: number; y: number; w: number; h: number } => {
  const state = store.getState();
  const dimensions = dimensionsForAnchor(store, object);
  try {
    const rect = cellRectUnclamped(
      layoutForView(state),
      state.viewport,
      object.fromRow,
      object.fromCol,
    );
    const zoom = viewportZoom(store);
    return {
      x: rect.x + (object.fromColOff / EMU_PER_PX) * zoom,
      y: rect.y + (object.fromRowOff / EMU_PER_PX) * zoom,
      w: Math.max(1, dimensions.widthEmu / EMU_PER_PX) * zoom,
      h: Math.max(1, dimensions.heightEmu / EMU_PER_PX) * zoom,
    };
  } catch {
    const zoom = viewportZoom(store);
    const { defaultColWidth, defaultRowHeight } = state.layout;
    return {
      x: (object.fromCol * defaultColWidth + object.fromColOff / EMU_PER_PX) * zoom,
      y: (object.fromRow * defaultRowHeight + object.fromRowOff / EMU_PER_PX) * zoom,
      w: Math.max(1, dimensions.widthEmu / EMU_PER_PX) * zoom,
      h: Math.max(1, dimensions.heightEmu / EMU_PER_PX) * zoom,
    };
  }
};

/** Grid pixels of the top-left corner of cell (row, col) under the current view. */
export const cellOrigin = (
  store: SpreadsheetStore,
  row: number,
  col: number,
): { x: number; y: number } => {
  const state = store.getState();
  try {
    const rect = cellRectUnclamped(layoutForView(state), state.viewport, row, col);
    return { x: rect.x, y: rect.y };
  } catch {
    const zoom = viewportZoom(store);
    return {
      x: col * state.layout.defaultColWidth * zoom,
      y: row * state.layout.defaultRowHeight * zoom,
    };
  }
};

/** How far a drawing's own decorations, such as sizing handles, may reach past it. */
const DECORATION_OUTSET = 64;

/** The CSS `clip-path` that hides the part of an object at `position` lying
 *  under the headers, or under the frozen panes when its anchor scrolls; empty
 *  when nothing is hidden. */
export const drawingClipPath = (
  store: SpreadsheetStore,
  object: EngineDrawingAnchor,
  position: { x: number; y: number; w: number; h: number },
): string => {
  try {
    const state = store.getState();
    const layout = layoutForView(state);
    const band = bodyBandOrigin(layout, state.viewport);
    const top = object.fromRow >= layout.freezeRows ? band.y : gridOriginY(layout);
    const hiddenTop = top - position.y;
    let hiddenLead: number;
    if (layout.rtl) {
      const edge =
        object.fromCol >= layout.freezeCols ? band.x : layout.viewWidth - gridOriginX(layout);
      hiddenLead = position.x + position.w - edge;
    } else {
      const edge = object.fromCol >= layout.freezeCols ? band.x : gridOriginX(layout);
      hiddenLead = edge - position.x;
    }
    if (hiddenTop <= 0 && hiddenLead <= 0) return '';
    const inset = (hidden: number): string =>
      hidden > 0 ? `${hidden}px` : `-${DECORATION_OUTSET}px`;
    const outset = `-${DECORATION_OUTSET}px`;
    return layout.rtl
      ? `inset(${inset(hiddenTop)} ${inset(hiddenLead)} ${outset} ${outset})`
      : `inset(${inset(hiddenTop)} ${outset} ${outset} ${inset(hiddenLead)})`;
  } catch {
    return '';
  }
};
