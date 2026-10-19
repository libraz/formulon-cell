import type { ImagePlacement } from '@libraz/formulon';
import { cellRect, cellRectUnclamped, layoutForView } from '../render/geometry.js';
import type { SessionIllustration, SpreadsheetStore } from '../store/store.js';
import type { EngineDrawingAnchor } from './types.js';
import type { EngineDrawingObject } from './workbook-handle-annotations.js';

/** `AnchorKind.Absolute`: the engine cannot re-anchor such an object. */
export const ANCHOR_KIND_ABSOLUTE = 2;
/** `AnchorKind.TwoCell`: the size follows the `to` marker. */
const ANCHOR_KIND_TWO_CELL = 1;

export const EMU_PER_PX = 9525;

export const viewportZoom = (store: SpreadsheetStore | undefined): number => {
  const zoom = store?.getState().viewport.zoom;
  return zoom && Number.isFinite(zoom) && zoom > 0 ? zoom : 1;
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
      const origin = cellRectUnclamped(layout, state.viewport, 0, 0);
      col = Math.max(
        0,
        Math.floor((targetX - origin.x) / Math.max(1, layout.defaultColWidth * zoom)),
      );
      col = Math.min(16_383, col);
      for (let i = 0; i < 16_384; i += 1) {
        const rect = cellRectUnclamped(layout, state.viewport, 0, col);
        if (targetX < rect.x) col = Math.max(0, col - 1);
        else if (targetX >= rect.x + rect.w) col = Math.min(16_383, col + 1);
        else break;
      }
      const rect = cellRectUnclamped(layout, state.viewport, 0, col);
      colOffEmu = Math.max(0, Math.round(((targetX - rect.x) / zoom) * EMU_PER_PX));
    }
    if (patch.y !== undefined) {
      const targetY = patch.y;
      const origin = cellRectUnclamped(layout, state.viewport, 0, 0);
      row = Math.max(
        0,
        Math.floor((targetY - origin.y) / Math.max(1, layout.defaultRowHeight * zoom)),
      );
      row = Math.min(1_048_575, row);
      for (let i = 0; i < 1_048_576; i += 1) {
        const rect = cellRectUnclamped(layout, state.viewport, row, 0);
        if (targetY < rect.y) row = Math.max(0, row - 1);
        else if (targetY >= rect.y + rect.h) row = Math.min(1_048_575, row + 1);
        else break;
      }
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

/** Grid pixels of the object at `anchor` under the current view. */
export const sessionPosition = (
  store: SpreadsheetStore,
  object: EngineDrawingAnchor,
): { x: number; y: number; w: number; h: number } => {
  const state = store.getState();
  const dimensions = dimensionsForAnchor(store, object);
  try {
    const rect = cellRect(layoutForView(state), state.viewport, object.fromRow, object.fromCol);
    const zoom = viewportZoom(store);
    return {
      x: rect.x + (object.fromColOff / EMU_PER_PX) * zoom,
      y: rect.y + (object.fromRowOff / EMU_PER_PX) * zoom,
      w: Math.max(1, dimensions.widthEmu / EMU_PER_PX) * zoom,
      h: Math.max(1, dimensions.heightEmu / EMU_PER_PX) * zoom,
    };
  } catch {
    const zoom = viewportZoom(store);
    return {
      x: (object.fromCol * 64 + object.fromColOff / EMU_PER_PX) * zoom,
      y: (object.fromRow * 20 + object.fromRowOff / EMU_PER_PX) * zoom,
      w: Math.max(1, dimensions.widthEmu / EMU_PER_PX) * zoom,
      h: Math.max(1, dimensions.heightEmu / EMU_PER_PX) * zoom,
    };
  }
};
