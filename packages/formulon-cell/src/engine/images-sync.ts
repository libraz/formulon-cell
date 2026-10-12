import { cellRect, cellRectUnclamped, layoutForView } from '../render/geometry.js';
import type { SessionIllustration, SpreadsheetStore } from '../store/store.js';
import { mutators } from '../store/store.js';
import type { WorkbookHandle } from './workbook-handle.js';
import type { EngineDrawingObject, EngineImage } from './workbook-handle-annotations.js';

/** Extra fields kept on a session illustration for a persisted native picture.
 * They are intentionally optional so shapes and older hosts keep their old
 * session-only representation. The main store may later promote these fields
 * to its public object slice without changing this adapter. */
export interface PersistedImageFields {
  nativeObjectId: number;
  imageBytes: Uint8Array;
  imageFormat: number;
  nativeAnchor: {
    row: number;
    col: number;
    rowOffEmu: number;
    colOffEmu: number;
    widthEmu: number;
    heightEmu: number;
    name: string;
    descr: string;
    /** Older store snapshots omit these fields. Non-one-cell imported
     * drawings are gated before a remove/reinsert edit can change anchor
     * semantics that the public insert API cannot express. */
    anchorKind?: number;
    editAs?: number;
  };
}

export type PersistedSessionImage = SessionIllustration & PersistedImageFields;

const EMU_PER_PX = 9525;

const viewportZoom = (store: SpreadsheetStore | undefined): number => {
  const zoom = store?.getState().viewport.zoom;
  return zoom && Number.isFinite(zoom) && zoom > 0 ? zoom : 1;
};

const placementForPixels = (
  store: SpreadsheetStore | undefined,
  item: PersistedSessionImage,
  patch: Partial<Pick<SessionIllustration, 'x' | 'y' | 'w' | 'h'>>,
): PersistedImageFields['nativeAnchor'] => {
  const original = item.nativeAnchor;
  const zoom = viewportZoom(store);
  const nextWidth = Math.max(
    1,
    Math.round(
      (patch.w === undefined ? original.widthEmu / EMU_PER_PX : patch.w / zoom) * EMU_PER_PX,
    ),
  );
  const nextHeight = Math.max(
    1,
    Math.round(
      (patch.h === undefined ? original.heightEmu / EMU_PER_PX : patch.h / zoom) * EMU_PER_PX,
    ),
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
    let col = original.col;
    let row = original.row;
    let colOffEmu = original.colOffEmu;
    let rowOffEmu = original.rowOffEmu;
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

const base64 = (bytes: Uint8Array): string => {
  let binary = '';
  const chunk = 0x8000;
  for (let i = 0; i < bytes.length; i += chunk) {
    binary += String.fromCharCode(...bytes.subarray(i, i + chunk));
  }
  return globalThis.btoa(binary);
};

const mimeForFormat = (format: number): string => {
  switch (format) {
    case 1:
      return 'image/png';
    case 2:
      return 'image/jpeg';
    case 3:
      return 'image/gif';
    case 4:
      return 'image/bmp';
    default:
      return 'application/octet-stream';
  }
};

export function imageBytesToDataUrl(bytes: Uint8Array, format: number): string {
  return `data:${mimeForFormat(format)};base64,${base64(bytes)}`;
}

const dimensionsForObject = (
  store: SpreadsheetStore | undefined,
  object: EngineDrawingObject,
): { widthEmu: number; heightEmu: number } => {
  let widthEmu = object.cx;
  let heightEmu = object.cy;
  if (object.anchorKind !== 1 || (widthEmu > 0 && heightEmu > 0)) {
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

const anchorFor = (
  store: SpreadsheetStore | undefined,
  object: EngineDrawingObject,
): PersistedImageFields['nativeAnchor'] => {
  const dimensions = dimensionsForObject(store, object);
  return {
    row: object.fromRow,
    col: object.fromCol,
    rowOffEmu: object.fromRowOff,
    colOffEmu: object.fromColOff,
    widthEmu: dimensions.widthEmu,
    heightEmu: dimensions.heightEmu,
    name: object.name,
    descr: object.descr,
    anchorKind: object.anchorKind,
    editAs: object.editAs,
  };
};

const sessionPosition = (
  store: SpreadsheetStore,
  object: EngineDrawingObject,
): { x: number; y: number; w: number; h: number } => {
  const state = store.getState();
  const dimensions = dimensionsForObject(store, object);
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

const imageFromNative = (
  store: SpreadsheetStore,
  sheet: number,
  object: EngineDrawingObject,
  image: EngineImage,
): PersistedSessionImage => {
  const position = sessionPosition(store, object);
  return {
    id: `image-native-${sheet}-${object.objectId}`,
    kind: 'image',
    sheet,
    src: imageBytesToDataUrl(image.bytes, image.format),
    alt: object.descr || object.name,
    ...position,
    nativeObjectId: object.objectId,
    nativeGeometryEditable: false,
    imageBytes: Uint8Array.from(image.bytes),
    imageFormat: image.format,
    nativeAnchor: anchorFor(store, object),
  };
};

/** Hydrate persisted pictures while leaving shapes, charts and other drawing
 * objects untouched. Non-picture objects remain session-only by design. */
export function hydrateImagesFromEngine(
  wb: WorkbookHandle,
  store: SpreadsheetStore,
  sheet: number,
): void {
  if (!wb.capabilities.images) return;
  // A second hydration follows a workbook replacement or an undo. Remove
  // stale native picture projections first while preserving session shapes and
  // any images that have no native object yet.
  store.setState((state) => {
    const illustrations = state.illustrations.illustrations.filter(
      (item) => item.sheet !== sheet || item.nativeObjectId === undefined,
    );
    return illustrations.length === state.illustrations.illustrations.length
      ? state
      : { ...state, illustrations: { illustrations } };
  });
  for (const object of wb.listDrawingObjects(sheet)) {
    if (object.kind !== 0 || object.imageFormat === 0) continue;
    const image = wb.getImage(sheet, object.objectId);
    if (!image) continue;
    mutators.upsertIllustration(
      store,
      imageFromNative(store, sheet, object, image) as unknown as SessionIllustration,
    );
  }
}

export function persistentImageFromState(item: SessionIllustration): PersistedSessionImage | null {
  const candidate = item as PersistedSessionImage;
  return candidate.nativeObjectId && candidate.imageBytes && candidate.nativeAnchor
    ? candidate
    : null;
}

export function nativeImageAnchor(
  wb: WorkbookHandle,
  sheet: number,
  objectId: number,
  store?: SpreadsheetStore,
): PersistedImageFields['nativeAnchor'] | null {
  const object = wb.listDrawingObjects(sheet).find((candidate) => candidate.objectId === objectId);
  return object ? anchorFor(store, object) : null;
}

/** Reinsert one picture with updated dimensions/offsets. Native drawing
 * objects have immutable anchors, so this operation is transactional: the
 * original is restored if the replacement fails. */
export function replacePersistentImage(
  wb: WorkbookHandle,
  item: PersistedSessionImage,
  patch: Partial<Pick<SessionIllustration, 'x' | 'y' | 'w' | 'h'>> & {
    nativeAnchor?: PersistedImageFields['nativeAnchor'];
  },
  store?: SpreadsheetStore,
): number | null {
  if (
    !wb.capabilities.images ||
    !item.nativeObjectId ||
    !item.imageBytes ||
    !item.nativeAnchor ||
    item.nativeGeometryEditable !== true
  ) {
    return null;
  }
  const anchor = item.nativeAnchor;
  // The public native picture API can reinsert one-cell geometry only. A
  // two-cell or absolute imported anchor also carries editAs semantics that
  // would be silently changed by remove/reinsert, so leave the original
  // object untouched and let the UI report an unsupported edit.
  if (anchor.anchorKind !== undefined && anchor.anchorKind !== 0) return null;
  // History replay passes the captured native anchor. It is already in EMU
  // and must not be interpreted as screen pixels under the current zoom or
  // viewport, which may have changed since the original edit.
  const nextAnchor = patch.nativeAnchor
    ? { ...anchor, ...patch.nativeAnchor }
    : placementForPixels(store, item, patch);
  const inserted = wb.insertImage(item.sheet, item.imageBytes, {
    name: anchor.name,
    descr: anchor.descr,
    row: nextAnchor.row,
    col: nextAnchor.col,
    rowOffEmu: nextAnchor.rowOffEmu,
    colOffEmu: nextAnchor.colOffEmu,
    widthEmu: nextAnchor.widthEmu,
    heightEmu: nextAnchor.heightEmu,
  });
  if (inserted === null) return null;
  if (wb.removeImage(item.sheet, item.nativeObjectId)) return inserted;
  // Removing the old object failed, so remove the new copy if possible and
  // leave the original as the saved object. Even if this cleanup fails the
  // original remains present; the caller keeps the old UI identity.
  wb.removeImage(item.sheet, inserted);
  return null;
}

export { EMU_PER_PX };
