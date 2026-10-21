import type { SessionIllustration, SpreadsheetStore } from '../store/store.js';
import { mutators } from '../store/store.js';
import {
  ANCHOR_KIND_ABSOLUTE,
  drawingAnchorOf,
  placementForAnchor,
  placementForPixels,
  sessionPosition,
} from './drawing-geometry.js';
import type { EngineDrawingAnchor } from './types.js';
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
  nativeAnchor: EngineDrawingAnchor;
}

export type PersistedSessionImage = SessionIllustration & PersistedImageFields;

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
    nativeGeometryEditable: object.anchorKind !== ANCHOR_KIND_ABSOLUTE,
    imageBytes: Uint8Array.from(image.bytes),
    imageFormat: image.format,
    nativeAnchor: drawingAnchorOf(object),
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

/** Rebuild every native picture projection after sheets moved or were removed:
 * their stored sheet indexes are stale on every sheet, not just the active one.
 * Other sheets re-project when they are next shown. */
export function reloadImagesFromEngine(wb: WorkbookHandle, store: SpreadsheetStore): void {
  if (!wb.capabilities.images) return;
  store.setState((state) => {
    const illustrations = state.illustrations.illustrations.filter(
      (item) => item.nativeObjectId === undefined,
    );
    return illustrations.length === state.illustrations.illustrations.length
      ? state
      : { ...state, illustrations: { illustrations } };
  });
  hydrateImagesFromEngine(wb, store, store.getState().data.sheetIndex);
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
): EngineDrawingAnchor | null {
  const object = wb.listDrawingObjects(sheet).find((candidate) => candidate.objectId === objectId);
  return object ? drawingAnchorOf(object) : null;
}

/** Move or resize one picture in place. The engine keeps the object id, crop
 * and rotation, and an imported two-cell anchor keeps its kind and `editAs`.
 * Returns false when the object cannot be re-anchored or the engine refuses. */
export function movePersistentImage(
  wb: WorkbookHandle,
  item: PersistedSessionImage,
  patch: Partial<Pick<SessionIllustration, 'x' | 'y' | 'w' | 'h'>> & {
    nativeAnchor?: EngineDrawingAnchor;
  },
  store?: SpreadsheetStore,
): boolean {
  if (
    !wb.capabilities.drawings ||
    !item.nativeObjectId ||
    !item.nativeAnchor ||
    item.nativeGeometryEditable !== true
  ) {
    return false;
  }
  if (item.nativeAnchor.anchorKind === ANCHOR_KIND_ABSOLUTE) return false;
  // History replay passes the captured native anchor. It is already in EMU
  // and must not be interpreted as screen pixels under the current zoom or
  // viewport, which may have changed since the original edit.
  const placement = patch.nativeAnchor
    ? placementForAnchor(store, patch.nativeAnchor)
    : placementForPixels(store, item.nativeAnchor, patch);
  return wb.setDrawingAnchor(item.sheet, item.nativeObjectId, placement);
}
