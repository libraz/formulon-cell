import { AnchorEditAs, AnchorKind, DrawingObjectKind } from '@libraz/formulon';
import {
  cellOrigin,
  DUPLICATE_OFFSET_PX,
  placementForPixels,
  sessionPosition,
} from '../engine/drawing-geometry.js';
import {
  movePersistentImage,
  nativeImageAnchor,
  persistentImageFromState,
} from '../engine/images-sync.js';
import type { EngineDrawingAnchor, Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import {
  mutators,
  type SessionIllustration,
  type SessionShapeKind,
  type SpreadsheetStore,
} from '../store/store.js';
import type { History } from './history.js';
import type { InteractionOrigin } from './interaction-policy.js';
import { type ObjectActionContext, objectWrite } from './object-actions.js';
import { applyIllustrationsSnapshot, captureIllustrationsSnapshot } from './slice-history.js';

export interface CreateSessionShapeOptions {
  id?: string;
  shape: SessionShapeKind;
  x?: number;
  y?: number;
  w?: number;
  h?: number;
  color?: string;
  radius?: number;
  /** Outline weight in points; 0 is No Outline. */
  lineWidth?: number;
  opacity?: number;
  authorization?: IllustrationMutationAuthorization;
}

export interface CreateSessionImageOptions {
  id?: string;
  src: string;
  alt?: string;
  x?: number;
  y?: number;
  w?: number;
  h?: number;
  /** When present, insert a native drawing object and retain its bytes for
   * subsequent move/resize/delete operations. */
  bytes?: Uint8Array;
  wb?: WorkbookHandle;
  /** Anchor the native picture at `x` / `y` (grid pixels) instead of the top-left
   * of the range, so a reloaded file shows it where it was drawn. */
  anchorAtPixels?: boolean;
  /** Session-only fields carried on the illustration, including through redo. */
  extra?: Record<string, unknown>;
  authorization?: IllustrationMutationAuthorization;
}

export interface IllustrationMutationAuthorization {
  origin?: InteractionOrigin;
  commandId?: string;
}

export type SessionIllustrationPatch = Partial<Omit<SessionIllustration, 'id'>>;
export type SessionIllustrationArrangeAction =
  | 'bring-forward'
  | 'send-backward'
  | 'bring-front'
  | 'send-back';

/** A picture whose anchor the engine cannot change. A shape in that state can
 *  still be restyled and deleted. */
const isGeometryLockedPicture = (item: SessionIllustration): boolean =>
  item.kind === 'image' &&
  item.nativeObjectId !== undefined &&
  item.nativeGeometryEditable !== true;

/** Whether the illustration can be deleted. A shape drawn only as its
 *  bounding box cannot be put back faithfully by an undo, so it stays. */
export const canDeleteIllustration = (item: SessionIllustration): boolean =>
  !isGeometryLockedPicture(item) && !item.nativeLocked;

/** Whether the illustration can be copied; the same objects can be deleted. */
export const canDuplicateIllustration = canDeleteIllustration;

const sameIllustrations = (
  a: readonly SessionIllustration[],
  b: readonly SessionIllustration[],
): boolean => JSON.stringify(a) === JSON.stringify(b);

const recordIllustrationChange = (
  history: History | null,
  store: SpreadsheetStore,
  mutate: () => void,
): void => {
  if (!history || history.isReplaying()) {
    mutate();
    return;
  }
  const before = captureIllustrationsSnapshot(store.getState());
  mutate();
  const after = captureIllustrationsSnapshot(store.getState());
  if (sameIllustrations(before, after)) return;
  history.push({
    undo: () => applyIllustrationsSnapshot(store, before),
    redo: () => applyIllustrationsSnapshot(store, after),
  });
};

/** The entry point of an illustration change: `authorization`, or the instance API running
 *  `commandId`. */
const illustrationEntry = (
  store: SpreadsheetStore,
  history: History | null,
  authorization: IllustrationMutationAuthorization | undefined,
  commandId: string,
): ObjectActionContext => ({
  store,
  history,
  origin: authorization?.origin ?? 'instanceApi',
  commandId: authorization?.commandId ?? commandId,
});

const defaultShapeId = (range: Range, shape: SessionShapeKind): string =>
  `shape-${range.sheet}-${range.r0}-${range.c0}-${shape}`;

const defaultImageId = (range: Range): string => `image-${range.sheet}-${range.r0}-${range.c0}`;

export function createSessionShape(
  store: SpreadsheetStore,
  range: Range,
  options: CreateSessionShapeOptions,
  history: History | null = null,
): SessionIllustration | null {
  return objectWrite(
    illustrationEntry(store, history, options.authorization, 'shapeInsert'),
    range.sheet,
    null,
    () => {
      const item: SessionIllustration = {
        id: options.id ?? defaultShapeId(range, options.shape),
        kind: 'shape',
        shape: options.shape,
        sheet: range.sheet,
        x: options.x,
        y: options.y,
        w: options.w,
        h: options.h,
        color: options.color,
        radius: options.radius,
        lineWidth: options.lineWidth,
        opacity: options.opacity,
      };
      recordIllustrationChange(history, store, () => mutators.upsertIllustration(store, item));
      return item;
    },
  );
}

export function createSessionImage(
  store: SpreadsheetStore,
  range: Range,
  options: CreateSessionImageOptions,
  history: History | null = null,
): SessionIllustration | null {
  return objectWrite(
    illustrationEntry(store, history, options.authorization, 'pictureInsert'),
    range.sheet,
    null,
    () => {
      const item: SessionIllustration = {
        ...(options.extra as Partial<SessionIllustration> | undefined),
        id: options.id ?? defaultImageId(range),
        kind: 'image',
        src: options.src,
        alt: options.alt,
        sheet: range.sheet,
        x: options.x,
        y: options.y,
        w: options.w,
        h: options.h,
      };
      const previous = sessionIllustrationById(store.getState(), item.id);
      let nativeObjectId: number | null = null;
      let nativeFormat: number | undefined;
      let nativeWidthEmu: number | undefined;
      let nativeHeightEmu: number | undefined;
      if (options.bytes && options.bytes.length > 0 && options.wb?.capabilities.images) {
        const probe = options.wb.probeImage(options.bytes);
        if (!probe) return null;
        nativeWidthEmu = Math.max(1, Math.round((options.w ?? probe.widthPx) * 9525));
        nativeHeightEmu = Math.max(1, Math.round((options.h ?? probe.heightPx) * 9525));
        const placed = options.anchorAtPixels
          ? placementForPixels(
              store,
              {
                kind: AnchorKind.OneCell,
                editAs: AnchorEditAs.OneCell,
                fromRow: range.r0,
                fromCol: range.c0,
                fromRowOffsetEmu: 0,
                fromColOffsetEmu: 0,
                toRow: 0,
                toCol: 0,
                toRowOffsetEmu: 0,
                toColOffsetEmu: 0,
                widthEmu: nativeWidthEmu,
                heightEmu: nativeHeightEmu,
              },
              { x: options.x, y: options.y, w: options.w, h: options.h },
            )
          : null;
        if (placed) {
          nativeWidthEmu = placed.widthEmu ?? nativeWidthEmu;
          nativeHeightEmu = placed.heightEmu ?? nativeHeightEmu;
        }
        nativeObjectId = options.wb.insertImage(range.sheet, options.bytes, {
          name: options.alt ?? '',
          description: options.alt ?? '',
          anchor: {
            fromRow: placed?.fromRow ?? range.r0,
            fromCol: placed?.fromCol ?? range.c0,
            ...(placed
              ? {
                  fromRowOffsetEmu: placed.fromRowOffsetEmu,
                  fromColOffsetEmu: placed.fromColOffsetEmu,
                }
              : {}),
            widthEmu: nativeWidthEmu,
            heightEmu: nativeHeightEmu,
          },
        });
        if (nativeObjectId === null) return null;
        nativeFormat = probe.format;
      }
      const requestedAnchor: EngineDrawingAnchor = {
        kind: AnchorKind.OneCell,
        editAs: AnchorEditAs.OneCell,
        fromRow: range.r0,
        fromCol: range.c0,
        fromRowOffsetEmu: 0,
        fromColOffsetEmu: 0,
        toRow: 0,
        toCol: 0,
        toRowOffsetEmu: 0,
        toColOffsetEmu: 0,
        widthEmu: nativeWidthEmu ?? 1,
        heightEmu: nativeHeightEmu ?? 1,
      };
      const persistedItem: SessionIllustration =
        nativeObjectId === null
          ? item
          : {
              ...item,
              nativeObjectId,
              nativeGeometryEditable: true,
              imageBytes: Uint8Array.from(options.bytes ?? []),
              imageFormat: nativeFormat,
              nativeAnchor:
                options.wb && nativeObjectId !== null
                  ? (nativeImageAnchor(options.wb, range.sheet, nativeObjectId) ?? requestedAnchor)
                  : requestedAnchor,
            };
      if (nativeObjectId !== null && options.wb) {
        const nativeWb = options.wb;
        const previousPersisted = previous ? persistentImageFromState(previous) : null;
        const createdSnapshot = nativeWb.snapshotDrawing(range.sheet, nativeObjectId);
        const previousSnapshot = previousPersisted
          ? nativeWb.snapshotDrawing(previousPersisted.sheet, previousPersisted.nativeObjectId)
          : null;
        // A session id collision replaces the old projection. Remove its native
        // object only after the new object exists, so a failed insertion leaves
        // the original picture and projection untouched.
        if (
          !createdSnapshot ||
          (previousPersisted &&
            (!previousSnapshot ||
              !nativeWb.removeDrawing(previousPersisted.sheet, previousPersisted.nativeObjectId)))
        ) {
          nativeWb.removeDrawing(range.sheet, nativeObjectId);
          return null;
        }
        const inserted = persistentImageFromState(persistedItem);
        if (!inserted) return null;
        mutators.upsertIllustration(store, persistedItem);
        if (history && !history.isReplaying()) {
          history.push({
            undo: () => {
              const current = sessionIllustrationById(store.getState(), persistedItem.id);
              const currentPersisted = current ? persistentImageFromState(current) : null;
              if (
                !currentPersisted ||
                !current ||
                !nativeWb.removeDrawing(currentPersisted.sheet, currentPersisted.nativeObjectId)
              ) {
                throw new Error('image create undo failed');
              }
              mutators.removeIllustration(store, persistedItem.id);
              if (previous) {
                if (previousPersisted && previousSnapshot) {
                  const restoredId = nativeWb.restoreDrawing(
                    previousPersisted.sheet,
                    previousSnapshot,
                  );
                  if (restoredId === null) throw new Error('image create undo restore failed');
                  mutators.upsertIllustration(store, {
                    ...previous,
                    nativeObjectId: restoredId,
                    nativeAnchor:
                      nativeImageAnchor(nativeWb, previousPersisted.sheet, restoredId) ??
                      previousPersisted.nativeAnchor,
                  });
                } else {
                  mutators.upsertIllustration(store, previous);
                }
              }
            },
            redo: () => {
              const current = sessionIllustrationById(store.getState(), persistedItem.id);
              const currentPersisted = current ? persistentImageFromState(current) : null;
              if (
                currentPersisted &&
                !nativeWb.removeDrawing(currentPersisted.sheet, currentPersisted.nativeObjectId)
              ) {
                throw new Error('image create redo cleanup failed');
              }
              const redoneId = nativeWb.restoreDrawing(range.sheet, createdSnapshot);
              if (redoneId === null) throw new Error('image create redo failed');
              mutators.upsertIllustration(store, {
                ...persistedItem,
                nativeObjectId: redoneId,
                nativeAnchor:
                  nativeImageAnchor(nativeWb, range.sheet, redoneId) ?? inserted.nativeAnchor,
              });
            },
          });
        }
        return persistedItem;
      }
      recordIllustrationChange(history, store, () =>
        mutators.upsertIllustration(store, persistedItem),
      );
      return persistedItem;
    },
  );
}

/** Select a picture or shape the ribbon inserted, as the spreadsheet does;
 *  the overlay showing it names it. */
const selectInserted = (
  store: SpreadsheetStore,
  item: SessionIllustration | null,
): SessionIllustration | null => {
  if (item) mutators.setSelectedObject(store, { id: item.id, name: '' });
  return item;
};

export function createRibbonShapeFromSelection(
  store: SpreadsheetStore,
  range: Range,
  shape: SessionShapeKind,
  history: History | null = null,
  authorization?: IllustrationMutationAuthorization,
): SessionIllustration | null {
  const count = store.getState().illustrations.illustrations.length;
  const item = createSessionShape(
    store,
    range,
    {
      id: `ribbon-shape-${range.sheet}-${range.r0}-${range.c0}-${shape}-${count}`,
      shape,
      x: 300 + (count % 3) * 24,
      y: 88 + (count % 3) * 24,
      w: shape === 'line' || shape === 'arrow' ? 180 : 160,
      h: shape === 'line' || shape === 'arrow' ? 80 : 96,
      authorization,
    },
    history,
  );
  return selectInserted(store, item);
}

export function createRibbonImageFromSelection(
  store: SpreadsheetStore,
  range: Range,
  src: string,
  history: History | null = null,
  alt?: string,
  native?: { bytes?: Uint8Array; wb?: WorkbookHandle },
  authorization?: IllustrationMutationAuthorization,
): SessionIllustration | null {
  const count = store.getState().illustrations.illustrations.length;
  const item = createSessionImage(
    store,
    range,
    {
      id: `ribbon-image-${range.sheet}-${range.r0}-${range.c0}-${count}`,
      src,
      alt: alt ?? src,
      x: 300 + (count % 3) * 24,
      y: 88 + (count % 3) * 24,
      w: 240,
      h: 160,
      bytes: native?.bytes,
      wb: native?.wb,
      authorization,
    },
    history,
  );
  return selectInserted(store, item);
}

export function listSessionIllustrations(state: {
  illustrations: { illustrations: readonly SessionIllustration[] };
}): readonly SessionIllustration[] {
  return state.illustrations.illustrations;
}

export function sessionIllustrationById(
  state: { illustrations: { illustrations: readonly SessionIllustration[] } },
  id: string,
): SessionIllustration | null {
  return state.illustrations.illustrations.find((item) => item.id === id) ?? null;
}

/** The picture or shape selected on the active sheet, or null when none is selected. */
export function selectedSessionIllustration(state: {
  data: { sheetIndex: number };
  ui: { selectedObject: { id: string } | null };
  illustrations: { illustrations: readonly SessionIllustration[] };
}): SessionIllustration | null {
  const id = state.ui.selectedObject?.id;
  const item = id === undefined ? null : sessionIllustrationById(state, id);
  return item && item.sheet === state.data.sheetIndex ? item : null;
}

export function clearSessionIllustration(
  store: SpreadsheetStore,
  id: string,
  history: History | null = null,
  wb?: WorkbookHandle,
  authorization?: IllustrationMutationAuthorization,
): boolean {
  const item = sessionIllustrationById(store.getState(), id);
  if (!item) return false;
  if (!canDeleteIllustration(item)) return false;
  return objectWrite(
    illustrationEntry(store, history, authorization, 'pictureDelete'),
    item.sheet,
    false,
    () => {
      const persisted = persistentImageFromState(item);
      const snapshot =
        persisted && wb ? wb.snapshotDrawing(item.sheet, persisted.nativeObjectId) : null;
      if (
        persisted &&
        (!wb || !snapshot || !wb.removeDrawing(item.sheet, persisted.nativeObjectId))
      ) {
        return false;
      }
      if (persisted && wb && snapshot) {
        mutators.removeIllustration(store, id);
        if (history && !history.isReplaying()) {
          history.push({
            undo: () => {
              const nativeObjectId = wb.restoreDrawing(item.sheet, snapshot);
              if (nativeObjectId === null) throw new Error('image undo restore failed');
              mutators.upsertIllustration(store, { ...item, nativeObjectId });
            },
            redo: () => {
              const current = persistentImageFromState(
                sessionIllustrationById(store.getState(), id) ?? item,
              );
              if (!current || !wb.removeDrawing(item.sheet, current.nativeObjectId)) {
                throw new Error('image redo delete failed');
              }
              mutators.removeIllustration(store, id);
            },
          });
        }
        return true;
      }
      recordIllustrationChange(history, store, () => mutators.removeIllustration(store, id));
      return true;
    },
  );
}

export function updateSessionIllustration(
  store: SpreadsheetStore,
  id: string,
  patch: SessionIllustrationPatch,
  history: History | null = null,
  wb?: WorkbookHandle,
  authorization?: IllustrationMutationAuthorization,
): SessionIllustration | null {
  const item = sessionIllustrationById(store.getState(), id);
  if (!item) return null;
  if (isGeometryLockedPicture(item)) return null;
  return objectWrite(
    illustrationEntry(store, history, authorization, 'pictureUpdate'),
    item.sheet,
    null,
    () => {
      const persisted = persistentImageFromState(item);
      if (persisted && wb) {
        const before = { ...item } as SessionIllustration;
        if (!movePersistentImage(wb, persisted, patch, store)) return item;
        const after = {
          ...item,
          ...patch,
          nativeAnchor:
            nativeImageAnchor(wb, item.sheet, persisted.nativeObjectId) ?? persisted.nativeAnchor,
        } as SessionIllustration;
        mutators.updateIllustration(store, id, after as SessionIllustrationPatch);
        if (history && !history.isReplaying()) {
          history.push({
            undo: () => {
              const current = sessionIllustrationById(store.getState(), id);
              if (!current) throw new Error('image undo state missing');
              const currentPersisted = persistentImageFromState(current);
              if (!currentPersisted) throw new Error('image undo native state missing');
              if (!movePersistentImage(wb, currentPersisted, before, store)) {
                throw new Error('image undo restore failed');
              }
              mutators.updateIllustration(store, id, {
                ...before,
                nativeAnchor:
                  nativeImageAnchor(wb, item.sheet, currentPersisted.nativeObjectId) ??
                  currentPersisted.nativeAnchor,
              });
            },
            redo: () => {
              const current = sessionIllustrationById(store.getState(), id);
              if (!current) throw new Error('image redo state missing');
              const currentPersisted = persistentImageFromState(current);
              if (!currentPersisted) throw new Error('image redo native state missing');
              if (!movePersistentImage(wb, currentPersisted, after, store)) {
                throw new Error('image redo restore failed');
              }
              mutators.updateIllustration(store, id, {
                ...after,
                nativeAnchor:
                  nativeImageAnchor(wb, item.sheet, currentPersisted.nativeObjectId) ??
                  currentPersisted.nativeAnchor,
              });
            },
          });
        }
        return sessionIllustrationById(store.getState(), id);
      } else if (persisted && !wb) {
        return item;
      }
      recordIllustrationChange(history, store, () =>
        mutators.updateIllustration(store, id, patch as SessionIllustrationPatch),
      );
      return sessionIllustrationById(store.getState(), id);
    },
  );
}

const duplicateId = (
  state: { illustrations: { illustrations: readonly SessionIllustration[] } },
  id: string,
): string => {
  const ids = new Set(state.illustrations.illustrations.map((item) => item.id));
  let index = 1;
  let next = `${id}-copy`;
  while (ids.has(next)) {
    index += 1;
    next = `${id}-copy-${index}`;
  }
  return next;
};

export function duplicateSessionIllustration(
  store: SpreadsheetStore,
  id: string,
  history: History | null = null,
  wb?: WorkbookHandle,
  authorization?: IllustrationMutationAuthorization,
): SessionIllustration | null {
  const state = store.getState();
  const item = sessionIllustrationById(state, id);
  if (!item) return null;
  return objectWrite(
    illustrationEntry(store, history, authorization, 'pictureDuplicate'),
    item.sheet,
    null,
    () => {
      const clip = copySessionIllustration(store, id, wb);
      if (!clip) return null;
      const from = item.nativeAnchor
        ? sessionPosition(store, item.nativeAnchor)
        : { x: item.x ?? 300, y: item.y ?? 88 };
      return placeIllustrationCopy(
        store,
        clip,
        {
          id: duplicateId(state, id),
          sheet: item.sheet,
          row: item.nativeAnchor?.fromRow ?? 0,
          col: item.nativeAnchor?.fromCol ?? 0,
        },
        { x: from.x + DUPLICATE_OFFSET_PX, y: from.y + DUPLICATE_OFFSET_PX },
        history,
        wb,
      );
    },
  );
}

/** A picture or shape as copied to the clipboard; it outlives the object it was taken from. */
export interface IllustrationClip {
  readonly item: SessionIllustration;
  /** The drawing of a picture saved with the workbook. */
  readonly drawing: Uint8Array | null;
}

/** Copy a picture or shape for a later {@link pasteSessionIllustration}; null
 *  when it cannot be copied. */
export function copySessionIllustration(
  store: SpreadsheetStore,
  id: string,
  wb?: WorkbookHandle,
): IllustrationClip | null {
  const item = sessionIllustrationById(store.getState(), id);
  if (!item || !canDuplicateIllustration(item)) return null;
  const persisted = persistentImageFromState(item);
  if (!persisted) return { item: { ...item }, drawing: null };
  const drawing = wb?.snapshotDrawing(item.sheet, persisted.nativeObjectId) ?? null;
  return drawing ? { item: { ...item }, drawing } : null;
}

/**
 * Paste a copied picture or shape onto `at.sheet` with its top-left corner at
 * cell (row, col), as the spreadsheet does, moved by `offset` pixels. Undo
 * removes the pasted object.
 */
export function pasteSessionIllustration(
  store: SpreadsheetStore,
  clip: IllustrationClip,
  at: { sheet: number; row: number; col: number },
  history: History | null = null,
  wb?: WorkbookHandle,
  authorization?: IllustrationMutationAuthorization,
  offset: { x: number; y: number } = { x: 0, y: 0 },
): SessionIllustration | null {
  return objectWrite(
    illustrationEntry(store, history, authorization, 'picturePaste'),
    at.sheet,
    null,
    () => {
      return placeIllustrationCopy(
        store,
        clip,
        { ...at, id: duplicateId(store.getState(), clip.item.id) },
        (({ x, y }) => ({ x: x + offset.x, y: y + offset.y }))(cellOrigin(store, at.row, at.col)),
        history,
        wb,
      );
    },
  );
}

const sizeOf = (store: SpreadsheetStore, anchor: EngineDrawingAnchor): { w: number; h: number } => {
  const { w, h } = sessionPosition(store, anchor);
  return { w, h };
};

/** Add `source` as a new object `target.id` on `target.sheet` with its top-left
 *  corner at `to`, recorded as one undoable step. */
const placeIllustrationCopy = (
  store: SpreadsheetStore,
  source: IllustrationClip,
  target: { id: string; sheet: number; row: number; col: number },
  to: { x: number; y: number },
  history: History | null,
  wb: WorkbookHandle | undefined,
): SessionIllustration | null => {
  const { item } = source;
  const persisted = persistentImageFromState(item);
  const copy: SessionIllustration = {
    ...item,
    id: target.id,
    sheet: target.sheet,
    x: to.x,
    y: to.y,
    // A copied shape is a new engine shape, anchored where it is placed.
    ...(item.kind === 'shape' && item.nativeObjectId !== undefined
      ? {
          ...(item.nativeAnchor ? sizeOf(store, item.nativeAnchor) : {}),
          nativeObjectId: undefined,
          nativeAnchor: undefined,
          nativeGeometryEditable: undefined,
        }
      : {}),
  };
  if (!persisted) {
    recordIllustrationChange(history, store, () => mutators.upsertIllustration(store, copy));
    return copy;
  }
  if (!wb || !source.drawing) return null;
  const newId = wb.restoreDrawing(target.sheet, source.drawing, { newId: true });
  if (newId === null) return null;
  const placement = placementForPixels(store, persisted.nativeAnchor, to);
  const redoSnapshot = wb.setDrawingAnchor(target.sheet, newId, placement)
    ? wb.snapshotDrawing(target.sheet, newId)
    : null;
  if (!redoSnapshot) {
    wb.removeDrawing(target.sheet, newId);
    return null;
  }
  Object.assign(copy as object, {
    nativeObjectId: newId,
    imageBytes: Uint8Array.from(persisted.imageBytes),
    imageFormat: persisted.imageFormat,
    nativeAnchor: nativeImageAnchor(wb, target.sheet, newId) ?? persisted.nativeAnchor,
  });
  mutators.upsertIllustration(store, copy);
  if (history && !history.isReplaying()) {
    history.push({
      undo: () => {
        const current = persistentImageFromState(
          sessionIllustrationById(store.getState(), copy.id) ?? copy,
        );
        if (!current || !wb.removeDrawing(copy.sheet, current.nativeObjectId)) {
          throw new Error('image copy undo failed');
        }
        mutators.removeIllustration(store, copy.id);
      },
      redo: () => {
        const nativeObjectId = wb.restoreDrawing(copy.sheet, redoSnapshot);
        if (nativeObjectId === null) throw new Error('image copy redo failed');
        Object.assign(copy as object, { nativeObjectId });
        mutators.upsertIllustration(store, copy);
      },
    });
  }
  return copy;
};

const reorderIllustrations = (
  illustrations: readonly SessionIllustration[],
  id: string,
  action: SessionIllustrationArrangeAction,
): readonly SessionIllustration[] | null => {
  const item = illustrations.find((candidate) => candidate.id === id);
  if (!item) return null;
  const sameSheet = illustrations.filter((candidate) => candidate.sheet === item.sheet);
  const currentIndex = sameSheet.findIndex((candidate) => candidate.id === id);
  if (currentIndex < 0) return null;
  const targetIndex =
    action === 'bring-front'
      ? sameSheet.length - 1
      : action === 'send-back'
        ? 0
        : action === 'bring-forward'
          ? Math.min(sameSheet.length - 1, currentIndex + 1)
          : Math.max(0, currentIndex - 1);
  if (targetIndex === currentIndex) return null;

  const reorderedSameSheet = sameSheet.filter((candidate) => candidate.id !== id);
  reorderedSameSheet.splice(targetIndex, 0, item);
  let sheetCursor = 0;
  return illustrations.map((candidate) => {
    if (candidate.sheet !== item.sheet) return candidate;
    const next = reorderedSameSheet[sheetCursor];
    sheetCursor += 1;
    return next ?? candidate;
  });
};

/** Engine index to move a native picture to so its stacking among pictures
 * matches `arranged`, or null when it already sits there. Charts and other
 * objects keep their slots: the picture takes the slot of the picture that
 * currently holds its new rank. */
const nativeZTarget = (
  wb: WorkbookHandle,
  item: SessionIllustration,
  arranged: readonly SessionIllustration[],
): { from: number; to: number } | null => {
  const objects = wb.listDrawingObjects(item.sheet);
  const from = objects.findIndex((object) => object.objectId === item.nativeObjectId);
  if (from < 0) return null;
  const rank = arranged
    .filter(
      (candidate) =>
        candidate.sheet === item.sheet &&
        candidate.kind === 'image' &&
        candidate.nativeObjectId !== undefined,
    )
    .findIndex((candidate) => candidate.id === item.id);
  const slots = objects.flatMap((object, index) =>
    object.kind === DrawingObjectKind.Picture ? [index] : [],
  );
  const to = slots[rank];
  return rank < 0 || to === undefined || to === from ? null : { from, to };
};

/** Where `activeId` sits in a sheet's drawing stack, listed back to front:
 *  whether it is in the stack, and whether nothing lies in front of or
 *  behind it to bring it forward or send it backward past. */
export function arrangeStackPosition(
  stack: readonly { id: string }[],
  activeId: string | null,
): { hasTarget: boolean; atFront: boolean; atBack: boolean } {
  const index = activeId === null ? -1 : stack.findIndex((item) => item.id === activeId);
  const hasTarget = index >= 0;
  return {
    hasTarget,
    atFront: !hasTarget || index === stack.length - 1,
    atBack: !hasTarget || index === 0,
  };
}

export function arrangeSessionIllustration(
  store: SpreadsheetStore,
  id: string,
  action: SessionIllustrationArrangeAction,
  history: History | null = null,
  authorization?: IllustrationMutationAuthorization,
  wb?: WorkbookHandle,
): SessionIllustration | null {
  const state = store.getState();
  const item = sessionIllustrationById(state, id);
  if (!item) return null;
  return objectWrite(
    illustrationEntry(store, history, authorization, 'pictureArrange'),
    item.sheet,
    null,
    () => {
      const arranged = reorderIllustrations(state.illustrations.illustrations, id, action);
      if (!arranged) return null;
      if (item.nativeObjectId !== undefined) {
        // The engine orders pictures and charts only; a saved shape keeps its slot.
        if (item.kind === 'shape' || !wb?.capabilities.drawings) return null;
        const z = nativeZTarget(wb, item, arranged);
        if (!z || !wb.setDrawingZOrder(item.sheet, item.nativeObjectId, z.to)) return null;
        const before = state.illustrations.illustrations;
        const apply = (list: readonly SessionIllustration[]): void =>
          store.setState((current) => ({ ...current, illustrations: { illustrations: list } }));
        apply(arranged);
        if (history && !history.isReplaying()) {
          const nativeObjectId = item.nativeObjectId;
          const replay = (origin: 'undo' | 'redo', index: number, list: typeof before): void => {
            const current = sessionIllustrationById(store.getState(), id);
            if (!current || !wb.setDrawingZOrder(item.sheet, nativeObjectId, index)) {
              throw new Error(`image arrange ${origin} failed`);
            }
            apply(list);
          };
          history.push({
            undo: () => replay('undo', z.from, before),
            redo: () => replay('redo', z.to, arranged),
          });
        }
        return sessionIllustrationById(store.getState(), id);
      }
      recordIllustrationChange(history, store, () => {
        store.setState((state) => {
          return { ...state, illustrations: { illustrations: arranged } };
        });
      });
      return sessionIllustrationById(store.getState(), id);
    },
  );
}
