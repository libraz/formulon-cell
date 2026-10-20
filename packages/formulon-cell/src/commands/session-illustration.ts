import { AnchorEditAs, AnchorKind } from '@libraz/formulon';
import {
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
import { interactionControllerFor } from './interaction-controller.js';
import type { InteractionOrigin, OperationEffect, OperationIntent } from './interaction-policy.js';
import { isCellWritable, isSheetProtected } from './protection.js';
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

const imageIntent = (
  commandId: string,
  sheet: number,
  row: number,
  col: number,
  origin: OperationIntent['origin'] = 'instanceApi',
): OperationIntent => {
  const effects: readonly OperationEffect[] = [{ kind: 'cells', cells: [{ sheet, row, col }] }];
  return { operation: 'object', origin, commandId, effects };
};

const sameIllustrations = (
  a: readonly SessionIllustration[],
  b: readonly SessionIllustration[],
): boolean => JSON.stringify(a) === JSON.stringify(b);

const recordIllustrationChange = (
  history: History | null,
  store: SpreadsheetStore,
  mutate: () => void,
  intent: OperationIntent,
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
    intent: { ...intent, origin: 'redo' },
    inverseIntent: { ...intent, origin: 'undo' },
  });
};

const addressForItem = (item: SessionIllustration): { sheet: number; row: number; col: number } => {
  const candidate = item.nativeAnchor;
  return { sheet: item.sheet, row: candidate?.fromRow ?? 0, col: candidate?.fromCol ?? 0 };
};

const canMutateIllustration = (
  store: SpreadsheetStore,
  item: SessionIllustration,
  authorization: IllustrationMutationAuthorization = {},
  commandId = 'pictureUpdate',
): boolean => {
  const address = addressForItem(item);
  if (!isCellWritable(store.getState(), address)) return false;
  const controller = interactionControllerFor(store);
  const resolvedCommandId = authorization.commandId ?? commandId;
  return (
    !controller ||
    controller.canExecute(
      imageIntent(
        resolvedCommandId,
        address.sheet,
        address.row,
        address.col,
        authorization.origin ?? 'instanceApi',
      ),
    ).allowed
  );
};

const canMutateRange = (
  store: SpreadsheetStore,
  range: Range,
  authorization: IllustrationMutationAuthorization = {},
  commandId = 'pictureInsert',
): boolean => {
  const address = { sheet: range.sheet, row: range.r0, col: range.c0 };
  if (!isCellWritable(store.getState(), address)) return false;
  const controller = interactionControllerFor(store);
  const resolvedCommandId = authorization.commandId ?? commandId;
  return (
    !controller ||
    controller.canExecute(
      imageIntent(
        resolvedCommandId,
        address.sheet,
        address.row,
        address.col,
        authorization.origin ?? 'instanceApi',
      ),
    ).allowed
  );
};

const defaultShapeId = (range: Range, shape: SessionShapeKind): string =>
  `shape-${range.sheet}-${range.r0}-${range.c0}-${shape}`;

const defaultImageId = (range: Range): string => `image-${range.sheet}-${range.r0}-${range.c0}`;

export function createSessionShape(
  store: SpreadsheetStore,
  range: Range,
  options: CreateSessionShapeOptions,
  history: History | null = null,
): SessionIllustration | null {
  if (
    isSheetProtected(store.getState(), range.sheet) ||
    !canMutateRange(store, range, options.authorization, 'shapeInsert')
  )
    return null;
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
  recordIllustrationChange(
    history,
    store,
    () => mutators.upsertIllustration(store, item),
    imageIntent(
      options.authorization?.commandId ?? 'shapeInsert',
      range.sheet,
      range.r0,
      range.c0,
      options.authorization?.origin ?? 'instanceApi',
    ),
  );
  return item;
}

export function createSessionImage(
  store: SpreadsheetStore,
  range: Range,
  options: CreateSessionImageOptions,
  history: History | null = null,
): SessionIllustration | null {
  if (
    isSheetProtected(store.getState(), range.sheet) ||
    !canMutateRange(store, range, options.authorization, 'pictureInsert')
  )
    return null;
  const item: SessionIllustration = {
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
    nativeWidthEmu = Math.max(1, Math.round((options.w ?? probe.pxWidth) * 9525));
    nativeHeightEmu = Math.max(1, Math.round((options.h ?? probe.pxHeight) * 9525));
    nativeObjectId = options.wb.insertImage(range.sheet, options.bytes, {
      row: range.r0,
      col: range.c0,
      name: options.alt ?? '',
      descr: options.alt ?? '',
      widthEmu: nativeWidthEmu,
      heightEmu: nativeHeightEmu,
    });
    if (nativeObjectId === null) return null;
    nativeFormat = probe.format;
  }
  const requestedAnchor: EngineDrawingAnchor = {
    anchorKind: AnchorKind.OneCell,
    editAs: AnchorEditAs.OneCell,
    fromRow: range.r0,
    fromCol: range.c0,
    fromRowOff: 0,
    fromColOff: 0,
    toRow: 0,
    toCol: 0,
    toRowOff: 0,
    toColOff: 0,
    cx: nativeWidthEmu ?? 1,
    cy: nativeHeightEmu ?? 1,
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
      const intent = imageIntent(
        options.authorization?.commandId ?? 'pictureInsert',
        inserted.sheet,
        inserted.nativeAnchor.fromRow,
        inserted.nativeAnchor.fromCol,
        options.authorization?.origin ?? 'instanceApi',
      );
      history.push({
        undo: () => {
          const current = sessionIllustrationById(store.getState(), persistedItem.id);
          const currentPersisted = current ? persistentImageFromState(current) : null;
          if (
            !currentPersisted ||
            !current ||
            !canMutateIllustration(
              store,
              current,
              { origin: 'undo', commandId: 'pictureInsert' },
              'pictureInsert',
            ) ||
            !nativeWb.removeDrawing(currentPersisted.sheet, currentPersisted.nativeObjectId)
          ) {
            throw new Error('image create undo failed');
          }
          mutators.removeIllustration(store, persistedItem.id);
          if (previous) {
            if (previousPersisted && previousSnapshot) {
              const restoredId = nativeWb.restoreDrawing(previousPersisted.sheet, previousSnapshot);
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
            current &&
            !canMutateIllustration(
              store,
              current,
              { origin: 'redo', commandId: 'pictureInsert' },
              'pictureInsert',
            )
          ) {
            throw new Error('image create redo blocked by protection or policy');
          }
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
        intent: { ...intent, origin: 'redo' },
        inverseIntent: { ...intent, origin: 'undo' },
      });
    }
    return persistedItem;
  }
  recordIllustrationChange(
    history,
    store,
    () => mutators.upsertIllustration(store, persistedItem),
    imageIntent(
      options.authorization?.commandId ?? 'pictureInsert',
      range.sheet,
      range.r0,
      range.c0,
      options.authorization?.origin ?? 'instanceApi',
    ),
  );
  return persistedItem;
}

export function createRibbonShapeFromSelection(
  store: SpreadsheetStore,
  range: Range,
  shape: SessionShapeKind,
  history: History | null = null,
  authorization?: IllustrationMutationAuthorization,
): SessionIllustration | null {
  const count = store.getState().illustrations.illustrations.length;
  return createSessionShape(
    store,
    range,
    {
      id: `ribbon-shape-${range.sheet}-${range.r0}-${range.c0}-${shape}-${count}`,
      shape,
      x: 300 + (count % 3) * 24,
      y: 88 + (count % 3) * 24,
      w: shape === 'line' || shape === 'arrow' ? 180 : 160,
      h: shape === 'line' || shape === 'arrow' ? 80 : 96,
      color: '#0f6cbd',
      authorization,
    },
    history,
  );
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
  return createSessionImage(
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

export function clearSessionIllustration(
  store: SpreadsheetStore,
  id: string,
  history: History | null = null,
  wb?: WorkbookHandle,
  authorization?: IllustrationMutationAuthorization,
): boolean {
  const item = sessionIllustrationById(store.getState(), id);
  if (!item) return false;
  if (
    isSheetProtected(store.getState(), item.sheet) ||
    !canMutateIllustration(store, item, authorization, 'pictureDelete') ||
    (item.nativeObjectId !== undefined && item.nativeGeometryEditable !== true)
  )
    return false;
  const persisted = persistentImageFromState(item);
  const snapshot =
    persisted && wb ? wb.snapshotDrawing(item.sheet, persisted.nativeObjectId) : null;
  if (persisted && (!wb || !snapshot || !wb.removeDrawing(item.sheet, persisted.nativeObjectId))) {
    return false;
  }
  if (persisted && wb && snapshot) {
    mutators.removeIllustration(store, id);
    if (history && !history.isReplaying()) {
      const intent = imageIntent(
        authorization?.commandId ?? 'pictureDelete',
        persisted.sheet,
        persisted.nativeAnchor.fromRow,
        persisted.nativeAnchor.fromCol,
        authorization?.origin ?? 'instanceApi',
      );
      history.push({
        undo: () => {
          if (
            !canMutateIllustration(
              store,
              item,
              { origin: 'undo', commandId: 'pictureDelete' },
              'pictureDelete',
            )
          )
            throw new Error('image undo blocked by protection or policy');
          const nativeObjectId = wb.restoreDrawing(item.sheet, snapshot);
          if (nativeObjectId === null) throw new Error('image undo restore failed');
          mutators.upsertIllustration(store, { ...item, nativeObjectId });
        },
        redo: () => {
          if (
            !canMutateIllustration(
              store,
              item,
              { origin: 'redo', commandId: 'pictureDelete' },
              'pictureDelete',
            )
          )
            throw new Error('image redo blocked by protection or policy');
          const current = persistentImageFromState(
            sessionIllustrationById(store.getState(), id) ?? item,
          );
          if (!current || !wb.removeDrawing(item.sheet, current.nativeObjectId)) {
            throw new Error('image redo delete failed');
          }
          mutators.removeIllustration(store, id);
        },
        intent: { ...intent, origin: 'redo' },
        inverseIntent: { ...intent, origin: 'undo' },
      });
    }
    return true;
  }
  recordIllustrationChange(
    history,
    store,
    () => mutators.removeIllustration(store, id),
    imageIntent(
      authorization?.commandId ?? 'pictureDelete',
      item.sheet,
      item.nativeAnchor?.fromRow ?? 0,
      item.nativeAnchor?.fromCol ?? 0,
      authorization?.origin ?? 'instanceApi',
    ),
  );
  return true;
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
  if (
    isSheetProtected(store.getState(), item.sheet) ||
    !canMutateIllustration(store, item, authorization, 'pictureUpdate') ||
    (item.nativeObjectId !== undefined && item.nativeGeometryEditable !== true)
  )
    return null;
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
      const intent = imageIntent(
        authorization?.commandId ?? 'pictureUpdate',
        persisted.sheet,
        persisted.nativeAnchor.fromRow,
        persisted.nativeAnchor.fromCol,
        authorization?.origin ?? 'instanceApi',
      );
      history.push({
        undo: () => {
          const current = sessionIllustrationById(store.getState(), id);
          if (!current) throw new Error('image undo state missing');
          if (
            !canMutateIllustration(
              store,
              current,
              { origin: 'undo', commandId: 'pictureUpdate' },
              'pictureUpdate',
            )
          )
            throw new Error('image undo blocked by protection or policy');
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
          if (
            !canMutateIllustration(
              store,
              current,
              { origin: 'redo', commandId: 'pictureUpdate' },
              'pictureUpdate',
            )
          )
            throw new Error('image redo blocked by protection or policy');
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
        intent: { ...intent, origin: 'redo' },
        inverseIntent: { ...intent, origin: 'undo' },
      });
    }
    return sessionIllustrationById(store.getState(), id);
  } else if (persisted && !wb) {
    return item;
  }
  recordIllustrationChange(
    history,
    store,
    () => mutators.updateIllustration(store, id, patch as SessionIllustrationPatch),
    imageIntent(
      authorization?.commandId ?? 'pictureUpdate',
      item.sheet,
      item.nativeAnchor?.fromRow ?? 0,
      item.nativeAnchor?.fromCol ?? 0,
      authorization?.origin ?? 'instanceApi',
    ),
  );
  return sessionIllustrationById(store.getState(), id);
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
  if (
    isSheetProtected(state, item.sheet) ||
    !canMutateIllustration(store, item, authorization, 'pictureDuplicate') ||
    (item.nativeObjectId !== undefined && item.nativeGeometryEditable !== true)
  )
    return null;
  const copy: SessionIllustration = {
    ...item,
    id: duplicateId(state, id),
    x: (item.x ?? 300) + DUPLICATE_OFFSET_PX,
    y: (item.y ?? 88) + DUPLICATE_OFFSET_PX,
  };
  const persisted = persistentImageFromState(item);
  let copySnapshot: Uint8Array | null = null;
  if (persisted && wb) {
    const source = wb.snapshotDrawing(item.sheet, persisted.nativeObjectId);
    const newId = source ? wb.restoreDrawing(item.sheet, source, { newId: true }) : null;
    if (newId === null) return null;
    const at = sessionPosition(store, persisted.nativeAnchor);
    const offset = placementForPixels(store, persisted.nativeAnchor, {
      x: at.x + DUPLICATE_OFFSET_PX,
      y: at.y + DUPLICATE_OFFSET_PX,
    });
    copySnapshot = wb.setDrawingAnchor(item.sheet, newId, offset)
      ? wb.snapshotDrawing(item.sheet, newId)
      : null;
    if (!copySnapshot) {
      wb.removeDrawing(item.sheet, newId);
      return null;
    }
    Object.assign(copy as object, {
      nativeObjectId: newId,
      imageBytes: Uint8Array.from(persisted.imageBytes),
      imageFormat: persisted.imageFormat,
      nativeAnchor: nativeImageAnchor(wb, item.sheet, newId) ?? persisted.nativeAnchor,
    });
  } else if (persisted) {
    return null;
  }
  if (persisted && wb && copySnapshot) {
    const redoSnapshot = copySnapshot;
    mutators.upsertIllustration(store, copy);
    if (history && !history.isReplaying()) {
      const intent = imageIntent(
        authorization?.commandId ?? 'pictureDuplicate',
        persisted.sheet,
        persisted.nativeAnchor.fromRow,
        persisted.nativeAnchor.fromCol,
        authorization?.origin ?? 'instanceApi',
      );
      history.push({
        undo: () => {
          if (
            !canMutateIllustration(
              store,
              copy,
              { origin: 'undo', commandId: 'pictureDuplicate' },
              'pictureDuplicate',
            )
          )
            throw new Error('image duplicate undo blocked by protection or policy');
          const current = persistentImageFromState(
            sessionIllustrationById(store.getState(), copy.id) ?? copy,
          );
          if (!current || !wb.removeDrawing(copy.sheet, current.nativeObjectId)) {
            throw new Error('image duplicate undo failed');
          }
          mutators.removeIllustration(store, copy.id);
        },
        redo: () => {
          if (
            !canMutateIllustration(
              store,
              copy,
              { origin: 'redo', commandId: 'pictureDuplicate' },
              'pictureDuplicate',
            )
          )
            throw new Error('image duplicate redo blocked by protection or policy');
          const nativeObjectId = wb.restoreDrawing(copy.sheet, redoSnapshot);
          if (nativeObjectId === null) throw new Error('image duplicate redo failed');
          Object.assign(copy as object, { nativeObjectId });
          mutators.upsertIllustration(store, copy);
        },
        intent: { ...intent, origin: 'redo' },
        inverseIntent: { ...intent, origin: 'undo' },
      });
    }
    return copy;
  }
  recordIllustrationChange(
    history,
    store,
    () => mutators.upsertIllustration(store, copy),
    imageIntent(
      authorization?.commandId ?? 'pictureDuplicate',
      item.sheet,
      item.nativeAnchor?.fromRow ?? 0,
      item.nativeAnchor?.fromCol ?? 0,
      authorization?.origin ?? 'instanceApi',
    ),
  );
  return copy;
}

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
    .filter((candidate) => candidate.sheet === item.sheet && candidate.nativeObjectId !== undefined)
    .findIndex((candidate) => candidate.id === item.id);
  const slots = objects.flatMap((object, index) => (object.imageFormat !== 0 ? [index] : []));
  const to = slots[rank];
  return rank < 0 || to === undefined || to === from ? null : { from, to };
};

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
  if (
    isSheetProtected(state, item.sheet) ||
    !canMutateIllustration(store, item, authorization, 'pictureArrange')
  )
    return null;
  const arranged = reorderIllustrations(state.illustrations.illustrations, id, action);
  if (!arranged) return null;
  const intent = imageIntent(
    authorization?.commandId ?? 'pictureArrange',
    item.sheet,
    item.nativeAnchor?.fromRow ?? 0,
    item.nativeAnchor?.fromCol ?? 0,
    authorization?.origin ?? 'instanceApi',
  );
  if (item.nativeObjectId !== undefined) {
    if (!wb?.capabilities.drawings) return null;
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
        if (
          !current ||
          !canMutateIllustration(store, current, { origin, commandId: 'pictureArrange' }) ||
          !wb.setDrawingZOrder(item.sheet, nativeObjectId, index)
        ) {
          throw new Error(`image arrange ${origin} failed`);
        }
        apply(list);
      };
      history.push({
        undo: () => replay('undo', z.from, before),
        redo: () => replay('redo', z.to, arranged),
        intent: { ...intent, origin: 'redo' },
        inverseIntent: { ...intent, origin: 'undo' },
      });
    }
    return sessionIllustrationById(store.getState(), id);
  }
  recordIllustrationChange(
    history,
    store,
    () => {
      store.setState((state) => {
        return { ...state, illustrations: { illustrations: arranged } };
      });
    },
    intent,
  );
  return sessionIllustrationById(store.getState(), id);
}
