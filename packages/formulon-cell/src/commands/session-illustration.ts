import { AnchorEditAs, AnchorKind } from '@libraz/formulon';
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

/** The colour of a shape that has not been given one. */
export const DEFAULT_SHAPE_COLOR = '#0f6cbd';

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
      color: DEFAULT_SHAPE_COLOR,
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
    !canMutateIllustration(store, item, authorization, 'pictureDuplicate')
  )
    return null;
  const clip = copySessionIllustration(store, id, wb);
  if (!clip) return null;
  const persisted = persistentImageFromState(item);
  const from = persisted
    ? sessionPosition(store, persisted.nativeAnchor)
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
    {
      commandId: authorization?.commandId ?? 'pictureDuplicate',
      origin: authorization?.origin ?? 'instanceApi',
      historyCommandId: 'pictureDuplicate',
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
  if (!item || (item.nativeObjectId !== undefined && item.nativeGeometryEditable !== true)) {
    return null;
  }
  const persisted = persistentImageFromState(item);
  if (!persisted) return { item: { ...item }, drawing: null };
  const drawing = wb?.snapshotDrawing(item.sheet, persisted.nativeObjectId) ?? null;
  return drawing ? { item: { ...item }, drawing } : null;
}

/**
 * Paste a copied picture or shape onto `at.sheet` with its top-left corner at
 * cell (row, col), as the spreadsheet does. Undo removes the pasted object.
 */
export function pasteSessionIllustration(
  store: SpreadsheetStore,
  clip: IllustrationClip,
  at: { sheet: number; row: number; col: number },
  history: History | null = null,
  wb?: WorkbookHandle,
  authorization?: IllustrationMutationAuthorization,
): SessionIllustration | null {
  const range = { sheet: at.sheet, r0: at.row, c0: at.col, r1: at.row, c1: at.col };
  if (
    isSheetProtected(store.getState(), at.sheet) ||
    !canMutateRange(store, range, authorization, 'picturePaste')
  )
    return null;
  return placeIllustrationCopy(
    store,
    clip,
    { ...at, id: duplicateId(store.getState(), clip.item.id) },
    cellOrigin(store, at.row, at.col),
    history,
    wb,
    {
      commandId: authorization?.commandId ?? 'picturePaste',
      origin: authorization?.origin ?? 'instanceApi',
      historyCommandId: 'picturePaste',
    },
  );
}

/** Add `source` as a new object `target.id` on `target.sheet` with its top-left
 *  corner at `to`, recorded as one undoable step. */
const placeIllustrationCopy = (
  store: SpreadsheetStore,
  source: IllustrationClip,
  target: { id: string; sheet: number; row: number; col: number },
  to: { x: number; y: number },
  history: History | null,
  wb: WorkbookHandle | undefined,
  command: { commandId: string; origin: InteractionOrigin; historyCommandId: string },
): SessionIllustration | null => {
  const { item } = source;
  const copy: SessionIllustration = {
    ...item,
    id: target.id,
    sheet: target.sheet,
    x: to.x,
    y: to.y,
  };
  const persisted = persistentImageFromState(item);
  const intent = imageIntent(
    command.commandId,
    target.sheet,
    target.row,
    target.col,
    command.origin,
  );
  if (!persisted) {
    recordIllustrationChange(
      history,
      store,
      () => mutators.upsertIllustration(store, copy),
      intent,
    );
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
    const commandId = command.historyCommandId;
    history.push({
      undo: () => {
        if (!canMutateIllustration(store, copy, { origin: 'undo', commandId }, commandId))
          throw new Error('image copy undo blocked by protection or policy');
        const current = persistentImageFromState(
          sessionIllustrationById(store.getState(), copy.id) ?? copy,
        );
        if (!current || !wb.removeDrawing(copy.sheet, current.nativeObjectId)) {
          throw new Error('image copy undo failed');
        }
        mutators.removeIllustration(store, copy.id);
      },
      redo: () => {
        if (!canMutateIllustration(store, copy, { origin: 'redo', commandId }, commandId))
          throw new Error('image copy redo blocked by protection or policy');
        const nativeObjectId = wb.restoreDrawing(copy.sheet, redoSnapshot);
        if (nativeObjectId === null) throw new Error('image copy redo failed');
        Object.assign(copy as object, { nativeObjectId });
        mutators.upsertIllustration(store, copy);
      },
      intent: { ...intent, origin: 'redo' },
      inverseIntent: { ...intent, origin: 'undo' },
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
