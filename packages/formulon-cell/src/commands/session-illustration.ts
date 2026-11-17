import { AnchorEditAs, AnchorKind, type Color, ColorKind } from '@libraz/formulon';
import {
  cellOrigin,
  DUPLICATE_OFFSET_PX,
  placementForPixels,
  sessionPosition,
} from '../engine/drawing-geometry.js';
import {
  hasPictureAnchorChange,
  hasPictureTransform,
  movePersistentImage,
  nativeImageAnchor,
  persistentImageFromState,
  transformPersistentImage,
} from '../engine/images-sync.js';
import { captureShapeInput } from '../engine/shapes-sync.js';
import { cssColorFromArgb } from '../engine/theme-color.js';
import type { EngineDrawingAnchor, Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { EngineShapeInput } from '../engine/workbook-handle-drawing-objects.js';
import { defaultStrings, type Strings } from '../i18n/strings.js';
import {
  type DrawingStackEntry,
  type DrawingState,
  mintDrawingId,
  sheetDrawingStack,
  writeDrawingOrder,
} from '../store/drawing-order.js';
import {
  mutators,
  type SessionIllustration,
  type SessionShapeKind,
  type SpreadsheetStore,
  type State,
} from '../store/store.js';
import type { History } from './history.js';
import type { InteractionOrigin } from './interaction-policy.js';
import { type ObjectActionContext, objectWrite } from './object-actions.js';
import { captureIllustrationsSnapshot } from './slice-history.js';

export { type DrawingStackEntry, sheetDrawingStack } from '../store/drawing-order.js';

export interface CreateSessionShapeOptions {
  id?: string;
  shape: SessionShapeKind;
  x?: number;
  y?: number;
  w?: number;
  h?: number;
  color?: string;
  cornerAdj?: number;
  /** Outline weight in points; 0 is No Outline. */
  lineWidth?: number;
  opacity?: number;
  /** The shape's name, saved with it. */
  name?: string;
  /** Further fields carried on the shape. */
  extra?: Partial<SessionIllustration>;
  authorization?: IllustrationMutationAuthorization;
}

export interface CreateSessionImageOptions {
  id?: string;
  src: string;
  alt?: string;
  /** The picture's name, saved with it; the alt text when unset. */
  name?: string;
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

/** A saved picture or shape at a fixed position on the sheet: the engine can neither re-anchor it
 *  nor insert it afresh, though a snapshot of it restores. */
const isFixedInPlace = (item: SessionIllustration): boolean =>
  item.nativeObjectId !== undefined && item.nativeGeometryEditable !== true;

/** A picture fixed in place takes no edit; a shape there can still be restyled. */
const isGeometryLockedPicture = (item: SessionIllustration): boolean =>
  item.kind === 'image' && isFixedInPlace(item);

/** Whether the illustration can be deleted. A shape drawn only as its
 *  bounding box cannot be put back faithfully by an undo, so it stays. */
export const canDeleteIllustration = (item: SessionIllustration): boolean => !item.nativeLocked;

/** Whether a picture or shape can be inserted afresh: one fixed in place cannot be placed anew,
 *  and a paint the cell cannot state would be rewritten. */
export const canReinsertIllustration = (item: SessionIllustration): boolean =>
  !isFixedInPlace(item) && !item.nativeFillOther && !item.nativeLineOther;

/** Whether the illustration can be copied: one that can be deleted and inserted afresh. */
export const canDuplicateIllustration = (item: SessionIllustration): boolean =>
  canDeleteIllustration(item) && canReinsertIllustration(item);

/** Why a picture or shape cannot be copied, for the notice a refused copy shows; null when it can be. */
export const illustrationCopyRefusal = (
  item: SessionIllustration,
): 'drawing-fixed-position' | 'shape-locked' | 'shape-paint-kept' | null => {
  if (item.kind === 'shape' && item.nativeLocked) return 'shape-locked';
  if (isFixedInPlace(item)) return 'drawing-fixed-position';
  return canReinsertIllustration(item) ? null : 'shape-paint-kept';
};

/** The illustrations with the drawing order, which a history step restores together. */
const captureDrawings = (state: State) => ({
  items: captureIllustrationsSnapshot(state),
  order: state.drawings.order,
});

type DrawingsSnapshot = ReturnType<typeof captureDrawings>;

const sameDrawings = (a: DrawingsSnapshot, b: DrawingsSnapshot): boolean =>
  JSON.stringify(a) === JSON.stringify(b);

const restoreDrawings = (store: SpreadsheetStore, snap: DrawingsSnapshot): void =>
  mutators.restoreIllustrations(
    store,
    snap.items.map((item) => ({ ...item })),
    snap.order,
  );

const recordIllustrationChange = (
  history: History | null,
  store: SpreadsheetStore,
  mutate: () => void,
): void => {
  if (!history || history.isReplaying()) {
    mutate();
    return;
  }
  const before = captureDrawings(store.getState());
  mutate();
  const after = captureDrawings(store.getState());
  if (sameDrawings(before, after)) return;
  history.push({
    undo: () => restoreDrawings(store, before),
    redo: () => restoreDrawings(store, after),
  });
};

/** Add `item` on top of its sheet's drawings. */
const addIllustration = (store: SpreadsheetStore, item: SessionIllustration): void => {
  mutators.raiseDrawing(store, item.id);
  mutators.upsertIllustration(store, item);
};

/** Put `order` back and restack the engine's objects of `sheet` to match. */
export const settleDrawingOrder = (
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  sheet: number,
  order: readonly string[],
): void => {
  mutators.setDrawingOrder(store, order);
  if (!writeDrawingOrder(wb, store.getState(), sheet)) {
    throw new Error('drawing order restore failed');
  }
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
        ...options.extra,
        id: options.id ?? defaultShapeId(range, options.shape),
        kind: 'shape',
        shape: options.shape,
        sheet: range.sheet,
        x: options.x,
        y: options.y,
        w: options.w,
        h: options.h,
        color: options.color,
        cornerAdj: options.cornerAdj,
        lineWidth: options.lineWidth,
        opacity: options.opacity,
        ...(options.name !== undefined ? { name: options.name } : {}),
      };
      recordIllustrationChange(history, store, () => addIllustration(store, item));
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
        ...(options.name !== undefined ? { name: options.name } : {}),
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
          ...((options.name ?? options.alt) ? { name: options.name ?? options.alt } : {}),
          ...(options.alt ? { description: options.alt } : {}),
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
        const orderBefore = store.getState().drawings.order;
        addIllustration(store, persistedItem);
        writeDrawingOrder(nativeWb, store.getState(), range.sheet);
        const orderAfter = store.getState().drawings.order;
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
              settleDrawingOrder(store, nativeWb, range.sheet, orderBefore);
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
              settleDrawingOrder(store, nativeWb, range.sheet, orderAfter);
            },
          });
        }
        return persistedItem;
      }
      recordIllustrationChange(history, store, () => addIllustration(store, persistedItem));
      return persistedItem;
    },
  );
}

/** Name templates of the drawing kinds; `{n}` is the drawing's number among its kind. */
export type IllustrationNameTemplates = Pick<
  Strings['sessionIllustrations'],
  | 'pictureName'
  | 'rectangleName'
  | 'roundedRectangleName'
  | 'ovalName'
  | 'triangleName'
  | 'diamondName'
  | 'lineName'
  | 'arrowName'
  | 'textBoxName'
>;

const DEFAULT_NAMES: IllustrationNameTemplates = defaultStrings.sessionIllustrations;

const SHAPE_NAME_TEMPLATE: Record<SessionShapeKind, keyof IllustrationNameTemplates> = {
  rectangle: 'rectangleName',
  'rounded-rectangle': 'roundedRectangleName',
  oval: 'ovalName',
  triangle: 'triangleName',
  diamond: 'diamondName',
  line: 'lineName',
  arrow: 'arrowName',
};

type NamedKind = Pick<SessionIllustration, 'kind' | 'shape' | 'textBox' | 'sheet'>;

const nameTemplate = (item: NamedKind, names: IllustrationNameTemplates): string =>
  item.kind === 'image'
    ? names.pictureName
    : item.textBox
      ? names.textBoxName
      : names[SHAPE_NAME_TEMPLATE[item.shape ?? 'rectangle']];

/** The numbered name of a drawing that carries none: its position among its kind on its sheet. */
export const defaultIllustrationName = (
  illustrations: readonly SessionIllustration[],
  item: SessionIllustration,
  names: IllustrationNameTemplates,
): string => {
  const sameKind = (entry: SessionIllustration): boolean =>
    entry.sheet === item.sheet &&
    entry.kind === item.kind &&
    entry.shape === item.shape &&
    !entry.textBox === !item.textBox;
  const index = illustrations.filter(sameKind).indexOf(item);
  return nameTemplate(item, names).replace('{n}', String(index + 1));
};

/** A drawing's name: the one it carries, else its numbered default. */
export const illustrationName = (
  illustrations: readonly SessionIllustration[],
  item: SessionIllustration,
  names: IllustrationNameTemplates,
): string => item.name ?? defaultIllustrationName(illustrations, item, names);

/** The name a new drawing takes: one past the highest number its kind's
 *  template already shows on the sheet. */
export const newIllustrationName = (
  illustrations: readonly SessionIllustration[],
  item: NamedKind,
  names: IllustrationNameTemplates,
): string => {
  const template = nameTemplate(item, names);
  const [head = '', tail = ''] = template.split('{n}');
  let highest = 0;
  for (const entry of illustrations) {
    if (entry.sheet !== item.sheet) continue;
    const name = illustrationName(illustrations, entry, names);
    if (!name.startsWith(head) || !name.endsWith(tail)) continue;
    const digits = name.slice(head.length, name.length - tail.length);
    if (/^\d+$/.test(digits)) highest = Math.max(highest, Number(digits));
  }
  return template.replace('{n}', String(highest + 1));
};

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
  names: IllustrationNameTemplates = DEFAULT_NAMES,
): SessionIllustration | null {
  const state = store.getState();
  const count = state.illustrations.illustrations.length;
  const item = createSessionShape(
    store,
    range,
    {
      id: mintDrawingId(
        state,
        `ribbon-shape-${range.sheet}-${range.r0}-${range.c0}-${shape}-${count}`,
      ),
      shape,
      name: newIllustrationName(
        state.illustrations.illustrations,
        { kind: 'shape', shape, sheet: range.sheet },
        names,
      ),
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

/** A new text box's fill: theme light 1. */
const TEXT_BOX_FILL: Color = {
  kind: ColorKind.Theme,
  argb: 0,
  theme: 0,
  tint: 0,
  indexed: 0,
};
/** A new text box's outline, light 1 shaded 50% as Excel draws it; the
 *  engine stores no shade, so it is written as RGB. */
const TEXT_BOX_LINE = '#bcbcbc';

export function createRibbonTextBoxFromSelection(
  store: SpreadsheetStore,
  range: Range,
  history: History | null = null,
  authorization?: IllustrationMutationAuthorization,
  wb?: WorkbookHandle,
  names: IllustrationNameTemplates = DEFAULT_NAMES,
): SessionIllustration | null {
  const state = store.getState();
  const count = state.illustrations.illustrations.length;
  const argb = wb?.resolveColorArgb(TEXT_BOX_FILL) ?? null;
  const fill = argb === null ? '#ffffff' : cssColorFromArgb(argb);
  const item = createSessionShape(
    store,
    range,
    {
      id: mintDrawingId(state, `ribbon-text-box-${range.sheet}-${range.r0}-${range.c0}-${count}`),
      shape: 'rectangle',
      name: newIllustrationName(
        state.illustrations.illustrations,
        { kind: 'shape', shape: 'rectangle', textBox: true, sheet: range.sheet },
        names,
      ),
      x: 300 + (count % 3) * 24,
      y: 88 + (count % 3) * 24,
      w: 160,
      h: 96,
      color: fill,
      lineWidth: 0.75,
      extra: {
        textBox: true,
        nativeFill: { spec: TEXT_BOX_FILL, css: fill },
        lineColor: TEXT_BOX_LINE,
      },
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
  names: IllustrationNameTemplates = DEFAULT_NAMES,
): SessionIllustration | null {
  const state = store.getState();
  const count = state.illustrations.illustrations.length;
  const item = createSessionImage(
    store,
    range,
    {
      id: mintDrawingId(state, `ribbon-image-${range.sheet}-${range.r0}-${range.c0}-${count}`),
      src,
      alt: alt ?? src,
      name: newIllustrationName(
        state.illustrations.illustrations,
        { kind: 'image', sheet: range.sheet },
        names,
      ),
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
        const order = store.getState().drawings.order;
        mutators.removeIllustration(store, id);
        if (history && !history.isReplaying()) {
          history.push({
            undo: () => {
              const nativeObjectId = wb.restoreDrawing(item.sheet, snapshot);
              if (nativeObjectId === null) throw new Error('image undo restore failed');
              mutators.upsertIllustration(store, { ...item, nativeObjectId });
              settleDrawingOrder(store, wb, item.sheet, order);
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

/** The transform keys of `item`, each named so an absent flip reads as cleared. */
const transformOf = (
  item: SessionIllustration,
): Pick<SessionIllustration, 'rotationDeg' | 'flipH' | 'flipV'> => ({
  rotationDeg: item.rotationDeg,
  flipH: item.flipH,
  flipV: item.flipV,
});

/** Whether the picture or shape can turn and flip: a shape drawn only as its bounding box cannot. */
export const canTransformIllustration = (item: SessionIllustration): boolean =>
  (item.kind === 'shape' || item.kind === 'image') && !item.nativeLocked;

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
  // A picture fixed in place can still turn and flip, but not move.
  const moves = hasPictureAnchorChange(patch);
  const turns = hasPictureTransform(patch);
  if (isGeometryLockedPicture(item) && moves) return null;
  // Another sheet takes the shape as a new insert.
  if (patch.sheet !== undefined && patch.sheet !== item.sheet && !canReinsertIllustration(item)) {
    return null;
  }
  return objectWrite(
    illustrationEntry(store, history, authorization, 'pictureUpdate'),
    item.sheet,
    null,
    () => {
      const persisted = persistentImageFromState(item);
      if (persisted && wb) {
        const before = { ...item } as SessionIllustration;
        // The transform applies first, then the anchor resolves against it.
        if (turns && !transformPersistentImage(wb, persisted, patch)) return item;
        if (moves && !movePersistentImage(wb, persisted, patch, store)) {
          if (turns) transformPersistentImage(wb, persisted, transformOf(before));
          return item;
        }
        const after = {
          ...item,
          ...patch,
          nativeAnchor:
            nativeImageAnchor(wb, item.sheet, persisted.nativeObjectId) ?? persisted.nativeAnchor,
        } as SessionIllustration;
        mutators.updateIllustration(store, id, after as SessionIllustrationPatch);
        if (history && !history.isReplaying()) {
          const restore = (target: SessionIllustration, phase: 'undo' | 'redo'): void => {
            const current = sessionIllustrationById(store.getState(), id);
            if (!current) throw new Error(`image ${phase} state missing`);
            const currentPersisted = persistentImageFromState(current);
            if (!currentPersisted) throw new Error(`image ${phase} native state missing`);
            if (turns && !transformPersistentImage(wb, currentPersisted, transformOf(target))) {
              throw new Error(`image ${phase} restore failed`);
            }
            if (moves && !movePersistentImage(wb, currentPersisted, target, store)) {
              throw new Error(`image ${phase} restore failed`);
            }
            mutators.updateIllustration(store, id, {
              ...target,
              ...(turns ? transformOf(target) : {}),
              nativeAnchor:
                nativeImageAnchor(wb, item.sheet, currentPersisted.nativeObjectId) ??
                currentPersisted.nativeAnchor,
            });
          };
          history.push({
            undo: () => restore(before, 'undo'),
            redo: () => restore(after, 'redo'),
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
  /** A shape saved with the workbook, as the engine holds it. */
  readonly shape: EngineShapeInput | null;
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
  if (!persisted) {
    const shape =
      wb && item.kind === 'shape' && item.nativeObjectId !== undefined
        ? captureShapeInput(wb, item.sheet, item.nativeObjectId)
        : null;
    return { item: { ...item }, drawing: null, shape };
  }
  const drawing = wb?.snapshotDrawing(item.sheet, persisted.nativeObjectId) ?? null;
  return drawing ? { item: { ...item }, drawing, shape: null } : null;
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
          ...(source.shape && item.nativeAnchor
            ? {
                nativeSource: {
                  ...source.shape,
                  anchor: placementForPixels(store, item.nativeAnchor, to),
                },
              }
            : {}),
        }
      : {}),
  };
  if (!persisted) {
    recordIllustrationChange(history, store, () => addIllustration(store, copy));
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
  addIllustration(store, copy);
  writeDrawingOrder(wb, store.getState(), target.sheet);
  const order = store.getState().drawings.order;
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
        settleDrawingOrder(store, wb, copy.sheet, order);
      },
    });
  }
  return copy;
};

/** A planned arrange step: the sheet's stack after it, back to front. */
const planDrawingArrange = (
  state: DrawingState,
  sheet: number,
  id: string,
  action: SessionIllustrationArrangeAction,
): readonly DrawingStackEntry[] | null => {
  const stack = sheetDrawingStack(state, sheet);
  const rank = stack.findIndex((candidate) => candidate.id === id);
  const entry = stack[rank];
  if (!entry) return null;
  const target =
    action === 'bring-front'
      ? stack.length - 1
      : action === 'send-back'
        ? 0
        : action === 'bring-forward'
          ? Math.min(stack.length - 1, rank + 1)
          : Math.max(0, rank - 1);
  if (target === rank) return null;
  const next = stack.filter((candidate) => candidate !== entry);
  next.splice(target, 0, entry);
  return next;
};

/** History steps that put the drawing order back and forward again. */
export interface DrawingOrderSteps {
  readonly undo: () => void;
  readonly redo: () => void;
}

/**
 * Move drawing `id` of `sheet` one step or to an end of its sheet's stack, the
 * way the spreadsheet's arrange commands do for every kind. The store order
 * changes and the engine follows it. Null when the drawing is already there or
 * the engine refuses.
 */
export function arrangeDrawing(
  store: SpreadsheetStore,
  sheet: number,
  id: string,
  action: SessionIllustrationArrangeAction,
  wb?: WorkbookHandle | null,
): DrawingOrderSteps | null {
  const state = store.getState();
  const stack = planDrawingArrange(state, sheet, id, action);
  if (!stack) return null;
  const before = state.drawings.order;
  mutators.restackDrawings(
    store,
    stack.map((entry) => entry.id),
  );
  const after = store.getState().drawings.order;
  const apply = (order: readonly string[]): boolean => {
    mutators.setDrawingOrder(store, order);
    return !wb || writeDrawingOrder(wb, store.getState(), sheet);
  };
  if (wb && !writeDrawingOrder(wb, store.getState(), sheet)) {
    apply(before);
    return null;
  }
  const replay = (order: readonly string[], phase: 'undo' | 'redo') => (): void => {
    if (!apply(order)) throw new Error(`drawing arrange ${phase} failed`);
  };
  return { undo: replay(before, 'undo'), redo: replay(after, 'redo') };
}

/** The z-index of each drawing panel on `sheet`: its place in the stack from 1,
 *  and the selected panel over every kind. */
export function drawingPanelZIndex(
  state: DrawingState,
  sheet: number,
  selectedId: string | null,
): (id: string) => string {
  const stack = sheetDrawingStack(state, sheet);
  return (id) =>
    String(id === selectedId ? stack.length + 1 : stack.findIndex((entry) => entry.id === id) + 1);
}

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
  const item = sessionIllustrationById(store.getState(), id);
  if (!item) return null;
  return objectWrite(
    illustrationEntry(store, history, authorization, 'pictureArrange'),
    item.sheet,
    null,
    () => {
      if (item.nativeObjectId !== undefined && !wb?.capabilities.drawings) return null;
      const steps = arrangeDrawing(store, item.sheet, id, action, wb);
      if (!steps) return null;
      if (history && !history.isReplaying()) history.push(steps);
      return sessionIllustrationById(store.getState(), id);
    },
  );
}
