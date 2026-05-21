import type { Range } from '../engine/types.js';
import {
  mutators,
  type SessionIllustration,
  type SessionShapeKind,
  type SpreadsheetStore,
} from '../store/store.js';
import { type History, recordIllustrationsChange } from './history.js';
import { isSheetProtected } from './protection.js';

export interface CreateSessionShapeOptions {
  id?: string;
  shape: SessionShapeKind;
  x?: number;
  y?: number;
  w?: number;
  h?: number;
  color?: string;
  radius?: number;
}

export interface CreateSessionImageOptions {
  id?: string;
  src: string;
  alt?: string;
  x?: number;
  y?: number;
  w?: number;
  h?: number;
}

export type SessionIllustrationPatch = Partial<Omit<SessionIllustration, 'id'>>;
export type SessionIllustrationArrangeAction =
  | 'bring-forward'
  | 'send-backward'
  | 'bring-front'
  | 'send-back';

const defaultShapeId = (range: Range, shape: SessionShapeKind): string =>
  `shape-${range.sheet}-${range.r0}-${range.c0}-${shape}`;

const defaultImageId = (range: Range): string => `image-${range.sheet}-${range.r0}-${range.c0}`;

export function createSessionShape(
  store: SpreadsheetStore,
  range: Range,
  options: CreateSessionShapeOptions,
  history: History | null = null,
): SessionIllustration | null {
  if (isSheetProtected(store.getState(), range.sheet)) return null;
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
  };
  recordIllustrationsChange(history, store, () => {
    mutators.upsertIllustration(store, item);
  });
  return item;
}

export function createSessionImage(
  store: SpreadsheetStore,
  range: Range,
  options: CreateSessionImageOptions,
  history: History | null = null,
): SessionIllustration | null {
  if (isSheetProtected(store.getState(), range.sheet)) return null;
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
  recordIllustrationsChange(history, store, () => {
    mutators.upsertIllustration(store, item);
  });
  return item;
}

export function createRibbonShapeFromSelection(
  store: SpreadsheetStore,
  range: Range,
  shape: SessionShapeKind,
  history: History | null = null,
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
): boolean {
  const item = sessionIllustrationById(store.getState(), id);
  if (!item) return false;
  if (isSheetProtected(store.getState(), item.sheet)) return false;
  recordIllustrationsChange(history, store, () => {
    mutators.removeIllustration(store, id);
  });
  return true;
}

export function updateSessionIllustration(
  store: SpreadsheetStore,
  id: string,
  patch: SessionIllustrationPatch,
  history: History | null = null,
): SessionIllustration | null {
  const item = sessionIllustrationById(store.getState(), id);
  if (!item) return null;
  if (isSheetProtected(store.getState(), item.sheet)) return null;
  recordIllustrationsChange(history, store, () => {
    mutators.updateIllustration(store, id, patch);
  });
  return sessionIllustrationById(store.getState(), id);
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

export function arrangeSessionIllustration(
  store: SpreadsheetStore,
  id: string,
  action: SessionIllustrationArrangeAction,
  history: History | null = null,
): SessionIllustration | null {
  const state = store.getState();
  const item = sessionIllustrationById(state, id);
  if (!item) return null;
  if (isSheetProtected(state, item.sheet)) return null;
  const arranged = reorderIllustrations(state.illustrations.illustrations, id, action);
  if (!arranged) return null;
  recordIllustrationsChange(history, store, () => {
    store.setState((state) => {
      return { ...state, illustrations: { illustrations: arranged } };
    });
  });
  return sessionIllustrationById(store.getState(), id);
}
