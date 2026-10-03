import { addrKey, parseAddrKey } from '../engine/address.js';
import { flushFormatToEngine } from '../engine/cell-format-sync.js';
import type { Addr } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { CellFormat, SpreadsheetStore, State } from '../store/store.js';
import type { History, HistoryEntry } from './history.js';

type FormatMapSnapshot = ReadonlyMap<string, CellFormat | undefined>;
type CommentSnapshot = ReadonlyMap<string, { author: string; text: string } | null>;

const clone = <T>(value: T): T =>
  value === undefined || value === null ? value : structuredClone(value);

const same = (left: unknown, right: unknown): boolean => {
  if (Object.is(left, right)) return true;
  if (left === null || right === null || typeof left !== 'object' || typeof right !== 'object') {
    return false;
  }
  if (Array.isArray(left) || Array.isArray(right)) {
    if (!Array.isArray(left) || !Array.isArray(right) || left.length !== right.length) return false;
    return left.every((value, index) => same(value, right[index]));
  }
  const leftRecord = left as Record<string, unknown>;
  const rightRecord = right as Record<string, unknown>;
  const leftKeys = Object.keys(leftRecord);
  const rightKeys = Object.keys(rightRecord);
  if (leftKeys.length !== rightKeys.length) return false;
  return leftKeys.every(
    (key) => Object.hasOwn(rightRecord, key) && same(leftRecord[key], rightRecord[key]),
  );
};

const captureFormats = (
  state: State,
  targets: readonly Addr[],
): Map<string, CellFormat | undefined> => {
  const result = new Map<string, CellFormat | undefined>();
  for (const addr of targets) {
    const key = addrKey(addr);
    result.set(key, clone(state.format.formats.get(key)));
  }
  return result;
};

const applyFormats = (
  store: SpreadsheetStore,
  formats: FormatMapSnapshot,
  pending?: { readonly value: State['ui']['pendingFormat'] },
): void => {
  store.setState((state) => {
    const next = new Map(state.format.formats);
    for (const [key, format] of formats) {
      if (format === undefined) next.delete(key);
      else next.set(key, clone(format));
    }
    return {
      ...state,
      format: { ...state.format, formats: next },
      ui: pending === undefined ? state.ui : { ...state.ui, pendingFormat: clone(pending.value) },
    };
  });
};

const captureComments = (
  workbook: WorkbookHandle | null,
  targets: readonly Addr[],
): Map<string, { author: string; text: string } | null> => {
  const result = new Map<string, { author: string; text: string } | null>();
  if (!workbook?.capabilities.comments) return result;
  for (const addr of targets) {
    result.set(addrKey(addr), clone(workbook.getComment(addr.sheet, addr.row, addr.col)));
  }
  return result;
};

const sameComment = (
  left: { author: string; text: string } | null | undefined,
  right: { author: string; text: string } | null | undefined,
): boolean => same(left ?? null, right ?? null);

const writeComments = (workbook: WorkbookHandle, desired: CommentSnapshot): void => {
  for (const [key, comment] of desired) {
    const addr = parseAddrKey(key);
    if (!addr) continue;
    const { sheet, row, col } = addr;
    const current = workbook.getComment(sheet, row, col);
    if (sameComment(current, comment)) continue;
    const ok = workbook.setCommentEntry(
      sheet,
      row,
      col,
      comment?.author ?? '',
      comment?.text ?? '',
    );
    if (!ok) throw new Error('dialog comment engine write failed');
  }
};

const commentsFromFormats = (
  formats: FormatMapSnapshot,
  targets: readonly Addr[],
  keys?: ReadonlySet<string>,
): Map<string, { author: string; text: string } | null> => {
  const result = new Map<string, { author: string; text: string } | null>();
  for (const addr of targets) {
    const key = addrKey(addr);
    if (keys && !keys.has(key)) continue;
    const format = formats.get(key);
    result.set(
      key,
      format?.comment ? { author: format.commentAuthor ?? '', text: format.comment } : null,
    );
  }
  return result;
};

const commentChangedKeys = (
  before: FormatMapSnapshot,
  after: FormatMapSnapshot,
  changed: ReadonlySet<string>,
): Set<string> => {
  const result = new Set<string>();
  for (const key of changed) {
    const previous = before.get(key);
    const next = after.get(key);
    if (
      !same(previous?.comment, next?.comment) ||
      !same(previous?.commentAuthor, next?.commentAuthor)
    ) {
      result.add(key);
    }
  }
  return result;
};

const selectComments = (
  comments: CommentSnapshot,
  keys: ReadonlySet<string>,
): Map<string, { author: string; text: string } | null> => {
  const result = new Map<string, { author: string; text: string } | null>();
  for (const key of keys) {
    result.set(key, clone(comments.get(key) ?? null));
  }
  return result;
};

const changedFormatKeys = (before: FormatMapSnapshot, after: FormatMapSnapshot): Set<string> => {
  const keys = new Set([...before.keys(), ...after.keys()]);
  return new Set([...keys].filter((key) => !same(before.get(key), after.get(key))));
};

const restoreAfterFailure = (
  store: SpreadsheetStore,
  workbook: WorkbookHandle | null,
  sheet: number,
  formats: FormatMapSnapshot,
  pending: State['ui']['pendingFormat'],
  comments: CommentSnapshot,
): void => {
  applyFormats(store, formats, { value: pending });
  if (!workbook) return;
  try {
    writeComments(workbook, comments);
  } finally {
    flushFormatToEngine(workbook, store, sheet, { strict: true });
  }
};

export interface RecordDialogFormatChangeOptions {
  readonly history: History | null;
  readonly store: SpreadsheetStore;
  readonly workbook: WorkbookHandle | null;
  readonly sheet: number;
  readonly targets: readonly Addr[];
  readonly pendingBefore: State['ui']['pendingFormat'];
  readonly mutate: () => boolean;
  readonly repeat?: () => void;
  /** Defer repeat registration when this entry is inside an outer transaction. */
  readonly registerRepeat?: boolean;
}

/** Record a dialog format mutation with scoped store and engine replay. */
export function recordDialogFormatChange(options: RecordDialogFormatChangeOptions): boolean {
  const {
    history,
    store,
    workbook,
    sheet,
    targets,
    pendingBefore,
    mutate,
    repeat,
    registerRepeat = true,
  } = options;
  const beforeFormats = captureFormats(store.getState(), targets);
  const beforePending = clone(pendingBefore);
  const beforeComments = captureComments(workbook, targets);

  let afterFormats: Map<string, CellFormat | undefined>;
  let afterComments: Map<string, { author: string; text: string } | null>;
  try {
    if (!mutate()) return false;
    afterFormats = captureFormats(store.getState(), targets);
    const changed = changedFormatKeys(beforeFormats, afterFormats);
    if (changed.size === 0) {
      // Accepted no-op and pending-only formats are UI-side repeat seeds. They
      // have no engine publication or material history entry.
      if (registerRepeat) history?.setRepeat(repeat ?? null);
      return true;
    }

    const changedComments = commentChangedKeys(beforeFormats, afterFormats, changed);
    if (workbook) {
      flushFormatToEngine(workbook, store, sheet, { strict: true });
      if (workbook.capabilities.comments) {
        writeComments(workbook, commentsFromFormats(afterFormats, targets, changedComments));
        afterComments = selectComments(captureComments(workbook, targets), changedComments);
      } else {
        afterComments = new Map();
      }
    } else {
      afterComments = new Map();
    }

    const changedBefore = new Map<string, CellFormat | undefined>();
    const changedAfter = new Map<string, CellFormat | undefined>();
    for (const key of changed) {
      changedBefore.set(key, clone(beforeFormats.get(key)));
      changedAfter.set(key, clone(afterFormats.get(key)));
    }
    const entryBeforeComments = selectComments(beforeComments, changedComments);
    const entryAfterComments = afterComments;
    const entry: HistoryEntry = {
      undo: () => replay(changedBefore, entryBeforeComments),
      redo: () => replay(changedAfter, entryAfterComments),
      repeat: registerRepeat ? repeat : undefined,
    };
    if (history && !history.isReplaying()) history.push(entry);
    return true;
  } catch (error) {
    try {
      restoreAfterFailure(store, workbook, sheet, beforeFormats, beforePending, beforeComments);
    } catch (rollbackError) {
      throw new AggregateError(
        [error, rollbackError],
        'Dialog format change failed and its rollback failed',
        { cause: error },
      );
    }
    throw error;
  }

  function replay(formats: FormatMapSnapshot, comments: CommentSnapshot): void {
    const immediateFormats = captureFormats(store.getState(), targets);
    const immediatePending = clone(store.getState().ui.pendingFormat);
    const immediateComments = captureComments(workbook, targets);
    try {
      // A material dialog history entry must not overwrite a later pending
      // format created by an unrelated blank-cell edit.
      applyFormats(store, formats);
      if (workbook) {
        flushFormatToEngine(workbook, store, sheet, { strict: true });
        if (workbook.capabilities.comments) writeComments(workbook, comments);
      }
    } catch (error) {
      try {
        restoreAfterFailure(
          store,
          workbook,
          sheet,
          immediateFormats,
          immediatePending,
          immediateComments,
        );
      } catch (rollbackError) {
        throw new AggregateError(
          [error, rollbackError],
          'Dialog format replay failed and its rollback failed',
          { cause: error },
        );
      }
      throw error;
    }
  }
}
