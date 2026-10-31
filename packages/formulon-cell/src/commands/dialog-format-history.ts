import { addrKey, parseAddrKey } from '../engine/address.js';
import { flushFormatToEngine } from '../engine/cell-format-sync.js';
import { formatHistoryBridgeFor } from '../engine/format-history-bridge.js';
import {
  captureSyncedFormatKeyMembership,
  restoreSyncedFormatKeyMembership,
} from '../engine/synced-format-keys.js';
import type { Addr, SheetAxisFormatOverride, SheetFontOverride } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { CellFormat, SpreadsheetStore, State } from '../store/store.js';
import { DEFAULT_COMMENT_AUTHOR } from './comment-author.js';
import type { History, HistoryEntry } from './history.js';

type FormatMapSnapshot = ReadonlyMap<string, CellFormat | undefined>;
type NoteSnapshot = { author: string; text: string; visible?: true };
type CommentSnapshot = ReadonlyMap<string, NoteSnapshot | null>;
interface NativeXfSnapshotEntry {
  readonly directXfIndex: number;
  readonly effectiveXfIndex: number | null;
  /** Formulon's EffectiveStyleSource ordinal: Cell=0, Row=1, Column=2, Default=3. */
  readonly source: number | null;
}
type NativeXfSnapshot = ReadonlyMap<string, NativeXfSnapshotEntry>;
type AxisSnapshot = SheetAxisFormatOverride | undefined;
type SheetFontSnapshot = SheetFontOverride | undefined;

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

const cloneAxis = (value: AxisSnapshot): AxisSnapshot =>
  value === undefined
    ? undefined
    : {
        baseline: {
          columns: value.baseline.columns.map((entry) => ({ ...entry })),
          rows: value.baseline.rows.map((entry) => ({ ...entry })),
        },
        desired: {
          columns: value.desired.columns.map((entry) => ({ ...entry })),
          rows: value.desired.rows.map((entry) => ({ ...entry })),
        },
        ...(value.directXfs
          ? {
              directXfs: value.directXfs.map((entry) => ({
                key: entry.key,
                before: { ...entry.before },
                afterXfIndex: entry.afterXfIndex,
              })),
            }
          : {}),
        ...(value.fontPatch ? { fontPatch: { ...value.fontPatch } } : {}),
      };

const cloneSheetFont = (value: SheetFontSnapshot): SheetFontSnapshot =>
  value === undefined
    ? undefined
    : {
        patch: { ...value.patch },
        ...(value.clearAxes ? { clearAxes: true } : {}),
        ...(value.baseline
          ? {
              baseline: {
                columns: value.baseline.columns.map((entry) => ({ ...entry })),
                rows: value.baseline.rows.map((entry) => ({ ...entry })),
              },
            }
          : {}),
      };

const applyAxis = (store: SpreadsheetStore, sheet: number, value: AxisSnapshot): void => {
  store.setState((state) => {
    const next = new Map(state.format.sheetAxisOverrides ?? new Map());
    if (value === undefined) next.delete(sheet);
    else next.set(sheet, cloneAxis(value) as SheetAxisFormatOverride);
    return { ...state, format: { ...state.format, sheetAxisOverrides: next } };
  });
};

const applySheetFont = (store: SpreadsheetStore, sheet: number, value: SheetFontSnapshot): void => {
  store.setState((state) => {
    const next = new Map(state.format.sheetFontOverrides ?? new Map());
    if (value === undefined) next.delete(sheet);
    else next.set(sheet, cloneSheetFont(value) as SheetFontOverride);
    return { ...state, format: { ...state.format, sheetFontOverrides: next } };
  });
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
): Map<string, NoteSnapshot | null> => {
  const result = new Map<string, NoteSnapshot | null>();
  if (!workbook?.capabilities.comments) return result;
  for (const addr of targets) {
    result.set(addrKey(addr), clone(workbook.getComment(addr.sheet, addr.row, addr.col)));
  }
  return result;
};

const sameComment = (
  left: NoteSnapshot | null | undefined,
  right: NoteSnapshot | null | undefined,
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
      comment ? comment.visible === true : undefined,
    );
    if (!ok) throw new Error('dialog comment engine write failed');
  }
};

const commentsFromFormats = (
  formats: FormatMapSnapshot,
  targets: readonly Addr[],
  keys?: ReadonlySet<string>,
): Map<string, NoteSnapshot | null> => {
  const result = new Map<string, NoteSnapshot | null>();
  for (const addr of targets) {
    const key = addrKey(addr);
    if (keys && !keys.has(key)) continue;
    const format = formats.get(key);
    result.set(
      key,
      format?.comment
        ? {
            author: format.commentAuthor ?? DEFAULT_COMMENT_AUTHOR,
            text: format.comment,
            ...(format.commentVisible === true ? { visible: true as const } : {}),
          }
        : null,
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
      !same(previous?.commentAuthor, next?.commentAuthor) ||
      !same(previous?.commentVisible === true, next?.commentVisible === true)
    ) {
      result.add(key);
    }
  }
  return result;
};

const selectComments = (
  comments: CommentSnapshot,
  keys: ReadonlySet<string>,
): Map<string, NoteSnapshot | null> => {
  const result = new Map<string, NoteSnapshot | null>();
  for (const key of keys) {
    result.set(key, clone(comments.get(key) ?? null));
  }
  return result;
};

const changedFormatKeys = (before: FormatMapSnapshot, after: FormatMapSnapshot): Set<string> => {
  const keys = new Set([...before.keys(), ...after.keys()]);
  return new Set([...keys].filter((key) => !same(before.get(key), after.get(key))));
};

/** Capture the direct native XF row for the changed cells only. StoreMirror is
 * muted around dialog mutations/replays, so these indices still describe the
 * pre-flush engine state when the forward transaction asks for them. */
const captureNativeXfs = (
  workbook: WorkbookHandle | null,
  keys: ReadonlySet<string>,
): Map<string, NativeXfSnapshotEntry> => {
  const result = new Map<string, NativeXfSnapshotEntry>();
  if (!workbook?.capabilities?.cellFormatting) return result;
  for (const key of keys) {
    const addr = parseAddrKey(key);
    if (!addr) continue;
    const xfIndex = workbook.getCellXfIndex(addr.sheet, addr.row, addr.col);
    if (xfIndex === null) {
      throw new Error(`dialog native XF snapshot unavailable at ${key}`);
    }
    const effective = workbook.capabilities.effectiveStyle
      ? workbook.getEffectiveStyle(addr.sheet, addr.row, addr.col)
      : null;
    result.set(key, {
      directXfIndex: xfIndex,
      effectiveXfIndex: effective?.xfIndex ?? null,
      source: effective?.source ?? null,
    });
  }
  return result;
};

const restoreNativeXfs = (workbook: WorkbookHandle, snapshot: NativeXfSnapshot): void => {
  if (!workbook.capabilities?.cellFormatting) return;
  const clearCellXfIndex = (
    workbook as WorkbookHandle & {
      clearCellXfIndex?: (sheet: number, row: number, col: number) => boolean;
    }
  ).clearCellXfIndex;
  for (const [key, entry] of snapshot) {
    const addr = parseAddrKey(key);
    if (!addr) continue;
    // A non-cell source is inherited. Use the native clear operation when the
    // loaded adapter exposes it; older bundles retain the effective-XF
    // fallback so replay remains compatible with their API surface.
    const inherited = entry.source !== null && entry.source !== 0;
    const xfIndex = inherited ? entry.effectiveXfIndex : entry.directXfIndex;
    if (inherited && workbook.capabilities.cellFormattingClear === true) {
      if (
        typeof clearCellXfIndex !== 'function' ||
        !clearCellXfIndex.call(workbook, addr.sheet, addr.row, addr.col)
      )
        throw new Error(`dialog native XF clear restore failed at ${key}`);
      // Non-empty cells may remain cell-owned in the native resolver even
      // after a successful clear. Confirm the requested inherited source and
      // fall back to the captured effective XF only for that semantic
      // mismatch; a reported clear failure above remains fatal.
      const effective = workbook.getEffectiveStyle(addr.sheet, addr.row, addr.col);
      if (
        effective?.source === entry.source &&
        (entry.effectiveXfIndex === null || effective.xfIndex === entry.effectiveXfIndex)
      ) {
        continue;
      }
      if (xfIndex === null || !workbook.setCellXfIndex(addr.sheet, addr.row, addr.col, xfIndex)) {
        throw new Error(`dialog native XF fallback restore failed at ${key}`);
      }
      continue;
    }
    if (xfIndex === null) continue;
    if (!workbook.setCellXfIndex(addr.sheet, addr.row, addr.col, xfIndex)) {
      throw new Error(`dialog native XF restore failed at ${key}`);
    }
  }
};

const withEngineSyncMuted = <T>(workbook: WorkbookHandle | null, fn: () => T): T =>
  workbook && typeof workbook.withEngineSyncMuted === 'function'
    ? workbook.withEngineSyncMuted(fn)
    : fn();

const restoreAfterFailure = (
  store: SpreadsheetStore,
  workbook: WorkbookHandle | null,
  sheet: number,
  formats: FormatMapSnapshot,
  sheetFont: SheetFontSnapshot,
  axis: AxisSnapshot,
  pending: State['ui']['pendingFormat'],
  comments: CommentSnapshot,
  nativeXfs: NativeXfSnapshot,
  tracking: ReadonlyMap<string, boolean>,
): void => {
  withEngineSyncMuted(workbook, () => {
    applyFormats(store, formats, { value: pending });
    applySheetFont(store, sheet, sheetFont);
    applyAxis(store, sheet, axis);
    if (!workbook) return;

    const failures: unknown[] = [];
    let nativeRestored = false;
    try {
      flushFormatToEngine(workbook, store, sheet, { strict: true });
    } catch (error) {
      failures.push(error);
    }
    try {
      restoreNativeXfs(workbook, nativeXfs);
      nativeRestored = true;
    } catch (rollbackError) {
      failures.push(rollbackError);
    }
    try {
      writeComments(workbook, comments);
    } catch (commentError) {
      failures.push(commentError);
    }
    try {
      if (nativeRestored) restoreSyncedFormatKeyMembership(workbook, tracking);
    } catch (trackingError) {
      failures.push(trackingError);
    }
    if (failures.length === 1) throw failures[0];
    if (failures.length > 1) {
      throw new AggregateError(failures, 'Dialog format rollback failed', { cause: failures[0] });
    }
  });
};

const restoreStoreAfterFailure = (
  store: SpreadsheetStore,
  workbook: WorkbookHandle | null,
  sheet: number,
  formats: FormatMapSnapshot,
  sheetFont: SheetFontSnapshot,
  axis: AxisSnapshot,
  pending: State['ui']['pendingFormat'],
): void => {
  withEngineSyncMuted(workbook, () => {
    applyFormats(store, formats, { value: pending });
    applySheetFont(store, sheet, sheetFont);
    applyAxis(store, sheet, axis);
  });
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
    workbook: suppliedWorkbook,
    sheet,
    targets,
    pendingBefore,
    mutate,
    repeat,
    registerRepeat = true,
  } = options;
  const workbook = suppliedWorkbook ?? formatHistoryBridgeFor(store)?.workbook ?? null;
  const beforeFormats = captureFormats(store.getState(), targets);
  const beforeSheetFont = cloneSheetFont(store.getState().format.sheetFontOverrides?.get(sheet));
  const beforeAxis = cloneAxis(store.getState().format.sheetAxisOverrides?.get(sheet));
  const beforePending = clone(pendingBefore);
  const beforeComments = captureComments(workbook, targets);

  let afterFormats: Map<string, CellFormat | undefined>;
  let afterComments: Map<string, NoteSnapshot | null>;
  let beforeNativeXfs = new Map<string, NativeXfSnapshotEntry>();
  let beforeTracking = new Map<string, boolean>();
  let nativePublicationAttempted = false;
  try {
    if (!withEngineSyncMuted(workbook, mutate)) return false;
    afterFormats = captureFormats(store.getState(), targets);
    const afterSheetFont = cloneSheetFont(store.getState().format.sheetFontOverrides?.get(sheet));
    const afterAxis = cloneAxis(store.getState().format.sheetAxisOverrides?.get(sheet));
    const changed = changedFormatKeys(beforeFormats, afterFormats);
    const sheetFontChanged = !same(beforeSheetFont, afterSheetFont);
    const axisChanged = !same(beforeAxis, afterAxis);
    if (changed.size === 0 && !sheetFontChanged && !axisChanged) {
      // Accepted no-op and pending-only formats are UI-side repeat seeds. They
      // have no engine publication or material history entry.
      if (registerRepeat) history?.setRepeat(repeat ?? null);
      return true;
    }

    // The store mutation has completed, but its mirror was muted, so the
    // engine still carries the direct XF rows from before the dialog edit.
    const nativeKeys = new Set(changed);
    for (const entry of beforeAxis?.directXfs ?? []) nativeKeys.add(entry.key);
    for (const entry of afterAxis?.directXfs ?? []) nativeKeys.add(entry.key);
    beforeNativeXfs = captureNativeXfs(workbook, nativeKeys);
    if (workbook) beforeTracking = captureSyncedFormatKeyMembership(workbook, nativeKeys);
    const changedComments = commentChangedKeys(beforeFormats, afterFormats, changed);
    if (workbook) {
      withEngineSyncMuted(workbook, () => {
        nativePublicationAttempted = true;
        flushFormatToEngine(workbook, store, sheet, { strict: true });
        if (workbook.capabilities.comments) {
          writeComments(workbook, commentsFromFormats(afterFormats, targets, changedComments));
        }
      });
      if (workbook.capabilities.comments)
        afterComments = selectComments(captureComments(workbook, targets), changedComments);
      else afterComments = new Map();
    } else {
      afterComments = new Map();
    }
    const afterNativeXfs = captureNativeXfs(workbook, nativeKeys);

    const changedBefore = new Map<string, CellFormat | undefined>();
    const changedAfter = new Map<string, CellFormat | undefined>();
    for (const key of changed) {
      changedBefore.set(key, clone(beforeFormats.get(key)));
      changedAfter.set(key, clone(afterFormats.get(key)));
    }
    const entryBeforeComments = selectComments(beforeComments, changedComments);
    const entryAfterComments = afterComments;
    const entry: HistoryEntry = {
      undo: () =>
        replay(changedBefore, beforeSheetFont, beforeAxis, entryBeforeComments, beforeNativeXfs),
      redo: () =>
        replay(changedAfter, afterSheetFont, afterAxis, entryAfterComments, afterNativeXfs),
      repeat: registerRepeat ? repeat : undefined,
    };
    if (history && !history.isReplaying()) history.push(entry);
    return true;
  } catch (error) {
    try {
      if (nativePublicationAttempted) {
        restoreAfterFailure(
          store,
          workbook,
          sheet,
          beforeFormats,
          beforeSheetFont,
          beforeAxis,
          beforePending,
          beforeComments,
          beforeNativeXfs,
          beforeTracking,
        );
      } else {
        restoreStoreAfterFailure(
          store,
          workbook,
          sheet,
          beforeFormats,
          beforeSheetFont,
          beforeAxis,
          beforePending,
        );
      }
    } catch (rollbackError) {
      throw new AggregateError(
        [error, rollbackError],
        'Dialog format change failed and its rollback failed',
        { cause: error },
      );
    }
    throw error;
  }

  function replay(
    formats: FormatMapSnapshot,
    sheetFont: SheetFontSnapshot,
    axis: AxisSnapshot,
    comments: CommentSnapshot,
    nativeXfs: NativeXfSnapshot,
  ): void {
    const immediateFormats = captureFormats(store.getState(), targets);
    const immediateSheetFont = cloneSheetFont(
      store.getState().format.sheetFontOverrides?.get(sheet),
    );
    const immediateAxis = cloneAxis(store.getState().format.sheetAxisOverrides?.get(sheet));
    const immediatePending = clone(store.getState().ui.pendingFormat);
    const immediateComments = captureComments(workbook, targets);
    const nativeKeys = new Set(nativeXfs.keys());
    const immediateNativeXfs = captureNativeXfs(workbook, nativeKeys);
    const immediateTracking = workbook
      ? captureSyncedFormatKeyMembership(workbook, nativeKeys)
      : new Map<string, boolean>();
    try {
      // A material dialog history entry must not overwrite a later pending
      // format created by an unrelated blank-cell edit.
      withEngineSyncMuted(workbook, () => {
        applyFormats(store, formats);
        applySheetFont(store, sheet, sheetFont);
        applyAxis(store, sheet, axis);
        if (workbook) {
          flushFormatToEngine(workbook, store, sheet, { strict: true });
          restoreNativeXfs(workbook, nativeXfs);
          if (workbook.capabilities.comments) writeComments(workbook, comments);
        }
      });
    } catch (error) {
      try {
        restoreAfterFailure(
          store,
          workbook,
          sheet,
          immediateFormats,
          immediateSheetFont,
          immediateAxis,
          immediatePending,
          immediateComments,
          immediateNativeXfs,
          immediateTracking,
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
