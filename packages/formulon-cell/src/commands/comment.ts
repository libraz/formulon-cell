import { addrKey, parseAddrKey } from '../engine/address.js';
import type { Addr } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type {
  EngineMention,
  EnginePerson,
  EngineThreadedComment,
} from '../engine/workbook-handle-annotations.js';
import { mutators, type SpreadsheetStore, type State } from '../store/store.js';
import { DEFAULT_COMMENT_AUTHOR } from './comment-author.js';
import type { History } from './history.js';
import { interactionControllerFor } from './interaction-controller.js';
import type { InteractionOrigin, OperationEffect } from './interaction-policy.js';
import { isCellWritable, warnProtected } from './protection.js';

export interface CommentEntry {
  addr: Addr;
  text: string;
  author?: string;
}

export type ThreadedCommentEntry = EngineThreadedComment;
export type ThreadedPersonEntry = EnginePerson;

export interface ThreadedMutationAuthorization {
  origin?: InteractionOrigin;
  commandId?: string;
}

const commentIntentFor = (
  addr: Addr,
  authorization: ThreadedMutationAuthorization = {},
): {
  operation: 'comment';
  origin: InteractionOrigin;
  commandId: string;
  effects: readonly OperationEffect[];
} => ({
  operation: 'comment',
  origin: authorization.origin ?? 'instanceApi',
  commandId: authorization.commandId ?? 'commentDialog',
  effects: [{ kind: 'cells', cells: [addr] }],
});

const canMutateThreadedCell = (
  store: SpreadsheetStore,
  addr: Addr,
  authorization?: ThreadedMutationAuthorization,
): boolean => {
  if (!isCellWritable(store.getState(), addr)) {
    warnProtected(addr);
    return false;
  }
  const controller = interactionControllerFor(store);
  return !controller || controller.canExecute(commentIntentFor(addr, authorization)).allowed;
};

const cloneThread = (entry: ThreadedCommentEntry): ThreadedCommentEntry => ({
  ...entry,
  mentions: entry.mentions.map((mention) => ({ ...mention })),
});

const cloneThreads = (entries: readonly ThreadedCommentEntry[]): ThreadedCommentEntry[] =>
  entries.map(cloneThread);

const sameThreads = (
  a: readonly ThreadedCommentEntry[],
  b: readonly ThreadedCommentEntry[],
): boolean => JSON.stringify(a) === JSON.stringify(b);

/** Replace the active sheet's thread cache after workbook hydration. */
export function hydrateThreadedComments(
  store: SpreadsheetStore,
  sheet: number,
  entries: readonly ThreadedCommentEntry[],
): void {
  const snapshot = cloneThreads(entries);
  store.setState((state) => ({
    ...state,
    threadedComments: {
      ...state.threadedComments,
      bySheet: new Map(state.threadedComments.bySheet).set(sheet, snapshot),
    },
  }));
}

/** Read threaded comments, preferring the native workbook so callers always
 * see edits made by another command or by a workbook replacement. */
export function listThreadedComments(
  store: SpreadsheetStore,
  sheet = store.getState().data.sheetIndex,
  wb?: WorkbookHandle,
): ThreadedCommentEntry[] {
  if (wb?.capabilities.threadedComments) {
    const entries = wb.getThreadedComments(sheet);
    hydrateThreadedComments(store, sheet, entries);
    store.setState((state) => ({
      ...state,
      threadedComments: { ...state.threadedComments, persons: wb.getPersons() },
    }));
    return cloneThreads(entries);
  }
  return cloneThreads(store.getState().threadedComments.bySheet.get(sheet) ?? []);
}

export function threadedCommentsAt(
  store: SpreadsheetStore,
  addr: Addr,
  wb?: WorkbookHandle,
): ThreadedCommentEntry[] {
  return listThreadedComments(store, addr.sheet, wb).filter(
    (entry) => entry.parentId === '' && entry.row === addr.row && entry.col === addr.col,
  );
}

const guid = (): string => {
  const random = globalThis.crypto?.randomUUID?.();
  if (random) return `{${random.toUpperCase()}}`;
  const bytes = new Uint8Array(16);
  if (globalThis.crypto?.getRandomValues) globalThis.crypto.getRandomValues(bytes);
  else for (let i = 0; i < bytes.length; i += 1) bytes[i] = (i * 47 + Date.now()) & 0xff;
  const version = bytes[6] ?? 0;
  const variant = bytes[8] ?? 0;
  bytes[6] = (version & 0x0f) | 0x40;
  bytes[8] = (variant & 0x3f) | 0x80;
  const hex = Array.from(bytes, (value) => value.toString(16).padStart(2, '0')).join('');
  return `{${hex.slice(0, 8)}-${hex.slice(8, 12)}-${hex.slice(12, 16)}-${hex.slice(16, 20)}-${hex.slice(20)}}`.toUpperCase();
};

const timestamp = (): string => {
  const value = new Date().toISOString();
  return `${value.slice(0, 19)}.${value.slice(20, 22)}`;
};

const ensurePerson = (wb: WorkbookHandle, author: string): string => {
  const displayName = author.trim() || DEFAULT_COMMENT_AUTHOR;
  const existing = wb.getPersons().find((person) => person.displayName === displayName);
  if (existing) return existing.id;
  const id = guid();
  const person: EnginePerson = { id, displayName, userId: '', providerId: 'None' };
  if (!wb.addPerson(person)) {
    const retry = wb.getPersons().find((candidate) => candidate.displayName === displayName);
    if (retry) return retry.id;
    throw new Error('threaded comment author registration failed');
  }
  return id;
};

const threadedEntryAddress = (
  sheet: number,
  entries: readonly ThreadedCommentEntry[],
  id: string,
): Addr | null => {
  const entry = entries.find((candidate) => candidate.id === id);
  if (!entry) return null;
  const opening = entry.parentId
    ? (entries.find((candidate) => candidate.id === entry.parentId) ?? entry)
    : entry;
  return { sheet, row: opening.row, col: opening.col };
};

const refreshThreadCache = (store: SpreadsheetStore, wb: WorkbookHandle, sheet: number): void => {
  hydrateThreadedComments(store, sheet, wb.getThreadedComments(sheet));
  const persons = wb.getPersons();
  store.setState((state) => ({
    ...state,
    threadedComments: { ...state.threadedComments, persons },
  }));
};

/** Add a native thread opener at `addr`; replies use `replyToThread`. */
export function addThreadedComment(
  store: SpreadsheetStore,
  addr: Addr,
  text: string,
  wb: WorkbookHandle | undefined,
  options: {
    author?: string;
    replyToThread?: string;
    mentions?: readonly EngineMention[];
    authorization?: ThreadedMutationAuthorization;
  } = {},
): ThreadedCommentEntry | null {
  if (!wb?.capabilities.threadedComments || text.length === 0) return null;
  if (!canMutateThreadedCell(store, addr, options.authorization)) return null;
  // formulon rejects a thread and a legacy note on the same cell. Preserve
  // the note and avoid registering an unused person when the UI explicitly
  // requests a thread on such a cell.
  if (commentAt(store.getState(), addr) !== null) return null;
  let personId: string;
  try {
    personId = ensurePerson(wb, options.author ?? '');
  } catch {
    return null;
  }
  const entry: Partial<ThreadedCommentEntry> & Pick<ThreadedCommentEntry, 'id' | 'text'> = {
    id: guid(),
    row: addr.row,
    col: addr.col,
    personId,
    created: timestamp(),
    text,
    parentId: options.replyToThread ?? '',
    done: false,
    mentions: options.mentions ? options.mentions.map((mention) => ({ ...mention })) : [],
  };
  if (!wb.addThreadedComment(addr.sheet, entry)) return null;
  refreshThreadCache(store, wb, addr.sheet);
  return (
    listThreadedComments(store, addr.sheet, wb).find((candidate) => candidate.id === entry.id) ??
    null
  );
}

export function editThreadedComment(
  store: SpreadsheetStore,
  sheet: number,
  id: string,
  text: string,
  wb: WorkbookHandle | undefined,
  mentions: readonly EngineMention[] = [],
  authorization?: ThreadedMutationAuthorization,
): boolean {
  if (!wb?.capabilities.threadedComments || text.length === 0) return false;
  const address = threadedEntryAddress(sheet, wb.getThreadedComments(sheet), id);
  if (!address || !canMutateThreadedCell(store, address, authorization)) return false;
  if (!wb.editThreadedComment(sheet, id, text, mentions)) return false;
  refreshThreadCache(store, wb, sheet);
  return true;
}

export function resolveThread(
  store: SpreadsheetStore,
  sheet: number,
  threadId: string,
  done: boolean,
  wb: WorkbookHandle | undefined,
  authorization?: ThreadedMutationAuthorization,
): boolean {
  if (!wb?.capabilities.threadedComments) return false;
  const address = threadedEntryAddress(sheet, wb.getThreadedComments(sheet), threadId);
  if (!address || !canMutateThreadedCell(store, address, authorization)) return false;
  if (!wb.setThreadResolved(sheet, threadId, done)) return false;
  refreshThreadCache(store, wb, sheet);
  return true;
}

export function removeThreadedComment(
  store: SpreadsheetStore,
  sheet: number,
  id: string,
  wb: WorkbookHandle | undefined,
  authorization?: ThreadedMutationAuthorization,
): boolean {
  if (!wb?.capabilities.threadedComments) return false;
  const address = threadedEntryAddress(sheet, wb.getThreadedComments(sheet), id);
  if (!address || !canMutateThreadedCell(store, address, authorization)) return false;
  if (!wb.removeThreadedComment(sheet, id)) return false;
  refreshThreadCache(store, wb, sheet);
  return true;
}

const applyThreadSnapshot = (
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  sheet: number,
  snapshot: readonly ThreadedCommentEntry[],
  authorization: ThreadedMutationAuthorization = {},
): void => {
  const current = listThreadedComments(store, sheet, wb);
  const addresses = new Map<string, Addr>();
  for (const entry of [...current, ...snapshot]) {
    const address = threadedEntryAddress(sheet, [...current, ...snapshot], entry.id);
    if (address) addresses.set(addrKey(address), address);
  }
  for (const address of addresses.values()) {
    if (!canMutateThreadedCell(store, address, authorization)) {
      throw new Error('thread history replay blocked by protection or policy');
    }
  }
  const removeAll = (entries: readonly ThreadedCommentEntry[]): void => {
    // Replies must be removed before their opener: native removal of an
    // opener intentionally removes the whole thread.
    for (const entry of [...entries].sort(
      (a, b) => Number(a.parentId === '') - Number(b.parentId === ''),
    )) {
      if (!wb.removeThreadedComment(sheet, entry.id)) {
        throw new Error('thread history remove failed');
      }
    }
  };
  const restore = (entries: readonly ThreadedCommentEntry[]): void => {
    // The native API requires a reply's opener to exist first.
    for (const entry of [...entries].sort(
      (a, b) => Number(a.parentId !== '') - Number(b.parentId !== ''),
    )) {
      if (!wb.addThreadedComment(sheet, cloneThread(entry))) {
        throw new Error('thread history restore failed');
      }
    }
  };
  try {
    removeAll(current);
    restore(snapshot);
    refreshThreadCache(store, wb, sheet);
  } catch (error) {
    // A failed native add/remove must not leave a half-replayed thread in the
    // workbook. Reconcile the native state first, then restore the original
    // snapshot. If the engine rejects the rollback too, reflect its actual
    // state in the store before surfacing the original failure.
    try {
      removeAll(wb.getThreadedComments(sheet));
      restore(current);
      refreshThreadCache(store, wb, sheet);
    } catch {
      refreshThreadCache(store, wb, sheet);
    }
    throw error;
  }
};

/** Wrap a native threaded-comment mutation in the same undo boundary as the
 * legacy note commands. The snapshot contains original IDs, replies, mentions
 * and resolution, so undo/redo never synthesizes a new thread identity. */
export function recordThreadedChange<T>(
  history: History | null,
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  sheet: number,
  mutate: () => T,
  authorization: ThreadedMutationAuthorization = {},
): T {
  const before = listThreadedComments(store, sheet, wb);
  const result = mutate();
  const after = listThreadedComments(store, sheet, wb);
  if (history && !history.isReplaying() && !sameThreads(before, after)) {
    const addresses = new Map<string, Addr>();
    for (const entry of [...before, ...after]) {
      const source = entry.parentId
        ? ([...before, ...after].find((candidate) => candidate.id === entry.parentId) ?? entry)
        : entry;
      const address = { sheet, row: source.row, col: source.col };
      addresses.set(addrKey(address), address);
    }
    const effects: readonly OperationEffect[] = [{ kind: 'cells', cells: [...addresses.values()] }];
    const intent = {
      operation: 'comment' as const,
      origin: authorization.origin ?? 'instanceApi',
      commandId: authorization.commandId ?? 'commentDialog',
      effects,
    };
    history.push({
      undo: () =>
        applyThreadSnapshot(store, wb, sheet, before, { ...authorization, origin: 'undo' }),
      redo: () =>
        applyThreadSnapshot(store, wb, sheet, after, { ...authorization, origin: 'redo' }),
      intent: { ...intent, origin: 'redo' },
      inverseIntent: { ...intent, origin: 'undo' },
    });
  }
  return result;
}

type CommentSnapshot = Array<{ addr: Addr; text: string | null; author: string | null }>;
type PhysicalCommentSnapshot = Array<{
  addr: Addr;
  comment: { author: string; text: string } | null;
}>;

/** Read the comment text on a cell, or null when unset. */
export function commentAt(state: State, addr: Addr): string | null {
  const fmt = state.format.formats.get(addrKey(addr));
  const c = fmt?.comment;
  return typeof c === 'string' && c.length > 0 ? c : null;
}

/** Read the comment author on a cell, or null when unset. */
export function commentAuthorAt(state: State, addr: Addr): string | null {
  const fmt = state.format.formats.get(addrKey(addr));
  const author = fmt?.commentAuthor;
  return typeof author === 'string' && author.length > 0 ? author : null;
}

/** List non-empty comments on `sheet` in row-major order. */
export function listComments(state: State, sheet = state.data.sheetIndex): CommentEntry[] {
  const out: CommentEntry[] = [];
  for (const [key, fmt] of state.format.formats) {
    if (typeof fmt.comment !== 'string' || fmt.comment.length === 0) continue;
    const addr = parseAddrKey(key);
    if (!addr || addr.sheet !== sheet) continue;
    const entry: CommentEntry = { addr, text: fmt.comment };
    if (fmt.commentAuthor) entry.author = fmt.commentAuthor;
    out.push(entry);
  }
  return out.sort((a, b) => a.addr.row - b.addr.row || a.addr.col - b.addr.col);
}

/** Set or replace the comment on a cell. Empty string clears the comment.
 *  When `wb` is provided and the engine supports comments, the change is
 *  mirrored to the engine so it survives a save/load round-trip. */
export function setComment(
  store: SpreadsheetStore,
  addr: Addr,
  text: string,
  wb?: WorkbookHandle,
  authorization?: ThreadedMutationAuthorization,
): boolean {
  if (!isCellWritable(store.getState(), addr)) {
    warnProtected(addr);
    return false;
  }
  const controller = interactionControllerFor(store);
  if (controller && !controller.canExecute(commentIntentFor(addr, authorization)).allowed) {
    return false;
  }
  if (
    text.length > 0 &&
    wb?.capabilities.threadedComments &&
    threadedCommentsAt(store, addr, wb).length > 0
  ) {
    return false;
  }
  const author =
    commentAuthorAt(store.getState(), addr) ?? (text.length > 0 ? DEFAULT_COMMENT_AUTHOR : '');
  if (wb?.capabilities.comments) {
    if (wb.setCommentEntry(addr.sheet, addr.row, addr.col, author, text) === false) return false;
  }
  if (text.length === 0) {
    mutators.setCellFormat(store, addr, { comment: undefined, commentAuthor: undefined });
  } else {
    mutators.setCellFormat(store, addr, { comment: text, commentAuthor: author });
  }
  return true;
}

/** Drop the comment from a cell. No-op when there isn't one. */
export function clearComment(
  store: SpreadsheetStore,
  addr: Addr,
  wb?: WorkbookHandle,
  authorization?: ThreadedMutationAuthorization,
): void {
  if (!isCellWritable(store.getState(), addr)) {
    warnProtected(addr);
    return;
  }
  if (commentAt(store.getState(), addr) === null) return;
  setComment(store, addr, '', wb, authorization);
}

const cloneAddr = (addr: Addr): Addr => ({ sheet: addr.sheet, row: addr.row, col: addr.col });

const captureCommentSnapshot = (state: State, addrs: readonly Addr[]): CommentSnapshot =>
  addrs.map((addr) => ({
    addr: cloneAddr(addr),
    text: commentAt(state, addr),
    author: commentAuthorAt(state, addr),
  }));

const sameCommentSnapshot = (a: CommentSnapshot, b: CommentSnapshot): boolean =>
  a.length === b.length &&
  a.every((entry, index) => {
    const other = b[index];
    return (
      !!other &&
      entry.addr.sheet === other.addr.sheet &&
      entry.addr.row === other.addr.row &&
      entry.addr.col === other.addr.col &&
      entry.text === other.text &&
      entry.author === other.author
    );
  });

const uniqueAddrs = (addrs: readonly Addr[]): Addr[] => {
  const seen = new Set<string>();
  const unique: Addr[] = [];
  for (const addr of addrs) {
    const key = addrKey(addr);
    if (seen.has(key)) continue;
    seen.add(key);
    unique.push(cloneAddr(addr));
  }
  return unique;
};

const capturePhysicalCommentSnapshot = (
  wb: WorkbookHandle | undefined,
  addrs: readonly Addr[],
): PhysicalCommentSnapshot => {
  if (!wb?.capabilities.comments || typeof wb.getComment !== 'function') return [];
  return addrs.map((addr) => ({
    addr: cloneAddr(addr),
    comment: wb.getComment(addr.sheet, addr.row, addr.col),
  }));
};

const applyCommentStoreSnapshot = (store: SpreadsheetStore, snapshot: CommentSnapshot): void => {
  store.setState((s) => {
    const formats = new Map(s.format.formats);
    for (const entry of snapshot) {
      const key = addrKey(entry.addr);
      const current = formats.get(key);
      if (entry.text === null) {
        if (!current) continue;
        const { comment: _comment, commentAuthor: _author, ...rest } = current;
        if (Object.keys(rest).length === 0) formats.delete(key);
        else formats.set(key, rest);
        continue;
      }
      const { comment: _comment, commentAuthor: _author, ...rest } = current ?? {};
      formats.set(key, {
        ...rest,
        comment: entry.text,
        ...(entry.author !== null ? { commentAuthor: entry.author } : {}),
      });
    }
    return { ...s, format: { ...s.format, formats } };
  });
};

const applyPhysicalCommentSnapshot = (
  wb: WorkbookHandle | undefined,
  snapshot: PhysicalCommentSnapshot,
): void => {
  if (!wb?.capabilities.comments) return;
  for (const entry of snapshot) {
    const comment = entry.comment;
    if (
      wb.setCommentEntry(
        entry.addr.sheet,
        entry.addr.row,
        entry.addr.col,
        comment?.author ?? '',
        comment?.text ?? '',
      ) === false
    ) {
      throw new Error('comment engine write failed');
    }
  }
};

const applyCommentSnapshot = (
  store: SpreadsheetStore,
  wb: WorkbookHandle | undefined,
  snapshot: CommentSnapshot,
  authorization: ThreadedMutationAuthorization = {},
): void => {
  const tracked = uniqueAddrs(snapshot.map((entry) => entry.addr));
  const controller = interactionControllerFor(store);
  if (
    controller &&
    tracked.some((addr) => !controller.canExecute(commentIntentFor(addr, authorization)).allowed)
  ) {
    throw new Error('comment history replay blocked by protection or policy');
  }
  const beforeStore = captureCommentSnapshot(store.getState(), tracked);
  const beforePhysical = capturePhysicalCommentSnapshot(wb, tracked);
  try {
    applyCommentStoreSnapshot(store, snapshot);
    if (wb?.capabilities.comments) {
      const physical = snapshot.map((entry) => ({
        addr: entry.addr,
        comment:
          entry.text === null
            ? null
            : { author: entry.author ?? DEFAULT_COMMENT_AUTHOR, text: entry.text },
      }));
      applyPhysicalCommentSnapshot(wb, physical);
    }
  } catch (error) {
    try {
      applyCommentStoreSnapshot(store, beforeStore);
    } catch {
      // Preserve the original engine error when a state observer itself fails.
    }
    try {
      applyPhysicalCommentSnapshot(wb, beforePhysical);
    } catch {
      // The engine may continue rejecting writes; the original error is primary.
    }
    throw error;
  }
};

/** Clear comments for a sparse address list in one store publication and one
 * physical write per target. A failed engine write restores every target. */
export function clearComments(
  store: SpreadsheetStore,
  addrs: readonly Addr[],
  wb?: WorkbookHandle,
  authorization?: ThreadedMutationAuthorization,
): void {
  const tracked = uniqueAddrs(addrs);
  const state = store.getState();
  const writable = tracked.filter(
    (addr) => isCellWritable(state, addr) && commentAt(state, addr) !== null,
  );
  if (writable.length === 0) return;
  applyCommentSnapshot(
    store,
    wb,
    writable.map((addr) => ({ addr, text: null, author: null })),
    authorization,
  );
}

export function recordCommentChange<T>(
  history: History | null,
  store: SpreadsheetStore,
  wb: WorkbookHandle | undefined,
  addrs: readonly Addr[],
  mutate: () => T,
  authorization: ThreadedMutationAuthorization = {},
): T {
  if (!history || history.isReplaying()) return mutate();
  const tracked = uniqueAddrs(addrs);
  const before = captureCommentSnapshot(store.getState(), tracked);
  const result = mutate();
  const after = captureCommentSnapshot(store.getState(), tracked);
  const changedBefore: CommentSnapshot = [];
  const changedAfter: CommentSnapshot = [];
  for (let index = 0; index < before.length; index += 1) {
    const beforeEntry = before[index];
    const afterEntry = after[index];
    if (!beforeEntry || !afterEntry || sameCommentSnapshot([beforeEntry], [afterEntry])) continue;
    changedBefore.push(beforeEntry);
    changedAfter.push(afterEntry);
  }
  if (changedBefore.length > 0) {
    const effects: readonly OperationEffect[] = [
      { kind: 'cells', cells: changedBefore.map((entry) => entry.addr) },
    ];
    const intent = {
      operation: 'comment' as const,
      origin: authorization.origin ?? 'instanceApi',
      commandId: authorization.commandId ?? 'commentChange',
      effects,
    };
    history.push({
      undo: () =>
        applyCommentSnapshot(store, wb, changedBefore, { ...authorization, origin: 'undo' }),
      redo: () =>
        applyCommentSnapshot(store, wb, changedAfter, { ...authorization, origin: 'redo' }),
      intent: { ...intent, origin: 'redo' },
      inverseIntent: { ...intent, origin: 'undo' },
    });
  }
  return result;
}
