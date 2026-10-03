import { addrKey, parseAddrKey } from '../engine/address.js';
import type { Addr } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { mutators, type SpreadsheetStore, type State } from '../store/store.js';
import type { History } from './history.js';
import { isCellWritable, warnProtected } from './protection.js';

export interface CommentEntry {
  addr: Addr;
  text: string;
  author?: string;
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
): void {
  if (!isCellWritable(store.getState(), addr)) {
    warnProtected(addr);
    return;
  }
  const author = commentAuthorAt(store.getState(), addr) ?? '';
  if (text.length === 0) {
    mutators.setCellFormat(store, addr, { comment: undefined, commentAuthor: undefined });
  } else {
    mutators.setCellFormat(store, addr, { comment: text });
  }
  if (wb?.capabilities.comments) {
    wb.setCommentEntry(addr.sheet, addr.row, addr.col, author, text);
  }
}

/** Drop the comment from a cell. No-op when there isn't one. */
export function clearComment(store: SpreadsheetStore, addr: Addr, wb?: WorkbookHandle): void {
  if (!isCellWritable(store.getState(), addr)) {
    warnProtected(addr);
    return;
  }
  if (commentAt(store.getState(), addr) === null) return;
  const author = commentAuthorAt(store.getState(), addr) ?? '';
  mutators.setCellFormat(store, addr, { comment: undefined, commentAuthor: undefined });
  if (wb?.capabilities.comments) {
    wb.setCommentEntry(addr.sheet, addr.row, addr.col, author, '');
  }
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
      !wb.setCommentEntry(
        entry.addr.sheet,
        entry.addr.row,
        entry.addr.col,
        comment?.author ?? '',
        comment?.text ?? '',
      )
    ) {
      throw new Error('comment engine write failed');
    }
  }
};

const applyCommentSnapshot = (
  store: SpreadsheetStore,
  wb: WorkbookHandle | undefined,
  snapshot: CommentSnapshot,
): void => {
  const tracked = uniqueAddrs(snapshot.map((entry) => entry.addr));
  const beforeStore = captureCommentSnapshot(store.getState(), tracked);
  const beforePhysical = capturePhysicalCommentSnapshot(wb, tracked);
  try {
    applyCommentStoreSnapshot(store, snapshot);
    if (wb?.capabilities.comments) {
      const physical = snapshot.map((entry) => ({
        addr: entry.addr,
        comment: entry.text === null ? null : { author: entry.author ?? '', text: entry.text },
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
  );
}

export function recordCommentChange<T>(
  history: History | null,
  store: SpreadsheetStore,
  wb: WorkbookHandle | undefined,
  addrs: readonly Addr[],
  mutate: () => T,
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
    history.push({
      undo: () => applyCommentSnapshot(store, wb, changedBefore),
      redo: () => applyCommentSnapshot(store, wb, changedAfter),
    });
  }
  return result;
}
