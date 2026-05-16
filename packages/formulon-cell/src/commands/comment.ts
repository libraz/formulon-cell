import { addrKey } from '../engine/address.js';
import type { Addr } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { mutators, type SpreadsheetStore, type State } from '../store/store.js';
import type { History } from './history.js';
import { isCellWritable, warnProtected } from './protection.js';

export interface CommentEntry {
  addr: Addr;
  text: string;
}

type CommentSnapshot = Array<{ addr: Addr; text: string | null }>;

/** Read the comment text on a cell, or null when unset. */
export function commentAt(state: State, addr: Addr): string | null {
  const fmt = state.format.formats.get(addrKey(addr));
  const c = fmt?.comment;
  return typeof c === 'string' && c.length > 0 ? c : null;
}

/** List non-empty comments on `sheet` in row-major order. */
export function listComments(state: State, sheet = state.data.sheetIndex): CommentEntry[] {
  const out: CommentEntry[] = [];
  for (const [key, fmt] of state.format.formats) {
    if (typeof fmt.comment !== 'string' || fmt.comment.length === 0) continue;
    const parts = key.split(':').map((n) => Number(n));
    const s = parts[0] ?? -1;
    const row = parts[1] ?? -1;
    const col = parts[2] ?? -1;
    if (s !== sheet) continue;
    out.push({ addr: { sheet: s, row, col }, text: fmt.comment });
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
  if (text.length === 0) {
    mutators.setCellFormat(store, addr, { comment: undefined });
  } else {
    mutators.setCellFormat(store, addr, { comment: text });
  }
  if (wb?.capabilities.comments) {
    wb.setCommentEntry(addr.sheet, addr.row, addr.col, '', text);
  }
}

/** Drop the comment from a cell. No-op when there isn't one. */
export function clearComment(store: SpreadsheetStore, addr: Addr, wb?: WorkbookHandle): void {
  if (!isCellWritable(store.getState(), addr)) {
    warnProtected(addr);
    return;
  }
  if (commentAt(store.getState(), addr) === null) return;
  mutators.setCellFormat(store, addr, { comment: undefined });
  if (wb?.capabilities.comments) {
    wb.setCommentEntry(addr.sheet, addr.row, addr.col, '', '');
  }
}

const cloneAddr = (addr: Addr): Addr => ({ sheet: addr.sheet, row: addr.row, col: addr.col });

const captureCommentSnapshot = (state: State, addrs: readonly Addr[]): CommentSnapshot =>
  addrs.map((addr) => ({ addr: cloneAddr(addr), text: commentAt(state, addr) }));

const sameCommentSnapshot = (a: CommentSnapshot, b: CommentSnapshot): boolean =>
  a.length === b.length &&
  a.every((entry, index) => {
    const other = b[index];
    return (
      !!other &&
      entry.addr.sheet === other.addr.sheet &&
      entry.addr.row === other.addr.row &&
      entry.addr.col === other.addr.col &&
      entry.text === other.text
    );
  });

const applyCommentSnapshot = (
  store: SpreadsheetStore,
  wb: WorkbookHandle | undefined,
  snapshot: CommentSnapshot,
): void => {
  for (const entry of snapshot) {
    mutators.setCellFormat(store, entry.addr, { comment: entry.text ?? undefined });
    if (wb?.capabilities.comments) {
      wb.setCommentEntry(entry.addr.sheet, entry.addr.row, entry.addr.col, '', entry.text ?? '');
    }
  }
};

export function recordCommentChange<T>(
  history: History | null,
  store: SpreadsheetStore,
  wb: WorkbookHandle | undefined,
  addrs: readonly Addr[],
  mutate: () => T,
): T {
  if (!history || history.isReplaying()) return mutate();
  const tracked = addrs.map(cloneAddr);
  const before = captureCommentSnapshot(store.getState(), tracked);
  const result = mutate();
  const after = captureCommentSnapshot(store.getState(), tracked);
  if (!sameCommentSnapshot(before, after)) {
    history.push({
      undo: () => applyCommentSnapshot(store, wb, before),
      redo: () => applyCommentSnapshot(store, wb, after),
    });
  }
  return result;
}
