import { addrKey } from '../engine/address.js';
import type { Addr } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { mutators, type SpreadsheetStore, type State } from '../store/store.js';

export interface CommentEntry {
  addr: Addr;
  text: string;
}

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
  mutators.setCellFormat(store, addr, { comment: undefined });
  if (wb?.capabilities.comments) {
    wb.setCommentEntry(addr.sheet, addr.row, addr.col, '', '');
  }
}
