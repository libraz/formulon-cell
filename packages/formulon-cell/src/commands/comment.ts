import type { Addr } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { addrKey } from '../engine/workbook-handle.js';
import { mutators, type SpreadsheetStore, type State } from '../store/store.js';

/** Read the comment text on a cell, or null when unset. */
export function commentAt(state: State, addr: Addr): string | null {
  const fmt = state.format.formats.get(addrKey(addr));
  const c = fmt?.comment;
  return typeof c === 'string' && c.length > 0 ? c : null;
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
