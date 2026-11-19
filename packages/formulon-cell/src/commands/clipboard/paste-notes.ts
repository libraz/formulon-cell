import { addrKey, parseAddrKey } from '../../engine/address.js';
import type { Range } from '../../engine/types.js';
import type { WorkbookHandle } from '../../engine/workbook-handle.js';
import type { CellFormat, SpreadsheetStore } from '../../store/store.js';
import { MAX_PASTE_CELLS } from './paste-destination.js';

export const commentText = (format: CellFormat | undefined): string | null =>
  typeof format?.comment === 'string' && format.comment.length > 0 ? format.comment : null;

export const commentAuthor = (format: CellFormat | undefined): string | null =>
  commentText(format) && typeof format?.commentAuthor === 'string' ? format.commentAuthor : null;

/** Bring engine-only notes into the store before a format/comment history
 * snapshot is captured. The engine is authoritative for persisted comments,
 * while the store remains the source used by the history helpers. */
export const hydrateEngineComments = (
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  ranges: readonly Range[],
): void => {
  if (!wb.capabilities.commentsEnumerable || ranges.length === 0) return;
  const sheets = new Set(ranges.map((range) => range.sheet));
  const comments: Array<{
    sheet: number;
    row: number;
    col: number;
    author: string | null;
    text: string;
    visible?: boolean;
  }> = [];
  for (const sheet of sheets) {
    for (const comment of wb.getComments(sheet)) {
      if (!comment.text) continue;
      if (
        ranges.some(
          (range) =>
            range.sheet === sheet &&
            comment.row >= range.r0 &&
            comment.row <= range.r1 &&
            comment.col >= range.c0 &&
            comment.col <= range.c1,
        )
      ) {
        comments.push({ sheet, ...comment });
      }
    }
  }
  if (comments.length === 0) return;
  store.setState((s) => {
    const formats = new Map(s.format.formats);
    for (const comment of comments) {
      const key = addrKey({ sheet: comment.sheet, row: comment.row, col: comment.col });
      formats.set(key, {
        ...formats.get(key),
        comment: comment.text,
        commentAuthor: comment.author || undefined,
        commentVisible: comment.visible || undefined,
      });
    }
    return { ...s, format: { ...s.format, formats } };
  });
};

export const formatWithoutComment = (format: CellFormat | undefined): CellFormat | null => {
  if (!format) return null;
  const next = { ...format };
  delete next.comment;
  delete next.commentAuthor;
  delete next.commentVisible;
  return next;
};

export const formatCommentsOnly = (format: CellFormat | undefined): CellFormat | null => {
  const text = commentText(format);
  if (!text) return null;
  const author = commentAuthor(format);
  return {
    comment: text,
    ...(author ? { commentAuthor: author } : {}),
    ...(format?.commentVisible === true ? { commentVisible: true } : {}),
  };
};

export const clearSourceFormats = (store: SpreadsheetStore, source: Range): void => {
  store.setState((s) => {
    const formats = new Map(s.format.formats);
    const area = (source.r1 - source.r0 + 1) * (source.c1 - source.c0 + 1);
    const clear = (key: string): void => {
      const comments = formatCommentsOnly(formats.get(key));
      if (comments) formats.set(key, comments);
      else formats.delete(key);
    };
    if (area > MAX_PASTE_CELLS) {
      for (const key of formats.keys()) {
        const addr = parseAddrKey(key);
        if (!addr) continue;
        const { sheet, row, col } = addr;
        if (
          sheet === source.sheet &&
          row >= source.r0 &&
          row <= source.r1 &&
          col >= source.c0 &&
          col <= source.c1
        ) {
          clear(key);
        }
      }
    } else {
      for (let row = source.r0; row <= source.r1; row += 1) {
        for (let col = source.c0; col <= source.c1; col += 1) {
          clear(addrKey({ sheet: source.sheet, row, col }));
        }
      }
    }
    return { ...s, format: { ...s.format, formats } };
  });
};
