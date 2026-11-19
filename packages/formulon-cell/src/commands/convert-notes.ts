import { hydrateCommentsAndHyperlinksFromEngine } from '../engine/format-sync.js';
import type { Addr } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import {
  addThreadedComment,
  clearComment,
  recordCommentChange,
  recordThreadedChange,
} from './comment.js';
import { type CommentActionContext, commentWrite, threadWrite } from './comment-actions.js';
import type { History } from './history.js';

/** Whether any sheet of `wb` holds a note Convert to Comments would turn into a thread. */
export const workbookHasNotes = (wb: WorkbookHandle): boolean => {
  if (!wb.capabilities.threadedComments || !wb.capabilities.commentsEnumerable) return false;
  for (let sheet = 0; sheet < wb.sheetCount; sheet += 1) {
    if (wb.getComments(sheet).some((note) => note.text.length > 0)) return true;
  }
  return false;
};

/**
 * Convert to Comments: every note of the workbook becomes a thread whose first comment is the
 * note's text, under the note's author, as one undo step. A note the policy or a locked cell keeps
 * stays a note; when a removed note's thread cannot be added, nothing is converted. Returns how
 * many notes became threads.
 */
export function convertNotesToThreads(
  ctx: CommentActionContext & { readonly history: History; readonly workbook: WorkbookHandle },
): number {
  const { store, workbook: wb, history } = ctx;
  if (!workbookHasNotes(wb)) return 0;
  const authorization = { origin: ctx.origin, commandId: ctx.commandId };
  let converted = 0;
  let lost = false;
  history.transact(
    () => {
      for (let sheet = 0; sheet < wb.sheetCount; sheet += 1) {
        const notes = wb.getComments(sheet).filter((note) => note.text.length > 0);
        if (notes.length === 0) continue;
        // The store mirrors the active sheet's notes only; the note commands read them there.
        hydrateCommentsAndHyperlinksFromEngine(wb, store, sheet);
        const cells: Addr[] = notes.map((note) => ({ sheet, row: note.row, col: note.col }));
        commentWrite(ctx, cells, undefined, () =>
          recordCommentChange(
            history,
            store,
            wb,
            cells,
            () => {
              for (const cell of cells) clearComment(store, cell, wb, authorization);
            },
            authorization,
          ),
        );
        const cleared = notes.filter((note) => !wb.getComment(sheet, note.row, note.col));
        const before = converted;
        threadWrite(ctx, cells, undefined, () =>
          recordThreadedChange(
            history,
            store,
            wb,
            sheet,
            () => {
              for (const note of cleared) {
                const cell = { sheet, row: note.row, col: note.col };
                const added = addThreadedComment(store, cell, note.text, wb, {
                  author: note.author ?? '',
                  authorization,
                });
                if (added) converted += 1;
              }
            },
            authorization,
          ),
        );
        if (converted - before < cleared.length) {
          lost = true;
          break;
        }
      }
    },
    { keep: () => !lost },
  );
  return lost ? 0 : converted;
}
