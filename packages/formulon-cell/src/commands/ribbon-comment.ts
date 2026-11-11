// Shared "Comment" ribbon split-button action — clears either the comment on the
// active cell ("delete-active") or every comment on the sheet ("delete-all").
// Host wrappers don't need to know the bookkeeping: we resolve the target set
// from the store and keep note and threaded-comment undo entries independent.

import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { mutators, type SpreadsheetStore } from '../store/store.js';
import {
  clearComment,
  commentAt,
  listComments,
  listThreadedComments,
  noteVisibleAt,
  recordCommentChange,
  recordThreadedChange,
  removeThreadedComment,
  setNotesVisible,
  threadedCommentsAt,
} from './comment.js';
import { commentWrite, threadWrite } from './comment-actions.js';
import type { History } from './history.js';
import type { InteractionOrigin } from './interaction-policy.js';

export type RibbonCommentAction = 'delete-active' | 'delete-all';

export interface ExecuteRibbonCommentActionDeps {
  store: SpreadsheetStore;
  workbook: WorkbookHandle;
  history: History;
  action: RibbonCommentAction;
  /** The entry point deleting the comments. */
  origin: InteractionOrigin;
  commandId: string;
}

export const executeRibbonCommentAction = (deps: ExecuteRibbonCommentActionDeps): void => {
  const { store, workbook, history, action, origin, commandId } = deps;
  const authorization = { origin, commandId };
  const state = store.getState();
  const targets =
    action === 'delete-active'
      ? commentAt(state, state.selection.active) === null
        ? []
        : [{ addr: state.selection.active }]
      : listComments(state);
  const threaded = workbook.capabilities.threadedComments
    ? action === 'delete-active'
      ? threadedCommentsAt(store, state.selection.active, workbook)
      : listThreadedComments(store, state.data.sheetIndex, workbook)
    : [];
  if (targets.length === 0 && threaded.length === 0) return;
  // Notes and threads are separate undo steps, each one authorized comment write.
  const ctx = { store, history, ...authorization };
  if (targets.length > 0) {
    const cells = targets.map((entry) => entry.addr);
    commentWrite(ctx, cells, undefined, () =>
      recordCommentChange(
        history,
        store,
        workbook,
        cells,
        () => {
          for (const entry of targets) clearComment(store, entry.addr, workbook, authorization);
        },
        authorization,
      ),
    );
  }
  if (threaded.length > 0 && workbook.capabilities.threadedComments) {
    const sheet = state.data.sheetIndex;
    const cells = threaded.map((entry) => ({ sheet, row: entry.cell.row, col: entry.cell.col }));
    threadWrite(ctx, cells, undefined, () =>
      recordThreadedChange(
        history,
        store,
        workbook,
        sheet,
        () => {
          for (const entry of [...threaded].sort(
            (a, b) => Number(a.parentId === null) - Number(b.parentId === null),
          )) {
            removeThreadedComment(store, sheet, entry.id, workbook, authorization);
          }
        },
        authorization,
      ),
    );
  }
};

/** `toggle-active` is Show/Hide Note on the active cell, which the workbook keeps;
 *  `toggle-all` is Show All Notes, a view preference that draws every note and leaves
 *  each note's own shown state alone. */
export type NoteVisibilityAction = 'toggle-active' | 'toggle-all';

export interface ExecuteNoteVisibilityActionDeps {
  store: SpreadsheetStore;
  workbook: WorkbookHandle;
  action: NoteVisibilityAction;
}

export const executeNoteVisibilityAction = (deps: ExecuteNoteVisibilityActionDeps): boolean => {
  const { store, workbook, action } = deps;
  const state = store.getState();
  if (action === 'toggle-all') {
    mutators.setShowAllNotes(store, !state.ui.showAllNotes);
    return true;
  }
  const active = state.selection.active;
  if (commentAt(state, active) === null) return false;
  return setNotesVisible(store, [active], !noteVisibleAt(state, active), workbook);
};
