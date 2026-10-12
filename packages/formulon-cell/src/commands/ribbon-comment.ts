// Shared "Comment" ribbon split-button action — clears either the comment on the
// active cell ("delete-active") or every comment on the sheet ("delete-all").
// Host wrappers don't need to know the bookkeeping: we resolve the target set
// from the store and keep note and threaded-comment undo entries independent.

import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { SpreadsheetStore } from '../store/store.js';
import {
  clearComment,
  commentAt,
  listComments,
  listThreadedComments,
  recordCommentChange,
  recordThreadedChange,
  removeThreadedComment,
  threadedCommentsAt,
} from './comment.js';
import type { History } from './history.js';

export type RibbonCommentAction = 'delete-active' | 'delete-all';

export interface ExecuteRibbonCommentActionDeps {
  store: SpreadsheetStore;
  workbook: WorkbookHandle;
  history: History;
  action: RibbonCommentAction;
}

export const executeRibbonCommentAction = (deps: ExecuteRibbonCommentActionDeps): void => {
  const { store, workbook, history, action } = deps;
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
  if (targets.length > 0) {
    recordCommentChange(
      history,
      store,
      workbook,
      targets.map((entry) => entry.addr),
      () => {
        for (const entry of targets)
          clearComment(store, entry.addr, workbook, {
            origin: 'ribbon',
            commandId: 'deleteComment',
          });
      },
      { origin: 'ribbon', commandId: 'deleteComment' },
    );
  }
  if (threaded.length > 0 && workbook.capabilities.threadedComments) {
    recordThreadedChange(
      history,
      store,
      workbook,
      state.data.sheetIndex,
      () => {
        for (const entry of [...threaded].sort(
          (a, b) => Number(a.parentId === '') - Number(b.parentId === ''),
        )) {
          removeThreadedComment(store, state.data.sheetIndex, entry.id, workbook, {
            origin: 'ribbon',
            commandId: 'deleteComment',
          });
        }
      },
      {
        origin: 'ribbon',
        commandId: 'deleteComment',
      },
    );
  }
};
