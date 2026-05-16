// Shared "Comment" ribbon split-button action — clears either the comment on the
// active cell ("delete-active") or every comment on the sheet ("delete-all").
// Host wrappers don't need to know the bookkeeping: we resolve the target set
// from the store, no-op when empty, and wrap the mutations in a single
// format-change entry so undo collapses them.

import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { SpreadsheetStore } from '../store/store.js';
import { clearComment, commentAt, listComments } from './comment.js';
import { type History, recordFormatChange } from './history.js';

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
  if (targets.length === 0) return;
  recordFormatChange(history, store, () => {
    for (const entry of targets) clearComment(store, entry.addr, workbook);
  });
};
