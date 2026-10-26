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
  noteVisibleAt,
  recordCommentChange,
  recordThreadedChange,
  removeThreadedComment,
  setNotesVisible,
  threadedCommentsAt,
} from './comment.js';
import type { History } from './history.js';
import { deferForAllowedRangeUnlock } from './protection.js';

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
  const footprint = [
    ...targets.map((entry) => entry.addr),
    ...threaded.map((entry) => ({ sheet: state.data.sheetIndex, row: entry.row, col: entry.col })),
  ].map((a) => ({ sheet: a.sheet, r0: a.row, c0: a.col, r1: a.row, c1: a.col }));
  if (deferForAllowedRangeUnlock(store, footprint, () => executeRibbonCommentAction(deps))) return;
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

/** `toggle-active` is Show/Hide Note on the active cell; `toggle-all` is Show
 *  All Notes, which shows every note unless all of them are already shown. */
export type NoteVisibilityAction = 'toggle-active' | 'toggle-all';

export interface ExecuteNoteVisibilityActionDeps {
  store: SpreadsheetStore;
  workbook: WorkbookHandle;
  action: NoteVisibilityAction;
}

export const executeNoteVisibilityAction = (deps: ExecuteNoteVisibilityActionDeps): boolean => {
  const { store, workbook, action } = deps;
  const state = store.getState();
  const active = state.selection.active;
  const addrs =
    action === 'toggle-active'
      ? commentAt(state, active) === null
        ? []
        : [active]
      : listComments(state).map((entry) => entry.addr);
  if (addrs.length === 0) return false;
  const visible =
    action === 'toggle-active'
      ? !noteVisibleAt(state, active)
      : !addrs.every((addr) => noteVisibleAt(state, addr));
  return setNotesVisible(store, addrs, visible, workbook);
};
