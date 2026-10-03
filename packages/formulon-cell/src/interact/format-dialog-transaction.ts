// Commit step of the Format Cells dialog: applies the resolved format action
// and the optional merge/unmerge as one undoable unit, and registers the
// format-only F4 repeat. A merge failure aborts the whole composite so a
// half-applied dialog never reaches history.

import { recordDialogFormatChange } from '../commands/dialog-format-history.js';
import {
  applySelectionFormatAction,
  planSelectionFormat,
  type SelectionFormatAction,
  type SelectionFormatPlan,
} from '../commands/format.js';
import { History, recordMergesChangeWithEngine } from '../commands/history.js';
import { applyMerge, applyUnmerge } from '../commands/merge.js';
import type { Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { mutators, type SpreadsheetStore, type State } from '../store/store.js';

export interface FormatDialogTransaction {
  readonly store: SpreadsheetStore;
  readonly history: History | null;
  /** Resolves the workbook afresh when F4 repeats the formatting later. */
  getWb(): WorkbookHandle | null;
  /** Store state and workbook captured when OK was pressed. */
  readonly state: State;
  readonly liveWb: WorkbookHandle | null;
  readonly plan: SelectionFormatPlan;
  readonly action: SelectionFormatAction;
  readonly merge: { action: 'merge' | 'unmerge' | null; range: Range };
}

/** Returns true when every step applied; false when one was rejected and the
 *  composite rolled back. Rethrows (after rollback) if a step throws. */
export function runFormatDialogTransaction(tx: FormatDialogTransaction): boolean {
  const { store, history, getWb, state, liveWb, plan, action } = tx;
  const range = state.selection.range;
  const mergeAction = tx.merge.action;
  const mergeRange = tx.merge.range;
  const hasFormatAction = Object.keys(action.patch).length > 0 || action.border !== undefined;

  // F4 repeats only formatting, never cell-identity metadata. It resolves a
  // fresh selection, workbook, policy, and plan at invocation time.
  const {
    hyperlink: _hyperlink,
    hyperlinkDisplay: _hyperlinkDisplay,
    hyperlinkTooltip: _hyperlinkTooltip,
    comment: _comment,
    commentAuthor: _commentAuthor,
    validation: _validation,
    ...repeatablePatch
  } = action.patch;
  const repeatAction: SelectionFormatAction = {
    patch: structuredClone(repeatablePatch),
    ...(action.border ? { border: structuredClone(action.border) } : {}),
  };
  const hasRepeatAction = Object.keys(repeatablePatch).length > 0 || action.border !== undefined;
  let repeatFormatting: (() => void) | undefined;
  if (hasRepeatAction) {
    repeatFormatting = (): void => {
      const current = store.getState();
      const currentPlan = planSelectionFormat(current);
      if (!currentPlan) return;
      const currentWb = getWb();
      recordDialogFormatChange({
        history,
        store,
        workbook: currentWb,
        sheet: current.selection.range.sheet,
        targets: currentPlan.cells,
        pendingBefore: current.ui.pendingFormat,
        mutate: () =>
          applySelectionFormatAction(current, store, repeatAction, {
            allowPending: false,
            origin: 'instanceApi',
            commandId: 'formatCells',
          }),
        repeat: repeatFormatting,
      });
    };
  }

  // A merge is a legacy multi-child history action. Use an ephemeral history
  // when the caller did not provide one so a later merge failure can abort
  // the already-applied format child atomically.
  const actionHistory = mergeAction && !history ? new History() : history;
  const transaction = mergeAction && actionHistory ? actionHistory.begin() : undefined;
  let transactionOpen = transaction !== undefined;
  const abortTransaction = (): void => {
    if (!transaction || !actionHistory || !transactionOpen) return;
    transactionOpen = false;
    try {
      actionHistory.abort(transaction);
    } finally {
      // Scoped material replay intentionally preserves a later pending
      // format. A failed composite dialog action instead restores the exact
      // pre-action pending snapshot.
      mutators.setPendingFormat(store, state.ui.pendingFormat);
    }
  };
  try {
    if (hasFormatAction) {
      const wrote = recordDialogFormatChange({
        history: actionHistory,
        store,
        workbook: liveWb,
        sheet: range.sheet,
        targets: plan.cells,
        pendingBefore: state.ui.pendingFormat,
        mutate: () =>
          applySelectionFormatAction(store.getState(), store, action, {
            allowPending: false,
            origin: 'instanceApi',
            commandId: 'formatCells',
          }),
        repeat: repeatFormatting,
        registerRepeat: !transaction,
      });
      if (!wrote) {
        abortTransaction();
        return false;
      }
    }
    if (mergeAction === 'merge') {
      if (mergeRange.r0 !== mergeRange.r1 || mergeRange.c0 !== mergeRange.c1) {
        const merged = liveWb
          ? applyMerge(store, liveWb, actionHistory, mergeRange)
          : (() => {
              recordMergesChangeWithEngine(actionHistory, store, null, mergeRange.sheet, () => {
                mutators.mergeRange(store, mergeRange);
              });
              return true;
            })();
        if (!merged) {
          abortTransaction();
          return false;
        }
      }
    } else if (mergeAction === 'unmerge') {
      if (!applyUnmerge(store, liveWb, actionHistory, range)) {
        abortTransaction();
        return false;
      }
    }
    if (transaction && actionHistory) {
      actionHistory.end(transaction);
      transactionOpen = false;
      if (hasFormatAction && repeatFormatting) actionHistory.setRepeat(repeatFormatting);
    }
    return true;
  } catch (error) {
    try {
      abortTransaction();
    } catch (abortError) {
      throw new AggregateError(
        [error, abortError],
        'Format dialog transaction failed and its rollback failed',
        { cause: error },
      );
    }
    throw error;
  }
}
