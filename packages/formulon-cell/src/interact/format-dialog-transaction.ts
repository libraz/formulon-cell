// Commit step of the Format Cells dialog: applies the resolved format action
// and the optional merge/unmerge as one undoable unit, and registers the
// format-only F4 repeat. A merge failure aborts the whole composite so a
// half-applied dialog never reaches history.

import { type AuthorizedIntentFactory, authorizedWrite } from '../commands/authorized-write.js';
import { recordDialogFormatChange } from '../commands/dialog-format-history.js';
import {
  applySelectionFormatAction,
  type SelectionFormatAction,
  selectionFormatActionOperations,
} from '../commands/format.js';
import { History } from '../commands/history.js';
import type { InteractionOperation } from '../commands/interaction-policy.js';
import { applyMerge, applyUnmerge } from '../commands/merge.js';
import {
  planSelectionFormat,
  type SelectionFormatPlan,
} from '../commands/selection-format-plan.js';
import { recordMergesChangeWithEngine } from '../commands/slice-history.js';
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

/** Origin and command id the dialog commit, its F4 repeat and their replay are decided under. */
const DIALOG_ORIGIN = 'instanceApi';
const DIALOG_COMMAND_ID = 'formatCells';

/** A step refused after the write started: the composite is rolled back and reported as false. */
class StepRejected extends Error {}

const dialogIntents = (
  operations: readonly InteractionOperation[],
  ranges: readonly Range[],
  mergeRange: Range | null,
): AuthorizedIntentFactory => {
  const effects = ranges.map((range) => ({ kind: 'range' as const, range: { ...range } }));
  return (origin) => [
    ...operations.map((operation) => ({
      operation,
      origin,
      commandId: DIALOG_COMMAND_ID,
      effects,
    })),
    ...(mergeRange
      ? [
          {
            operation: 'merge' as const,
            origin,
            commandId: DIALOG_COMMAND_ID,
            effects: [{ kind: 'range' as const, range: { ...mergeRange } }],
          },
        ]
      : []),
  ];
};

/** Returns true when every step applied; false when one was refused and the
 *  composite rolled back. Rethrows (after rollback) if a step throws. */
export function runFormatDialogTransaction(tx: FormatDialogTransaction): boolean {
  const { store, history, getWb, state, liveWb, plan, action } = tx;
  const range = state.selection.range;
  const mergeAction = tx.merge.action;
  const mergeRange = tx.merge.range;
  const hasFormatAction = Object.keys(action.patch).length > 0 || action.border !== undefined;
  // Without a caller history the composite still runs as one transaction, so a
  // later refusal can roll back the steps already applied.
  const actionHistory = history ?? new History();

  // F4 repeats only formatting, never cell-identity metadata. It resolves a
  // fresh selection, workbook, policy, and plan at invocation time.
  const {
    hyperlink: _hyperlink,
    hyperlinkLocation: _hyperlinkLocation,
    hyperlinkSpan: _hyperlinkSpan,
    hyperlinkDisplay: _hyperlinkDisplay,
    hyperlinkTooltip: _hyperlinkTooltip,
    comment: _comment,
    commentAuthor: _commentAuthor,
    commentVisible: _commentVisible,
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
      authorizedWrite({
        store,
        history: actionHistory,
        intents: dialogIntents(
          [...selectionFormatActionOperations(repeatAction)],
          currentPlan.ranges,
          null,
        ),
        origin: DIALOG_ORIGIN,
        write: () =>
          recordDialogFormatChange({
            history: actionHistory,
            store,
            workbook: currentWb,
            sheet: current.selection.range.sheet,
            targets: currentPlan.cells,
            mutate: () =>
              applySelectionFormatAction(current, store, repeatAction, {
                origin: DIALOG_ORIGIN,
                commandId: DIALOG_COMMAND_ID,
              }),
            repeat: repeatFormatting,
          }),
      });
    };
  }

  const write = (): void => {
    if (hasFormatAction) {
      const wrote = recordDialogFormatChange({
        history: actionHistory,
        store,
        workbook: liveWb,
        sheet: range.sheet,
        targets: plan.cells,
        mutate: () =>
          applySelectionFormatAction(store.getState(), store, action, {
            origin: DIALOG_ORIGIN,
            commandId: DIALOG_COMMAND_ID,
          }),
        repeat: repeatFormatting,
        registerRepeat: false,
      });
      if (!wrote) throw new StepRejected();
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
        if (!merged) throw new StepRejected();
      }
    } else if (mergeAction === 'unmerge') {
      if (!applyUnmerge(store, liveWb, actionHistory, range)) throw new StepRejected();
    }
  };

  let result: ReturnType<typeof authorizedWrite<void>>;
  try {
    result = authorizedWrite({
      store,
      history: actionHistory,
      intents: dialogIntents(
        hasFormatAction ? [...selectionFormatActionOperations(action)] : [],
        plan.ranges,
        mergeAction === 'merge' ? mergeRange : mergeAction === 'unmerge' ? range : null,
      ),
      origin: DIALOG_ORIGIN,
      write,
    });
  } catch (error) {
    if (error instanceof StepRejected) return false;
    throw error;
  }
  if (!result.ok) return false;
  if (hasFormatAction) actionHistory.setRepeat(repeatFormatting ?? null);
  return true;
}
