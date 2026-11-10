import {
  type AuthorizedIntentFactory,
  authorizedWrite,
  authorizeWrite,
} from '../commands/authorized-write.js';
import { applyFormatPatch } from '../commands/format.js';
import { History } from '../commands/history.js';
import type { InteractionOrigin } from '../commands/interaction-policy.js';
import {
  applyMerge,
  applyMergeAcross,
  applyUnmerge,
  expandRangeWithMerges,
  mergeAcrossWillLoseData,
  mergeIntents,
  mergeWillLoseData,
} from '../commands/merge.js';
import { deferForAllowedRangeUnlock } from '../commands/protection.js';
import { recordFormatChange } from '../commands/slice-history.js';
import type { Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { Strings } from '../i18n/strings.js';
import type { SpreadsheetStore } from '../store/store.js';
import { confirmMergeLoseData } from './dialogs/merge-confirm.js';

/** Actions exposed by the merge controls in wrappers and ribbon menus. */
export type MergeAction =
  | 'mergeCenter'
  | 'mergeAcross'
  | 'mergeCells'
  | 'unmergeCells'
  | 'centerAcross';

export interface MergeActionContext {
  store: SpreadsheetStore;
  workbook: WorkbookHandle;
  history: History | null;
  strings: Strings;
  /** The entry point the action is authorized under. */
  origin: InteractionOrigin;
  commandId: string;
}

/**
 * Apply one merge-control action as one authorized history step. The helper
 * deliberately performs no UI follow-up (focus, redraw, or projection);
 * callers own those callbacks.
 *
 * The first await is reached only when a merge would discard data. That keeps
 * the no-warning path synchronous for the existing toolbar callers while
 * ensuring a cancelled warning leaves the workbook untouched. A refused action
 * is decided before that prompt.
 */
export async function applyMergeAction(
  context: MergeActionContext,
  action: MergeAction,
): Promise<boolean> {
  const { store, workbook, history, strings, origin, commandId } = context;
  const selection = store.getState().selection.range;
  // A password-guarded range under the selection is unlocked first; the action then repeats.
  if (
    deferForAllowedRangeUnlock(store, [selection], () => void applyMergeAction(context, action))
  ) {
    return false;
  }

  const state = store.getState();
  // A merge that intersects an existing merge acts on the complete merged
  // region. The command layer also normalizes this, but using the same target
  // for the authorization, the warning and the mutation keeps them in step.
  const range = expandRangeWithMerges(state, selection);
  const formatIntents =
    (target: Range): AuthorizedIntentFactory =>
    (intentOrigin) => [
      {
        operation: 'format',
        origin: intentOrigin,
        commandId,
        effects: [{ kind: 'range', range: { ...target } }],
      },
    ];
  const intents: AuthorizedIntentFactory =
    action === 'centerAcross'
      ? formatIntents(selection)
      : action === 'mergeCenter'
        ? (intentOrigin) => [
            ...mergeIntents(range, commandId)(intentOrigin),
            ...formatIntents(range)(intentOrigin),
          ]
        : mergeIntents(range, commandId);
  if (!authorizeWrite(store, intents, origin).allowed) return false;

  if (action === 'mergeAcross' && mergeAcrossWillLoseData(state, selection)) {
    if (!(await confirmMergeLoseData(strings, state, selection, true, workbook))) return false;
  } else if (
    (action === 'mergeCenter' || action === 'mergeCells') &&
    mergeWillLoseData(state, range)
  ) {
    if (!(await confirmMergeLoseData(strings, state, range, false, workbook))) return false;
  }

  const write = (): boolean => {
    switch (action) {
      case 'unmergeCells':
        return applyUnmerge(store, workbook, history, selection);
      case 'centerAcross':
        recordFormatChange(history, store, () => {
          applyFormatPatch(store.getState(), store, selection, { align: 'centerContinuous' });
        });
        return true;
      case 'mergeAcross':
        return applyMergeAcross(store, workbook, history, selection);
      case 'mergeCenter': {
        const merged = applyMerge(store, workbook, history, range);
        if (merged) {
          recordFormatChange(history, store, () => {
            applyFormatPatch(store.getState(), store, range, { align: 'center' });
          });
        }
        return merged;
      }
      case 'mergeCells':
        // Keep the explicit Merge Cells menu as a pure merge operation. In
        // particular, it must not toggle an exact existing merge or alter alignment.
        return applyMerge(store, workbook, history, range);
    }
  };
  const result = authorizedWrite({
    store,
    history: history ?? new History(),
    intents,
    origin,
    write,
  });
  return result.ok && result.value;
}
