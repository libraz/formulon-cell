import { applyFormatPatch } from '../commands/format.js';
import type { History } from '../commands/history.js';
import {
  applyMerge,
  applyMergeAcross,
  applyUnmerge,
  expandRangeWithMerges,
  mergeAcrossWillLoseData,
  mergeWillLoseData,
} from '../commands/merge.js';
import { recordFormatChange } from '../commands/slice-history.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { Strings } from '../i18n/strings.js';
import type { SpreadsheetStore } from '../store/store.js';
import { confirmMergeLoseData } from './dialogs/merge-confirm.js';

/** Actions exposed by the merge controls in wrappers and ribbon menus. */
export type MergeAction = 'mergeCenter' | 'mergeAcross' | 'mergeCells' | 'unmergeCells';

export interface MergeActionContext {
  store: SpreadsheetStore;
  workbook: WorkbookHandle;
  history: History | null;
  strings: Strings;
}

/**
 * Apply one merge-control action. The helper deliberately performs no UI
 * follow-up (focus, redraw, or projection); callers own those callbacks.
 *
 * The first await is reached only when a merge would discard data. That keeps
 * the no-warning path synchronous for the existing toolbar callers while
 * ensuring a cancelled warning leaves the workbook untouched.
 */
export async function applyMergeAction(
  context: MergeActionContext,
  action: MergeAction,
): Promise<boolean> {
  const { store, workbook, history, strings } = context;
  const selection = store.getState().selection.range;

  if (action === 'unmergeCells') {
    return applyUnmerge(store, workbook, history, selection);
  }

  if (action === 'mergeAcross') {
    const state = store.getState();
    if (mergeAcrossWillLoseData(state, selection)) {
      const confirmed = await confirmMergeLoseData(strings, state, selection, true);
      if (!confirmed) return false;
    }
    return applyMergeAcross(store, workbook, history, selection);
  }

  const state = store.getState();
  // A merge that intersects an existing merge acts on the complete merged
  // region. The command layer also normalizes this, but using the same target
  // for the warning and the mutation prevents a false-negative prompt.
  const range = expandRangeWithMerges(state, selection);
  if (mergeWillLoseData(state, range)) {
    const confirmed = await confirmMergeLoseData(strings, state, range);
    if (!confirmed) return false;
  }

  if (action === 'mergeCenter') {
    let merged = false;
    history?.begin();
    try {
      merged = applyMerge(store, workbook, history, range);
      if (merged) {
        recordFormatChange(history, store, () => {
          applyFormatPatch(store.getState(), store, range, { align: 'center' });
        });
      }
    } finally {
      history?.end();
    }
    return merged;
  }

  // Keep the explicit Merge Cells menu as a pure merge operation. In
  // particular, it must not toggle an exact existing merge or alter alignment.
  return applyMerge(store, workbook, history, range);
}
