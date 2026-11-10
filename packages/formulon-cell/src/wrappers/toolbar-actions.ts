// Framework-agnostic toolbar action handlers. Used by both the React and
// Vue wrappers so the per-button dispatch logic only exists in one place.
// Each handler accepts a `SpreadsheetInstance | null` (matching the prop
// shape both wrappers expose) and short-circuits if no instance is mounted.

import {
  type AutoSumFunction,
  autoSum,
  autoSumSuggestion,
  autoSumTargets,
} from '../commands/auto-sum.js';
import { canExecuteBuiltIn } from '../commands/built-in-command-policy.js';
import { deleteCells, insertCells } from '../commands/cell-shift.js';
import type { PasteSpecialOptions } from '../commands/clipboard/paste-special.js';
import { recordRecentFunction } from '../commands/function-history.js';
import { deferForAllowedRangeUnlock } from '../commands/protection.js';
import { applyFreezeMode, type FreezeMode } from '../commands/row-col-layout.js';
import {
  deleteSheetAction,
  insertSheetAction,
  type SheetActionContext,
} from '../commands/sheet-actions.js';
import { deleteCols, deleteVisibleRows, insertCols, insertRows } from '../commands/structure.js';
import type { SpreadsheetInstance } from '../mount/types.js';
import { mutators } from '../store/store.js';
import { confirmSheetDelete } from '../toolbar/dialogs/sheet-delete-confirm.js';
import type { MergeAction } from '../toolbar/merge-action.js';

/** A sheet command run from the ribbon's Insert or Delete menu. */
const ribbonSheetContext = (instance: SpreadsheetInstance): SheetActionContext => ({
  store: instance.store,
  workbook: instance.workbook,
  history: instance.history,
  origin: 'ribbon',
});

export type AutoSumAction = AutoSumFunction | 'MORE';

export type { MergeAction };

export type FreezeAction = FreezeMode;

export type PasteAction =
  | 'paste'
  | 'pasteFormulas'
  | 'pasteFormulasNumFmt'
  | 'pasteValues'
  | 'pasteValuesNumFmt'
  | 'pasteFormatsOnly'
  | 'pasteTranspose'
  | 'insertCopiedCells'
  | 'pasteSpecial';

export type CellInsertAction = 'shiftDown' | 'shiftRight' | 'rows' | 'cols' | 'sheet';
export type CellDeleteAction = 'shiftUp' | 'shiftLeft' | 'rows' | 'cols' | 'sheet';

/** Routes a ribbon Copy/Cut/Paste click through the same `runShortcut` the
 *  keyboard router uses, so React/Vue hosts using `createDefaultRibbonHooks`
 *  get working Copy/Paste buttons without re-implementing snapshot tracking.
 *
 *  Falls back to `document.execCommand` only when the clipboard handle is
 *  absent (the `clipboard` feature flag is off). That fallback is
 *  best-effort: the grid host is `user-select: none` so `execCommand` won't
 *  fire copy/paste events on it, but the user can still use Ctrl/⌘+C/V. */
export const dispatchHostClipboard = (
  instance: SpreadsheetInstance | null,
  kind: 'copy' | 'cut' | 'paste',
): void => {
  if (!instance) return;
  instance.host.focus();
  if (instance.clipboard) {
    void instance.clipboard.runShortcut(kind);
    return;
  }
  try {
    document.execCommand(kind);
  } catch {
    /* best-effort */
  }
};

const PASTE_SPECIAL_PRESETS: Record<
  Exclude<PasteAction, 'paste' | 'insertCopiedCells' | 'pasteSpecial'>,
  PasteSpecialOptions
> = {
  pasteFormulas: { what: 'formulas', operation: 'none', skipBlanks: false, transpose: false },
  pasteFormulasNumFmt: {
    what: 'formulas-and-numfmt',
    operation: 'none',
    skipBlanks: false,
    transpose: false,
  },
  pasteValues: { what: 'values', operation: 'none', skipBlanks: false, transpose: false },
  pasteValuesNumFmt: {
    what: 'values-and-numfmt',
    operation: 'none',
    skipBlanks: false,
    transpose: false,
  },
  pasteFormatsOnly: { what: 'formats', operation: 'none', skipBlanks: false, transpose: false },
  pasteTranspose: { what: 'all', operation: 'none', skipBlanks: false, transpose: true },
};

export const handlePasteAction = (
  instance: SpreadsheetInstance | null,
  action: PasteAction,
): void => {
  if (!instance) return;
  if (action === 'paste') {
    dispatchHostClipboard(instance, 'paste');
    return;
  }
  if (action === 'insertCopiedCells') {
    instance.openInsertCopiedCells();
    return;
  }
  if (action === 'pasteSpecial') {
    instance.openPasteSpecial();
    return;
  }
  const preset = PASTE_SPECIAL_PRESETS[action];
  if (preset) instance.pasteSpecial(preset);
};

export const handleFreezeAction = (
  instance: SpreadsheetInstance | null,
  action: FreezeAction,
): void => {
  if (!instance) return;
  applyFreezeMode(instance.store, action, instance.workbook);
};

/** AutoSum from the toolbar. An empty single cell opens the suggested
 *  formula in the cell editor; any other selection is written directly. The
 *  "MORE" sentinel opens the function arguments dialog instead. */
export const handleAutoSum = (
  instance: SpreadsheetInstance | null,
  functionName: AutoSumFunction = 'SUM',
): boolean => {
  if (!instance) return false;
  const suggestion = autoSumSuggestion(instance.store.getState(), functionName);
  if (suggestion) {
    instance.cellEditor.begin(suggestion.formula, { mode: 'edit' });
    recordRecentFunction(instance.store, functionName);
    return true;
  }
  const run = (): boolean => {
    const r = instance.store.getState().selection.range;
    const single = r.r0 === r.r1 && r.c0 === r.c1;
    instance.history.begin();
    let result: ReturnType<typeof autoSum> = null;
    try {
      result = autoSum(instance.store.getState(), instance.workbook, functionName);
    } finally {
      instance.history.end();
    }
    if (!result) return false;
    mutators.replaceCells(instance.store, instance.workbook.cells(result.addr.sheet));
    // A range selection stays as it was; a single cell moves to the new total.
    if (single) mutators.setActive(instance.store, result.addr);
    recordRecentFunction(instance.store, functionName);
    return true;
  };
  const targets = autoSumTargets(instance.store.getState());
  // A password prompt makes the AutoSum asynchronous; it reports as not done yet.
  return deferForAllowedRangeUnlock(instance.store, targets, run) ? false : run();
};

export const handleAutoSumAction = (
  instance: SpreadsheetInstance | null,
  action: AutoSumAction,
): boolean => {
  if (!instance) return false;
  if (action === 'MORE') {
    instance.openFunctionArguments();
    return true;
  }
  return handleAutoSum(instance, action);
};

const insertSelectedRows = (instance: SpreadsheetInstance | null): void => {
  if (!instance) return;
  const r = instance.store.getState().selection.range;
  insertRows(instance.store, instance.workbook, instance.history, r.r0, r.r1 - r.r0 + 1);
};

const deleteSelectedRows = (instance: SpreadsheetInstance | null): void => {
  if (!instance) return;
  const r = instance.store.getState().selection.range;
  deleteVisibleRows(instance.store, instance.workbook, instance.history, r.r0, r.r1 - r.r0 + 1);
};

const insertSelectedCols = (instance: SpreadsheetInstance | null): void => {
  if (!instance) return;
  const r = instance.store.getState().selection.range;
  insertCols(instance.store, instance.workbook, instance.history, r.c0, r.c1 - r.c0 + 1);
};

const deleteSelectedCols = (instance: SpreadsheetInstance | null): void => {
  if (!instance) return;
  const r = instance.store.getState().selection.range;
  deleteCols(instance.store, instance.workbook, instance.history, r.c0, r.c1 - r.c0 + 1);
};

export const handleInsertCellsAction = (
  instance: SpreadsheetInstance | null,
  action: CellInsertAction,
): void => {
  if (!instance) return;
  const r = instance.store.getState().selection.range;
  if (action === 'rows') {
    insertSelectedRows(instance);
    return;
  }
  if (action === 'cols') {
    insertSelectedCols(instance);
    return;
  }
  if (action === 'sheet') {
    const added = insertSheetAction(ribbonSheetContext(instance));
    if (added >= 0) mutators.setSheetIndex(instance.store, added);
    return;
  }
  insertCells(
    instance.store,
    instance.workbook,
    instance.history,
    r,
    action === 'shiftDown' ? 'down' : 'right',
  );
};

export const handleDeleteCellsAction = (
  instance: SpreadsheetInstance | null,
  action: CellDeleteAction,
): void => {
  if (!instance) return;
  const r = instance.store.getState().selection.range;
  if (action === 'rows') {
    deleteSelectedRows(instance);
    return;
  }
  if (action === 'cols') {
    deleteSelectedCols(instance);
    return;
  }
  if (action === 'sheet') {
    const sheet = instance.store.getState().data.sheetIndex;
    // A refused delete is decided before the confirmation, not after it.
    if (!canExecuteBuiltIn(instance.store, 'deleteSheet', 'ribbon').allowed) return;
    void confirmSheetDelete(instance.i18n.strings, instance.workbook, sheet).then((ok) => {
      if (ok) deleteSheetAction(ribbonSheetContext(instance), sheet);
    });
    return;
  }
  deleteCells(
    instance.store,
    instance.workbook,
    instance.history,
    r,
    action === 'shiftUp' ? 'up' : 'left',
  );
};
