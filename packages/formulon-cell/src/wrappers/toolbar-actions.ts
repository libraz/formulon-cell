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
import type { PasteSpecialOptions } from '../commands/clipboard/paste-special.js';
import { recordRecentFunction } from '../commands/function-history.js';
import { deferForAllowedRangeUnlock } from '../commands/protection.js';
import {
  deleteCellsAction,
  deleteColumnsAction,
  deleteRowsAction,
  insertCellsAction,
  insertColumnsAction,
  insertRowsAction,
  type RowColActionContext,
} from '../commands/row-col-actions.js';
import { applyFreezeMode, type FreezeMode } from '../commands/row-col-layout.js';
import {
  deleteSheetAction,
  insertSheetAction,
  type SheetActionContext,
} from '../commands/sheet-actions.js';
import { runClipboardShortcutFrom } from '../interact/clipboard.js';
import {
  applyPasteSpecialFrom,
  openInsertCopiedCellsFrom,
  openPasteSpecialFrom,
} from '../interact/dialog-entry.js';
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
    void runClipboardShortcutFrom(instance.clipboard, kind, { origin: 'ribbon', commandId: kind });
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
    openInsertCopiedCellsFrom(instance.store, { origin: 'ribbon', commandId: 'insertCopiedCells' });
    return;
  }
  if (action === 'pasteSpecial') {
    openPasteSpecialFrom(instance.store, { origin: 'ribbon', commandId: 'pasteSpecial' });
    return;
  }
  const preset = PASTE_SPECIAL_PRESETS[action];
  if (preset)
    applyPasteSpecialFrom(instance.store, preset, { origin: 'ribbon', commandId: action });
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

/** A Home Insert (Delete) menu item, run as its own command. */
const homeCellsContext = (
  instance: SpreadsheetInstance,
  commandId: string,
): RowColActionContext => ({
  store: instance.store,
  workbook: instance.workbook,
  history: instance.history,
  origin: 'ribbon',
  commandId,
});

export const handleInsertCellsAction = (
  instance: SpreadsheetInstance | null,
  action: CellInsertAction,
): void => {
  if (!instance) return;
  const r = instance.store.getState().selection.range;
  if (action === 'rows') {
    insertRowsAction(homeCellsContext(instance, 'rowInsertAbove'), r.r0, r.r1 - r.r0 + 1);
    return;
  }
  if (action === 'cols') {
    insertColumnsAction(homeCellsContext(instance, 'colInsertLeft'), r.c0, r.c1 - r.c0 + 1);
    return;
  }
  if (action === 'sheet') {
    const added = insertSheetAction(ribbonSheetContext(instance));
    if (added >= 0) mutators.setSheetIndex(instance.store, added);
    return;
  }
  insertCellsAction(
    homeCellsContext(instance, 'insertCells'),
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
    deleteRowsAction(homeCellsContext(instance, 'rowDelete'), r.r0, r.r1 - r.r0 + 1);
    return;
  }
  if (action === 'cols') {
    deleteColumnsAction(homeCellsContext(instance, 'colDelete'), r.c0, r.c1 - r.c0 + 1);
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
  deleteCellsAction(
    homeCellsContext(instance, 'deleteCells'),
    r,
    action === 'shiftUp' ? 'up' : 'left',
  );
};
