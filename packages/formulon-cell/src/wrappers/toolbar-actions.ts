// Framework-agnostic toolbar action handlers. Used by both the React and
// Vue wrappers so the per-button dispatch logic only exists in one place.
// Each handler accepts a `SpreadsheetInstance | null` (matching the prop
// shape both wrappers expose) and short-circuits if no instance is mounted.

import type { AutoSumFunction } from '../commands/auto-sum.js';
import type { PasteSpecialOptions } from '../commands/clipboard/paste-special.js';
import type { ConditionalPresetAction } from '../commands/conditional-format.js';
import {
  addSheet,
  applyConditionalPresetAction,
  applyMerge,
  applyUnmerge,
  autoSum,
  deleteCells,
  deleteCols,
  deleteRows,
  hiddenInSelection,
  hideCols,
  hideRows,
  insertCells,
  insertCols,
  insertRows,
  mutators,
  recordConditionalRulesChange,
  recordFormatChange,
  recordMergesChangeWithEngine,
  removeSheet,
  setAlign,
  setFreezePanes,
  showColsAroundSelection,
  showRowsAroundSelection,
} from '../index.js';
import type { SpreadsheetInstance } from '../mount/types.js';

export type AutoSumAction = AutoSumFunction | 'MORE';

export type MergeAction = 'mergeCenter' | 'mergeAcross' | 'mergeCells' | 'unmergeCells';

export type FreezeAction = 'none' | 'topRow' | 'firstColumn' | 'panes';

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

export type WindowAction = 'hideRows' | 'showRows' | 'hideCols' | 'showCols';

export type ConditionalMenuAction =
  | ConditionalPresetAction
  | 'new-rule'
  | 'manage'
  | 'highlight-more'
  | 'top-bottom-more'
  | 'data-bars-more'
  | 'color-scales-more'
  | 'icon-sets-more'
  | 'cell-greater'
  | 'cell-less'
  | 'cell-between'
  | 'cell-equal'
  | 'text-contains'
  | 'date-occurring';

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

export const handleMergeAction = (
  instance: SpreadsheetInstance | null,
  action: MergeAction,
): void => {
  if (!instance) return;
  const s = instance.store.getState();
  const r = s.selection.range;
  if (action === 'unmergeCells') {
    applyUnmerge(instance.store, instance.workbook, instance.history, r);
    return;
  }
  if (action === 'mergeAcross') {
    recordMergesChangeWithEngine(
      instance.history,
      instance.store,
      instance.workbook,
      r.sheet,
      () => {
        for (let row = r.r0; row <= r.r1; row += 1) {
          if (r.c0 === r.c1) continue;
          mutators.mergeRange(instance.store, {
            sheet: r.sheet,
            r0: row,
            c0: r.c0,
            r1: row,
            c1: r.c1,
          });
        }
      },
    );
    return;
  }
  applyMerge(instance.store, instance.workbook, instance.history, r);
  if (action === 'mergeCenter') {
    recordFormatChange(instance.history, instance.store, () =>
      setAlign(instance.store.getState(), instance.store, 'center'),
    );
  }
};

export const handleFreezeAction = (
  instance: SpreadsheetInstance | null,
  action: FreezeAction,
): void => {
  if (!instance) return;
  const s = instance.store.getState();
  if (action === 'none') {
    setFreezePanes(instance.store, instance.history, 0, 0, instance.workbook);
    return;
  }
  if (action === 'topRow') {
    setFreezePanes(instance.store, instance.history, 1, 0, instance.workbook);
    return;
  }
  if (action === 'firstColumn') {
    setFreezePanes(instance.store, instance.history, 0, 1, instance.workbook);
    return;
  }
  const a = s.selection.active;
  // When the cursor is at A1 freeze just the top row — matches Excel's
  // "Freeze Panes" default which freezes everything above + left of the
  // active cell.
  const rows = a.row === 0 && a.col === 0 ? 1 : a.row;
  const cols = a.row === 0 && a.col === 0 ? 0 : a.col;
  setFreezePanes(instance.store, instance.history, rows, cols, instance.workbook);
};

/** AutoSum from the toolbar. The "MORE" sentinel opens the function
 *  arguments dialog instead of inserting an aggregate. */
export const handleAutoSum = (
  instance: SpreadsheetInstance | null,
  functionName: AutoSumFunction = 'SUM',
): boolean => {
  if (!instance) return false;
  instance.history.begin();
  let result: ReturnType<typeof autoSum> = null;
  try {
    result = autoSum(instance.store.getState(), instance.workbook, functionName);
  } finally {
    instance.history.end();
  }
  if (!result) return false;
  mutators.replaceCells(instance.store, instance.workbook.cells(result.addr.sheet));
  mutators.setActive(instance.store, result.addr);
  return true;
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

export const insertSelectedRows = (instance: SpreadsheetInstance | null): void => {
  if (!instance) return;
  const r = instance.store.getState().selection.range;
  insertRows(instance.store, instance.workbook, instance.history, r.r0, r.r1 - r.r0 + 1);
};

export const deleteSelectedRows = (instance: SpreadsheetInstance | null): void => {
  if (!instance) return;
  const r = instance.store.getState().selection.range;
  deleteRows(instance.store, instance.workbook, instance.history, r.r0, r.r1 - r.r0 + 1);
};

export const insertSelectedCols = (instance: SpreadsheetInstance | null): void => {
  if (!instance) return;
  const r = instance.store.getState().selection.range;
  insertCols(instance.store, instance.workbook, instance.history, r.c0, r.c1 - r.c0 + 1);
};

export const deleteSelectedCols = (instance: SpreadsheetInstance | null): void => {
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
    const added = addSheet(instance.store, instance.workbook, instance.history);
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
    removeSheet(instance.store, instance.workbook, instance.store.getState().data.sheetIndex);
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

export const handleWindowAction = (
  instance: SpreadsheetInstance | null,
  action: WindowAction,
): void => {
  if (!instance) return;
  const r = instance.store.getState().selection.range;
  if (action === 'hideRows') {
    hideRows(instance.store, instance.history, r.r0, r.r1, instance.workbook);
  } else if (action === 'showRows') {
    showRowsAroundSelection(instance.store, instance.history, r.r0, r.r1, instance.workbook);
  } else if (action === 'hideCols') {
    hideCols(instance.store, instance.history, r.c0, r.c1, instance.workbook);
  } else {
    showColsAroundSelection(instance.store, instance.history, r.c0, r.c1, instance.workbook);
  }
};

/** Mirrors the toolbar's "hide/show rows" toggle: if any rows in the
 *  selection are already hidden, restore them; otherwise hide the lot. */
export const toggleSelectedRowsHidden = (instance: SpreadsheetInstance | null): void => {
  if (!instance) return;
  const s = instance.store.getState();
  const r = s.selection.range;
  if (hiddenInSelection(s.layout, 'row', r.r0, r.r1).length > 0) {
    showRowsAroundSelection(instance.store, instance.history, r.r0, r.r1, instance.workbook);
  } else {
    hideRows(instance.store, instance.history, r.r0, r.r1, instance.workbook);
  }
};

export const toggleSelectedColsHidden = (instance: SpreadsheetInstance | null): void => {
  if (!instance) return;
  const s = instance.store.getState();
  const r = s.selection.range;
  if (hiddenInSelection(s.layout, 'col', r.c0, r.c1).length > 0) {
    showColsAroundSelection(instance.store, instance.history, r.c0, r.c1, instance.workbook);
  } else {
    hideCols(instance.store, instance.history, r.c0, r.c1, instance.workbook);
  }
};

export const handleConditionalAction = (
  instance: SpreadsheetInstance | null,
  action: ConditionalMenuAction,
): void => {
  if (!instance) return;
  if (action === 'new-rule') {
    instance.openConditionalDialog({ mode: 'new' });
    return;
  }
  if (action === 'cell-greater') {
    instance.openConditionalDialog({ mode: 'new', kind: 'cell-value', cellValueOp: '>' });
    return;
  }
  if (action === 'cell-less') {
    instance.openConditionalDialog({ mode: 'new', kind: 'cell-value', cellValueOp: '<' });
    return;
  }
  if (action === 'cell-between') {
    instance.openConditionalDialog({ mode: 'new', kind: 'cell-value', cellValueOp: 'between' });
    return;
  }
  if (action === 'cell-equal') {
    instance.openConditionalDialog({ mode: 'new', kind: 'cell-value', cellValueOp: '=' });
    return;
  }
  if (action === 'text-contains') {
    instance.openConditionalDialog({ mode: 'new', kind: 'text-contains' });
    return;
  }
  if (action === 'date-occurring') {
    instance.openConditionalDialog({ mode: 'new', kind: 'date-occurring', datePeriod: 'today' });
    return;
  }
  if (action === 'manage') {
    instance.openCfRulesDialog();
    return;
  }
  if (action === 'highlight-more') {
    instance.openConditionalDialog({ mode: 'new', kind: 'cell-value' });
    return;
  }
  if (action === 'top-bottom-more') {
    instance.openConditionalDialog({ mode: 'new', kind: 'top-bottom' });
    return;
  }
  if (action === 'data-bars-more') {
    instance.openConditionalDialog({ mode: 'new', kind: 'data-bar' });
    return;
  }
  if (action === 'color-scales-more') {
    instance.openConditionalDialog({ mode: 'new', kind: 'color-scale' });
    return;
  }
  if (action === 'icon-sets-more') {
    instance.openConditionalDialog({ mode: 'new', kind: 'icon-set' });
    return;
  }
  recordConditionalRulesChange(instance.history, instance.store, () => {
    applyConditionalPresetAction(instance.store, action);
  });
};
