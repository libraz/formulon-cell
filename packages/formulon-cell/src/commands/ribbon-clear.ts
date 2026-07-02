// Shared "Clear" ribbon split-button action — invoked by every host wrapper.
// Each variant either targets a single cross-cutting concern (formats, comments,
// hyperlinks, conditional rules) or piles them all together under "all". The
// host wrappers stay thin: they read the active range from the store, call
// here, and let the helper own the history/transaction wiring.

import { flushFormatToEngine } from '../engine/cell-format-sync.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { mutators, type SpreadsheetStore } from '../store/store.js';
import { clearComment } from './comment.js';
import { clearConditionalRulesInRange } from './conditional-format.js';
import { clearFormat, clearVisualFormat } from './format.js';
import { type History, recordConditionalRulesChange, recordFormatChange } from './history.js';
import { clearHyperlink } from './hyperlinks.js';
import { writableAddrs } from './protection.js';
import { clearValidationInRangeWithEngine } from './validate.js';

export type RibbonClearAction =
  | 'all'
  | 'formats'
  | 'contents'
  | 'comments'
  | 'hyperlinks'
  | 'conditional';

export interface ExecuteRibbonClearActionDeps {
  store: SpreadsheetStore;
  workbook: WorkbookHandle;
  history: History;
  action: RibbonClearAction;
}

export const executeRibbonClearAction = (deps: ExecuteRibbonClearActionDeps): void => {
  const { store, workbook, history, action } = deps;
  const range = store.getState().selection.range;
  const eachCell = (fn: (row: number, col: number) => void): void => {
    for (let row = range.r0; row <= range.r1; row += 1) {
      for (let col = range.c0; col <= range.c1; col += 1) fn(row, col);
    }
  };
  if (action === 'formats') {
    recordFormatChange(history, store, () => clearVisualFormat(store.getState(), store));
    // Reset the engine XF for the cleared cells so the format does not
    // resurrect on the next save.
    flushFormatToEngine(workbook, store, range.sheet);
    return;
  }
  if (action === 'conditional') {
    recordConditionalRulesChange(history, store, () => {
      clearConditionalRulesInRange(store, range);
    });
    return;
  }
  history.begin();
  try {
    if (action === 'contents' || action === 'all') {
      for (const addr of writableAddrs(store.getState(), range)) {
        workbook.setBlank(addr);
      }
    }
    if (action === 'comments' || action === 'all') {
      recordFormatChange(history, store, () => {
        eachCell((row, col) => clearComment(store, { sheet: range.sheet, row, col }, workbook));
      });
    }
    if (action === 'hyperlinks' || action === 'all') {
      recordFormatChange(history, store, () => {
        eachCell((row, col) => clearHyperlink(store, { sheet: range.sheet, row, col }, workbook));
      });
    }
    if (action === 'all') {
      clearValidationInRangeWithEngine(store, history, workbook, range);
      recordFormatChange(history, store, () => {
        clearFormat(store.getState(), store);
      });
      recordConditionalRulesChange(history, store, () => {
        clearConditionalRulesInRange(store, range);
      });
    }
    if (action === 'all') {
      // Flush the cleared formats to the engine so the XF resets.
      flushFormatToEngine(workbook, store, range.sheet);
    }
  } finally {
    history.end();
  }
  mutators.replaceCells(store, workbook.cells(range.sheet));
};
