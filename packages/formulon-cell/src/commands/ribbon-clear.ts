// Shared "Clear" ribbon split-button action — invoked by every host wrapper.
// Each variant either targets a single cross-cutting concern (formats, comments,
// hyperlinks, conditional rules) or piles them all together under "all". The
// host wrappers stay thin: they read the active range from the store, call
// here, and let the helper own the history/transaction wiring.

import { flushFormatToEngine } from '../engine/cell-format-sync.js';
import type { Addr, Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { type CellFormat, mutators, type SpreadsheetStore } from '../store/store.js';
import { clearComment } from './comment.js';
import { clearConditionalRulesInRange } from './conditional-format.js';
import { clearFormat, clearVisualFormat } from './format.js';
import { type History, recordConditionalRulesChange, recordFormatChange } from './history.js';
import { clearHyperlink } from './hyperlinks.js';
import { isCellWritable } from './protection.js';
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

const rangeContainsAddr = (range: Range, addr: Addr): boolean =>
  addr.sheet === range.sheet &&
  addr.row >= range.r0 &&
  addr.row <= range.r1 &&
  addr.col >= range.c0 &&
  addr.col <= range.c1;

const addrFromKey = (key: string): Addr | null => {
  const parts = key.split(':').map((n) => Number(n));
  const sheet = parts[0];
  const row = parts[1];
  const col = parts[2];
  if (
    typeof sheet !== 'number' ||
    typeof row !== 'number' ||
    typeof col !== 'number' ||
    !Number.isInteger(sheet) ||
    !Number.isInteger(row) ||
    !Number.isInteger(col)
  ) {
    return null;
  }
  return { sheet, row, col };
};

const physicalCellAddrsInRange = (workbook: WorkbookHandle, range: Range): Addr[] => {
  const source =
    typeof (workbook as WorkbookHandle & { physicalCells?: WorkbookHandle['cells'] })
      .physicalCells === 'function'
      ? (workbook as WorkbookHandle & { physicalCells: WorkbookHandle['cells'] }).physicalCells(
          range.sheet,
        )
      : workbook.cells(range.sheet);
  const out: Addr[] = [];
  for (const cell of source) {
    if (rangeContainsAddr(range, cell.addr)) out.push(cell.addr);
  }
  return out;
};

const formattedAddrsInRange = (
  store: SpreadsheetStore,
  range: Range,
  predicate: (format: CellFormat) => boolean,
): Addr[] => {
  const out: Addr[] = [];
  for (const [key, format] of store.getState().format.formats) {
    if (!predicate(format)) continue;
    const addr = addrFromKey(key);
    if (addr && rangeContainsAddr(range, addr)) out.push(addr);
  }
  return out;
};

export const executeRibbonClearAction = (deps: ExecuteRibbonClearActionDeps): void => {
  const { store, workbook, history, action } = deps;
  const range = store.getState().selection.range;
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
      for (const addr of physicalCellAddrsInRange(workbook, range)) {
        if (!isCellWritable(store.getState(), addr)) continue;
        workbook.setBlank(addr);
      }
    }
    if (action === 'comments' || action === 'all') {
      recordFormatChange(history, store, () => {
        const addrs = formattedAddrsInRange(
          store,
          range,
          (format) => typeof format.comment === 'string' && format.comment.length > 0,
        );
        for (const addr of addrs) clearComment(store, addr, workbook);
      });
    }
    if (action === 'hyperlinks' || action === 'all') {
      recordFormatChange(history, store, () => {
        const addrs = formattedAddrsInRange(store, range, (format) => !!format.hyperlink);
        for (const addr of addrs) clearHyperlink(store, addr, workbook);
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
