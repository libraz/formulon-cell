// Cell Format ribbon dispatcher. Extracted from main.ts so the playground
// doesn't carry the dispatch table in module scope. Side effects live on the
// caller's side via the `deps` struct (status text, tab list refresh, etc).

import {
  hiddenInSelection,
  hideCols,
  hideRows,
  moveSheet,
  mutators,
  type Range,
  recordFormatChange,
  recordLayoutChange,
  renameSheet,
  type SpreadsheetInstance,
  setCellLocked,
  setSheetHidden,
  showCols,
  showRows,
  type ToolbarMenuText,
} from '@libraz/formulon-cell';
import { autofitColWidth, autofitRowHeight } from './autofit.js';

export interface CellFormatActionDeps {
  inst: SpreadsheetInstance | null;
  ribbonLang: 'ja' | 'en';
  /** Result of `normalizedSelectionRange()`; main.ts computes this. */
  range: Range | null;
  statusMetric: HTMLElement | null;
  ribbonMenuText: ToolbarMenuText;
  /** Sheet-tab "Rename" label from `dictionaries[ribbonLang].sheetTabs`. */
  renameSheetLabel: string;
  runSheetProtectionFlow: () => Promise<void>;
  showPrompt: (opts: {
    title: string;
    label: string;
    initial?: string;
    placeholder?: string;
    okLabel?: string;
    cancelLabel?: string;
    validate?: (value: string) => string | null;
  }) => Promise<string | null>;
  promptDimension: (
    title: string,
    label: string,
    initial: number,
    max: number,
  ) => Promise<number | null>;
  renderSheetTabs: () => void;
  switchSheet: (idx: number) => void;
  refreshWorkbookCells: () => void;
  sheetTabColorByAction: (action: string) => string | null | undefined;
  projectFormatToolbar: () => void;
  focusSheet: () => void;
}

export const applyCellFormatAction = async (
  action: string,
  deps: CellFormatActionDeps,
): Promise<void> => {
  const { inst: i, range, ribbonLang, statusMetric, ribbonMenuText } = deps;
  if (!i || !range) return;
  if (action === 'dialog') {
    i.openFormatDialog();
    return;
  }
  if (action === 'protect-sheet') {
    await deps.runSheetProtectionFlow();
    return;
  }
  if (action === 'rename-sheet') {
    const sheet = i.store.getState().data.sheetIndex;
    const current = i.workbook.sheetName(sheet);
    const name = await deps.showPrompt({
      title: deps.renameSheetLabel,
      label: ribbonLang === 'ja' ? 'シート名' : 'Sheet name',
      initial: current,
      validate: (raw) =>
        raw.trim()
          ? null
          : ribbonLang === 'ja'
            ? 'シート名を入力してください。'
            : 'Enter a sheet name.',
    });
    if (name !== null && renameSheet(i.workbook, sheet, name.trim(), i.store, i.history)) {
      deps.renderSheetTabs();
    }
    return;
  }
  if (action === 'move-sheet-left' || action === 'move-sheet-right') {
    const sheet = i.store.getState().data.sheetIndex;
    const target = action === 'move-sheet-left' ? sheet - 1 : sheet + 1;
    if (target >= 0 && target < i.workbook.sheetCount) {
      moveSheet(i.store, i.workbook, sheet, target, i.history);
      deps.renderSheetTabs();
    }
    return;
  }
  if (action === 'hide-sheet') {
    const sheet = i.store.getState().data.sheetIndex;
    if (setSheetHidden(i.store, i.workbook, i.history, sheet, true)) {
      deps.renderSheetTabs();
      deps.switchSheet(i.store.getState().data.sheetIndex);
      deps.refreshWorkbookCells();
    }
    return;
  }
  if (action === 'unhide-sheet') {
    const hidden = [...i.store.getState().layout.hiddenSheets].sort((a, b) => a - b)[0];
    if (hidden != null && setSheetHidden(i.store, i.workbook, i.history, hidden, false)) {
      deps.renderSheetTabs();
    }
    return;
  }
  const tabColor = deps.sheetTabColorByAction(action);
  if (tabColor !== undefined) {
    const sheet = i.store.getState().data.sheetIndex;
    recordLayoutChange(i.history, i.store, () => {
      mutators.setSheetTabColor(i.store, sheet, tabColor);
    });
    deps.renderSheetTabs();
    return;
  }
  if (action === 'lock-cell' || action === 'unlock-cell') {
    const locked = action === 'lock-cell';
    recordFormatChange(i.history, i.store, () => {
      setCellLocked(i.store, range, locked);
    });
    if (statusMetric) {
      statusMetric.textContent = locked
        ? ribbonMenuText.cellsLockedStatus
        : ribbonMenuText.cellsUnlockedStatus;
    }
    deps.projectFormatToolbar();
    deps.focusSheet();
    return;
  }
  if (action === 'hide-rows') {
    hideRows(i.store, i.history, range.r0, range.r1, i.workbook);
    return;
  }
  if (action === 'hide-cols') {
    hideCols(i.store, i.history, range.c0, range.c1, i.workbook);
    return;
  }
  if (action === 'show-rows') {
    const targets = hiddenInSelection(i.store.getState().layout, 'row', range.r0, range.r1);
    if (targets.length > 0)
      showRows(i.store, i.history, targets[0] ?? range.r0, targets.at(-1) ?? range.r1, i.workbook);
    return;
  }
  if (action === 'show-cols') {
    const targets = hiddenInSelection(i.store.getState().layout, 'col', range.c0, range.c1);
    if (targets.length > 0)
      showCols(i.store, i.history, targets[0] ?? range.c0, targets.at(-1) ?? range.c1, i.workbook);
    return;
  }
  if (action === 'row-height') {
    const n = await deps.promptDimension(
      ribbonLang === 'ja' ? '行の高さ' : 'Row Height',
      ribbonLang === 'ja' ? '高さ (px)' : 'Height (px)',
      i.store.getState().layout.defaultRowHeight,
      409,
    );
    if (n === null) return;
    recordLayoutChange(i.history, i.store, () => {
      for (let row = range.r0; row <= range.r1; row += 1) {
        mutators.setRowHeight(i.store, row, n);
        i.workbook.setRowHeight(range.sheet, row, n);
      }
    });
    return;
  }
  if (action === 'row-autofit') {
    recordLayoutChange(i.history, i.store, () => {
      for (let row = range.r0; row <= range.r1; row += 1) {
        const height = autofitRowHeight(i, row, range.c0, range.c1, ribbonLang);
        mutators.setRowHeight(i.store, row, height);
        i.workbook.setRowHeight(range.sheet, row, height);
      }
    });
    return;
  }
  if (action === 'col-width') {
    const n = await deps.promptDimension(
      ribbonLang === 'ja' ? '列の幅' : 'Column Width',
      ribbonLang === 'ja' ? '幅 (px)' : 'Width (px)',
      i.store.getState().layout.defaultColWidth,
      2048,
    );
    if (n === null) return;
    recordLayoutChange(i.history, i.store, () => {
      for (let col = range.c0; col <= range.c1; col += 1) {
        mutators.setColWidth(i.store, col, n);
        i.workbook.setColumnWidth(range.sheet, col, col, n);
      }
    });
    // NOTE: original code intentionally (or accidentally) falls through to
    // `col-autofit` here — preserved verbatim.
  }
  if (action === 'col-autofit') {
    recordLayoutChange(i.history, i.store, () => {
      for (let col = range.c0; col <= range.c1; col += 1) {
        const width = autofitColWidth(i, col, range.r0, range.r1, ribbonLang);
        mutators.setColWidth(i.store, col, width);
        i.workbook.setColumnWidth(range.sheet, col, col, width);
      }
    });
  }
};
