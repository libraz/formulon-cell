import { canExecuteBuiltIn } from '../../commands/built-in-command-policy.js';
import { runColumnWidthDialog, runRowHeightDialog } from '../../commands/row-col-dimension.js';
import {
  autofitColsWidthForRanges,
  autofitRowsHeightForRanges,
} from '../../commands/row-col-layout.js';
import { addrKey } from '../../engine/address.js';
import type { Locale } from '../../i18n/strings.js';
import { openFormatDialogFrom } from '../../interact/dialog-entry.js';
import { showMessage } from '../dialogs/prompt.js';

// Cell Format ribbon dispatcher. Extracted from main.ts so the playground
// doesn't carry the dispatch table in module scope. Side effects live on the
// caller's side via the `deps` struct (status text, tab list refresh, etc).

import { setCellLocked } from '../../commands/protection.js';
import {
  hiddenExtentInSelection,
  hideCols,
  hideRows,
  showCols,
  showRows,
} from '../../commands/row-col-layout.js';
import {
  copySheetAction,
  moveSheetAction,
  renameSheetAction,
  type SheetActionContext,
  setSheetHiddenAction,
  setSheetTabColorAction,
} from '../../commands/sheet-actions.js';
import { recordFormatChange } from '../../commands/slice-history.js';
import type { Range } from '../../engine/types.js';
import type { SpreadsheetInstance } from '../../mount.js';
import { resolveTheme } from '../../theme/resolve.js';
import { showMoveCopySheetDialog } from '../dialogs/move-copy-sheet.js';
import type { ToolbarMenuText } from '../menu-text.js';

type CellFormatMenuText = ToolbarMenuText & {
  sheetNameLabel: string;
  sheetNameRequired: string;
};

export interface CellFormatActionDeps {
  inst: SpreadsheetInstance | null;
  ribbonLang: Locale;
  /** Result of `normalizedSelectionRange()`; main.ts computes this. */
  range: Range | null;
  statusMetric: HTMLElement | null;
  ribbonMenuText: ToolbarMenuText;
  /** Sheet-tab "Rename" label from `dictionaries[ribbonLang].sheetTabs`. */
  renameSheetLabel: string;
  runSheetProtectionFlow: () => Promise<void>;
  showRenameSheetDialog: (opts: {
    title: string;
    label: string;
    initial: string;
    requiredMessage: string;
  }) => Promise<string | null>;
  promptDimension: (
    axis: 'row' | 'col',
    initial: number | undefined,
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
  const t = ribbonMenuText as CellFormatMenuText;
  if (action === 'dialog') {
    openFormatDialogFrom(i.store, undefined, { origin: 'ribbon', commandId: 'formatCells' });
    return;
  }
  if (action === 'protect-sheet') {
    await deps.runSheetProtectionFlow();
    return;
  }
  const sheetCtx: SheetActionContext = {
    store: i.store,
    workbook: i.workbook,
    history: i.history,
    origin: 'ribbon',
  };
  if (action === 'rename-sheet') {
    const sheet = i.store.getState().data.sheetIndex;
    const current = i.workbook.sheetName(sheet);
    const name = await deps.showRenameSheetDialog({
      title: deps.renameSheetLabel,
      label: t.sheetNameLabel,
      initial: current,
      requiredMessage: t.sheetNameRequired,
    });
    if (name !== null && renameSheetAction(sheetCtx, sheet, name)) {
      deps.renderSheetTabs();
    }
    return;
  }
  if (action === 'move-sheet-left' || action === 'move-sheet-right') {
    const sheet = i.store.getState().data.sheetIndex;
    const target = action === 'move-sheet-left' ? sheet - 1 : sheet + 1;
    if (target >= 0 && target < i.workbook.sheetCount) {
      moveSheetAction(sheetCtx, sheet, target);
      deps.renderSheetTabs();
    }
    return;
  }
  if (action === 'move-sheet-copy') {
    const sheet = i.store.getState().data.sheetIndex;
    const hidden = i.store.getState().layout.hiddenSheets;
    const shown = Array.from({ length: i.workbook.sheetCount }, (_, index) => index).filter(
      (index) => !hidden.has(index),
    );
    const sheetTabs = i.i18n.strings.sheetTabs;
    const result = await showMoveCopySheetDialog({
      host: i.host,
      strings: {
        title: sheetTabs.moveCopyTitle,
        targetBook: sheetTabs.moveCopyTargetBook,
        before: sheetTabs.moveCopyBefore,
        toEnd: sheetTabs.moveCopyToEnd,
        makeCopy: sheetTabs.moveCopyMakeCopy,
        ok: i.i18n.strings.hyperlinkDialog.ok,
        cancel: i.i18n.strings.hyperlinkDialog.cancel,
      },
      bookName: '',
      sheets: shown.map((index) => i.workbook.sheetName(index)),
      current: Math.max(0, shown.indexOf(sheet)),
      copyDisabledReason:
        canExecuteBuiltIn(i.store, 'copySheet', 'ribbon').allowed &&
        i.workbook.capabilities.sheetMutate
          ? null
          : i.i18n.strings.ribbonMenu.sheetActionUnavailable,
    });
    if (!result) return;
    const beforeSheet = shown[result.before];
    if (result.copy) {
      const copy = copySheetAction(sheetCtx, sheet, beforeSheet ?? i.workbook.sheetCount);
      if (copy < 0) {
        await showMessage({
          host: i.host,
          title: sheetTabs.moveCopyTitle,
          message: sheetTabs.moveCopyFailed,
          okLabel: i.i18n.strings.hyperlinkDialog.ok,
        });
        return;
      }
      // The copy is already the active sheet, loaded from the engine.
      deps.renderSheetTabs();
      return;
    }
    const last = shown.at(-1) ?? sheet;
    const to =
      beforeSheet === undefined ? last : sheet < beforeSheet ? beforeSheet - 1 : beforeSheet;
    if (moveSheetAction(sheetCtx, sheet, Math.max(0, Math.min(i.workbook.sheetCount - 1, to)))) {
      deps.renderSheetTabs();
    }
    return;
  }
  if (action === 'tab-color-more') {
    const sheet = i.store.getState().data.sheetIndex;
    const picker = document.createElement('input');
    picker.type = 'color';
    picker.value = i.store.getState().layout.sheetTabColors.get(sheet) ?? '#c00000';
    picker.style.cssText = 'position:fixed;left:-9999px;opacity:0';
    picker.addEventListener('change', () => {
      setSheetTabColorAction(sheetCtx, sheet, picker.value);
      deps.renderSheetTabs();
      picker.remove();
    });
    document.body.appendChild(picker);
    picker.click();
    return;
  }
  if (action === 'hide-sheet') {
    const sheet = i.store.getState().data.sheetIndex;
    if (setSheetHiddenAction(sheetCtx, sheet, true)) {
      deps.renderSheetTabs();
      deps.switchSheet(i.store.getState().data.sheetIndex);
      deps.refreshWorkbookCells();
    }
    return;
  }
  if (action === 'unhide-sheet') {
    const hidden = [...i.store.getState().layout.hiddenSheets].sort((a, b) => a - b)[0];
    if (hidden != null && setSheetHiddenAction(sheetCtx, hidden, false)) {
      deps.renderSheetTabs();
    }
    return;
  }
  const tabColor = deps.sheetTabColorByAction(action);
  if (tabColor !== undefined) {
    const sheet = i.store.getState().data.sheetIndex;
    setSheetTabColorAction(sheetCtx, sheet, tabColor);
    deps.renderSheetTabs();
    return;
  }
  if (action === 'lock-cell' || action === 'unlock-cell' || action === 'toggle-lock') {
    const first = i.store
      .getState()
      .format.formats.get(addrKey({ sheet: range.sheet, row: range.r0, col: range.c0 }));
    const locked = action === 'toggle-lock' ? first?.locked === false : action === 'lock-cell';
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
    const hidden = hiddenExtentInSelection(i.store.getState().layout, 'row', range.r0, range.r1);
    if (hidden) showRows(i.store, i.history, hidden.first, hidden.last, i.workbook);
    return;
  }
  if (action === 'show-cols') {
    const hidden = hiddenExtentInSelection(i.store.getState().layout, 'col', range.c0, range.c1);
    if (hidden) showCols(i.store, i.history, hidden.first, hidden.last, i.workbook);
    return;
  }
  if (action === 'row-height') {
    await runRowHeightDialog({
      store: i.store,
      history: i.history,
      wb: i.workbook,
      range,
      prompt: (initial, max) => deps.promptDimension('row', initial, max),
    });
    return;
  }
  if (action === 'row-autofit') {
    autofitRowsHeightForRanges(
      i.store,
      i.history,
      [range, ...(i.store.getState().selection.extraRanges ?? [])],
      i.workbook,
      {
        locale: ribbonLang,
        theme: resolveTheme(i.host),
      },
    );
    return;
  }
  if (action === 'col-width') {
    await runColumnWidthDialog({
      store: i.store,
      history: i.history,
      wb: i.workbook,
      range,
      prompt: (initial, max) => deps.promptDimension('col', initial, max),
    });
    return;
  }
  if (action === 'col-autofit') {
    autofitColsWidthForRanges(
      i.store,
      i.history,
      [range, ...(i.store.getState().selection.extraRanges ?? [])],
      i.workbook,
      {
        locale: ribbonLang,
        theme: resolveTheme(i.host),
      },
    );
  }
};
