import { insertCopiedBand } from '../../commands/clipboard/insert-copied-cells.js';
import { isWorkbookStructureProtected } from '../../commands/protection.js';
import { addrKey } from '../../engine/address.js';
import { dictionaryLocaleFor } from '../../i18n/strings.js';
import { openCellShiftDialog, runCellShiftChoice } from '../../interact/cell-shift-dialog.js';
import type { SpreadsheetInstance } from '../../mount.js';
import { sheetTabColorActionForColor, sheetTabColorByAction } from '../../sheet-tab-colors.js';
import { isWholeColumnRange, isWholeRowRange } from '../../store/selection-geometry.js';
import { mutators } from '../../store/store.js';
import { showDimensionDialog } from '../../toolbar/dialogs/dimension.js';
import { showRenameSheetDialog } from '../../toolbar/dialogs/rename-sheet.js';
import { applyCellFormatAction } from '../../toolbar/ribbon/cell-format-action.js';
import type { DynamicDropdownsCtx } from '../../toolbar/ribbon/dynamic-dropdowns.js';
import {
  handleDeleteCellsAction,
  handleInsertCellsAction,
} from '../../wrappers/toolbar-actions.js';
import type { DefaultDynamicDropdownsOptions } from '../dynamic-dropdowns-defaults.js';
import { setMenuControlDisabled, showInstanceReport } from './menu-feedback.js';
import { hasActiveCopy, normalizedSelectionRange } from './selection.js';

const noop = (): void => undefined;

const cellShiftTarget = (instance: SpreadsheetInstance) => ({
  store: instance.store,
  wb: instance.workbook,
  history: instance.history,
});

const buildCellInsertAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyCellInsertAction'] =>
  (action) => {
    if (action === 'cells') {
      const range = normalizedSelectionRange(instance);
      if (hasActiveCopy(instance)) {
        const snapshot = instance.clipboard?.getSnapshot();
        const sourceRange = snapshot?.logicalRange ?? snapshot?.range;
        const sourceIsWholeBand =
          sourceRange !== undefined &&
          (isWholeRowRange(sourceRange) || isWholeColumnRange(sourceRange));
        if (snapshot && sourceIsWholeBand) {
          const inserted = insertCopiedBand(
            instance.store,
            instance.workbook,
            instance.history,
            snapshot,
            range,
          );
          if (inserted) {
            mutators.replaceCells(
              instance.store,
              instance.workbook.cells(instance.store.getState().data.sheetIndex),
            );
            instance.host.focus();
            return;
          }
          return;
        }
      }
      if (!hasActiveCopy(instance) && (isWholeRowRange(range) || isWholeColumnRange(range))) {
        handleInsertCellsAction(instance, isWholeRowRange(range) ? 'rows' : 'cols');
        mutators.replaceCells(
          instance.store,
          instance.workbook.cells(instance.store.getState().data.sheetIndex),
        );
        instance.host.focus();
        return;
      }
      openCellShiftDialog({
        host: instance.host,
        strings: instance.i18n.strings,
        kind: 'insert',
        onSubmit: (choice) => {
          if (!runCellShiftChoice(cellShiftTarget(instance), 'insert', range, choice)) return;
          mutators.replaceCells(
            instance.store,
            instance.workbook.cells(instance.store.getState().data.sheetIndex),
          );
          instance.host.focus();
        },
      });
      return;
    }
    const mapped: Record<string, Parameters<typeof handleInsertCellsAction>[1]> = {
      'shift-down': 'shiftDown',
      'shift-right': 'shiftRight',
      rows: 'rows',
      cols: 'cols',
      sheet: 'sheet',
    };
    const next = mapped[action];
    if (!next) return;
    handleInsertCellsAction(instance, next);
    instance.host.focus();
  };

/** Excel for Mac leaves Hide Sheet and Delete Sheet enabled on a one-sheet workbook and answers with an alert. */
const isMacInstance = (instance: SpreadsheetInstance): boolean =>
  (instance.host.closest<HTMLElement>('.fc-host') ?? instance.host).dataset.fcPlatform === 'mac';

const reportSheetGuard = (instance: SpreadsheetInstance, message: string): void => {
  void showInstanceReport(instance, message, [{ severity: 'warning', label: message, detail: '' }]);
};

const buildCellDeleteAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyCellDeleteAction'] =>
  (action) => {
    if (action === 'sheet' && isMacInstance(instance) && instance.workbook.sheetCount <= 1) {
      reportSheetGuard(instance, instance.i18n.strings.ribbonMenu.sheetDeleteRequiresAnotherSheet);
      return;
    }
    if (action === 'cells') {
      const range = normalizedSelectionRange(instance);
      if (isWholeRowRange(range) || isWholeColumnRange(range)) {
        handleDeleteCellsAction(instance, isWholeRowRange(range) ? 'rows' : 'cols');
        mutators.replaceCells(
          instance.store,
          instance.workbook.cells(instance.store.getState().data.sheetIndex),
        );
        instance.host.focus();
        return;
      }
      openCellShiftDialog({
        host: instance.host,
        strings: instance.i18n.strings,
        kind: 'delete',
        onSubmit: (choice) => {
          if (!runCellShiftChoice(cellShiftTarget(instance), 'delete', range, choice)) return;
          mutators.replaceCells(
            instance.store,
            instance.workbook.cells(instance.store.getState().data.sheetIndex),
          );
          instance.host.focus();
        },
      });
      return;
    }
    const mapped: Record<string, Parameters<typeof handleDeleteCellsAction>[1]> = {
      'shift-up': 'shiftUp',
      'shift-left': 'shiftLeft',
      rows: 'rows',
      cols: 'cols',
      row: 'rows',
      col: 'cols',
      sheet: 'sheet',
    };
    const next = mapped[action];
    if (!next) return;
    handleDeleteCellsAction(instance, next);
    instance.host.focus();
  };

const buildCellFormatAction =
  (
    instance: SpreadsheetInstance,
    opts: Pick<
      DefaultDynamicDropdownsOptions,
      'projectFormatToolbar' | 'refreshCells' | 'renderSheetTabs'
    >,
  ): DynamicDropdownsCtx['applyCellFormatAction'] =>
  async (action) => {
    const strings = instance.i18n.strings;
    if (action === 'hide-sheet' && isMacInstance(instance)) {
      const state = instance.store.getState();
      if (instance.workbook.sheetCount - state.layout.hiddenSheets.size <= 1) {
        reportSheetGuard(instance, strings.ribbonMenu.sheetHideRequiresVisibleSheet);
        return;
      }
    }
    await applyCellFormatAction(action, {
      inst: instance,
      ribbonLang: dictionaryLocaleFor(instance.i18n.locale),
      range: normalizedSelectionRange(instance),
      statusMetric: null,
      ribbonMenuText: strings.ribbonMenu,
      renameSheetLabel: strings.sheetTabs.rename,
      runSheetProtectionFlow: async () => {
        instance.toggleSheetProtection();
      },
      showRenameSheetDialog: (opts) =>
        showRenameSheetDialog({
          okLabel: strings.hyperlinkDialog.ok,
          cancelLabel: strings.hyperlinkDialog.cancel,
          ...opts,
        }),
      promptDimension: (axis, initial, max) =>
        showDimensionDialog({
          axis,
          strings,
          locale: () => instance.locale,
          initial,
          max,
        }),
      renderSheetTabs: opts.renderSheetTabs ?? noop,
      switchSheet: (idx) => {
        mutators.setSheetIndex(instance.store, idx);
      },
      refreshWorkbookCells:
        opts.refreshCells ??
        (() => {
          mutators.replaceCells(
            instance.store,
            instance.workbook.cells(instance.store.getState().data.sheetIndex),
          );
        }),
      sheetTabColorByAction,
      projectFormatToolbar: opts.projectFormatToolbar ?? noop,
      focusSheet: () => instance.host.focus(),
    });
    instance.host.focus();
  };

const updateFormatCellsMenu =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['updateFormatCellsMenu'] =>
  (menu) => {
    const state = instance.store.getState();
    const range = normalizedSelectionRange(instance);
    const representativeFormat = state.format.formats.get(
      addrKey({ sheet: range.sheet, row: range.r0, col: range.c0 }),
    );
    const activeLocked = representativeFormat?.locked !== false;
    const sheet = state.data.sheetIndex;
    const hiddenSheetCount = state.layout.hiddenSheets.size;
    const visibleSheetCount = instance.workbook.sheetCount - hiddenSheetCount;
    const canMoveSheet = instance.workbook.capabilities.sheetMutate;
    const activeTabColorAction = sheetTabColorActionForColor(
      state.layout.sheetTabColors.get(sheet),
    );
    const reasonForFormatAction = (action: string | undefined): string | undefined => {
      const t = instance.i18n.strings.ribbonMenu;
      if (
        (action === 'rename-sheet' ||
          action === 'move-sheet-copy' ||
          action === 'move-sheet-left' ||
          action === 'move-sheet-right' ||
          action === 'hide-sheet' ||
          action === 'unhide-sheet') &&
        !canMoveSheet
      ) {
        return t.sheetActionUnavailable;
      }
      if (action === 'move-sheet-left' && sheet <= 0) return t.sheetMoveAtBoundary;
      if (action === 'move-sheet-right' && sheet >= instance.workbook.sheetCount - 1) {
        return t.sheetMoveAtBoundary;
      }
      if (action === 'default-width') return instance.i18n.strings.ribbon.disabled;
      if (action === 'tab-color-high-contrast') {
        return t.sheetActionUnavailable;
      }
      if (action === 'hide-sheet' && visibleSheetCount <= 1 && !isMacInstance(instance)) {
        return t.sheetHideRequiresVisibleSheet;
      }
      if (action === 'unhide-sheet' && hiddenSheetCount === 0) {
        return t.sheetUnhideRequiresHiddenSheet;
      }
      return undefined;
    };
    for (const button of menu.querySelectorAll<HTMLButtonElement>('[data-cell-format]')) {
      const action = button.dataset.cellFormat;
      const disabled =
        ((action === 'rename-sheet' || action === 'move-sheet-copy') && !canMoveSheet) ||
        action === 'default-width' ||
        (action === 'move-sheet-left' && (sheet <= 0 || !canMoveSheet)) ||
        (action === 'move-sheet-right' &&
          (sheet >= instance.workbook.sheetCount - 1 || !canMoveSheet)) ||
        (action === 'hide-sheet' && visibleSheetCount <= 1 && !isMacInstance(instance)) ||
        (action === 'unhide-sheet' && hiddenSheetCount === 0) ||
        action === 'tab-color-high-contrast';
      setMenuControlDisabled(button, disabled, reasonForFormatAction(action));
      if (action === 'lock-cell' || action === 'toggle-lock') {
        button.setAttribute('role', 'menuitemcheckbox');
        button.setAttribute('aria-checked', String(activeLocked));
        button.classList.toggle('fc-tb__menu-item--checked', activeLocked);
      }
      if (action?.startsWith('tab-color-')) {
        const active = action === activeTabColorAction;
        button.setAttribute('role', 'menuitemradio');
        button.setAttribute('aria-checked', String(active));
        button.classList.toggle('fc-tb__color-swatch--active', active);
      }
    }
  };

const updateCellInsertMenu =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['updateCellInsertMenu'] =>
  (menu) => {
    const structureProtected = isWorkbookStructureProtected(instance.store.getState());
    const disabledReason = instance.i18n.strings.ribbonMenu.workbookStructureProtectedBlocked;
    for (const button of menu.querySelectorAll<HTMLButtonElement>('[data-cell-insert]')) {
      const disabled = button.dataset.cellInsert === 'sheet' && structureProtected;
      setMenuControlDisabled(button, disabled, disabledReason);
    }
  };

const updateCellDeleteMenu =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['updateCellDeleteMenu'] =>
  (menu) => {
    const state = instance.store.getState();
    const structureProtected = isWorkbookStructureProtected(state);
    const canRemoveSheet = instance.workbook.capabilities.sheetMutate;
    const strings = instance.i18n.strings.ribbonMenu;
    for (const button of menu.querySelectorAll<HTMLButtonElement>('[data-cell-delete]')) {
      let disabledReason: string | undefined;
      if (button.dataset.cellDelete === 'sheet') {
        if (structureProtected) disabledReason = strings.workbookStructureProtectedBlocked;
        else if (instance.workbook.sheetCount <= 1 && !isMacInstance(instance)) {
          disabledReason = strings.sheetDeleteRequiresAnotherSheet;
        } else if (!canRemoveSheet) disabledReason = strings.sheetMutationUnavailable;
      }
      const lastSheet = instance.workbook.sheetCount <= 1 && !isMacInstance(instance);
      const disabled =
        button.dataset.cellDelete === 'sheet' &&
        (structureProtected || lastSheet || !canRemoveSheet);
      setMenuControlDisabled(button, disabled, disabledReason);
    }
  };

type CellsDropdownDefaults = Pick<
  DynamicDropdownsCtx,
  | 'applyCellInsertAction'
  | 'applyCellDeleteAction'
  | 'applyCellFormatAction'
  | 'updateFormatCellsMenu'
  | 'updateCellInsertMenu'
  | 'updateCellDeleteMenu'
>;

export function createCellsDropdownDefaults(
  instance: SpreadsheetInstance,
  opts: Pick<
    DefaultDynamicDropdownsOptions,
    'projectFormatToolbar' | 'refreshCells' | 'renderSheetTabs'
  >,
): CellsDropdownDefaults {
  return {
    applyCellInsertAction: buildCellInsertAction(instance),
    applyCellDeleteAction: buildCellDeleteAction(instance),
    applyCellFormatAction: buildCellFormatAction(instance, opts),
    updateFormatCellsMenu: updateFormatCellsMenu(instance),
    updateCellInsertMenu: updateCellInsertMenu(instance),
    updateCellDeleteMenu: updateCellDeleteMenu(instance),
  };
}
