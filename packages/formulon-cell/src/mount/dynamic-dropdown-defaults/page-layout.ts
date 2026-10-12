import type { OperationIntent } from '../../commands/interaction-policy.js';
import { workbookThemeMatchesPreset } from '../../engine/workbook-theme-presets.js';
import {
  addPrintArea,
  clearPrintArea,
  clearSheetBackgroundImage,
  formatA1Range,
  getPageSetup,
  insertManualPageBreak,
  removeManualPageBreak,
  resetManualPageBreaks,
  type SpreadsheetInstance,
  setPrintArea,
  setSheetBackgroundImage,
  type ThemeName,
} from '../../index.js';
import { pickImageFileDataUrl } from '../../toolbar/dialogs/image-file.js';
import type { DynamicDropdownsCtx } from '../../toolbar/ribbon/dynamic-dropdowns.js';
import { applyWorkbookTheme } from '../workbook-theme-action.js';
import { setMenuControlDisabled } from './menu-feedback.js';
import { normalizedSelectionRange } from './selection.js';

const buildUiTheme =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyUiTheme'] =>
  (theme) => {
    instance.setTheme(theme);
  };

const buildWorkbookTheme =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyWorkbookTheme'] =>
  (presetId) => {
    applyWorkbookTheme(instance, presetId);
    instance.host.focus();
  };

/** Normalize the grid's stored theme to the `ThemeName` vocabulary the
 *  page-theme tiles carry in `data-page-theme-action`, so the active tile is
 *  highlighted. Anything unrecognized falls back to the default `paper`. */
const currentPageThemeAction = (theme: string | undefined): ThemeName => {
  if (theme === 'ink') return 'ink';
  if (theme === 'contrast') return 'contrast';
  return 'paper';
};

const updatePageThemeMenu =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['updatePageThemeMenu'] =>
  (menu) => {
    const current = currentPageThemeAction(instance.store.getState().ui.theme);
    for (const button of menu.querySelectorAll<HTMLButtonElement>('[data-page-theme-action]')) {
      const active = button.dataset.pageThemeAction === current;
      button.setAttribute('role', 'menuitemradio');
      button.setAttribute('aria-checked', String(active));
      button.classList.toggle('fc-tb__visual-tile--active', active);
    }
    const workbookTheme = instance.workbook.capabilities.workbookTheme
      ? instance.workbook.getTheme()
      : null;
    const workbookThemeIntent: OperationIntent = {
      operation: 'format',
      origin: 'ribbon',
      commandId: 'workbookTheme',
      effects: [{ kind: 'workbook' }],
    };
    const workbookFormatAllowed =
      instance.commands?.canExecute(workbookThemeIntent).allowed ?? true;
    const workbookThemeUnavailable =
      !workbookTheme ||
      workbookTheme.source === 2 ||
      instance.store.getState().protection.protectedSheets.size > 0 ||
      !workbookFormatAllowed;
    const unavailableReason = instance.i18n.strings.ribbonMenu.themeWorkbookUnavailable;
    for (const button of menu.querySelectorAll<HTMLButtonElement>('[data-workbook-theme-action]')) {
      const action = button.dataset.workbookThemeAction;
      const active =
        !workbookThemeUnavailable &&
        (action === 'office' || action === 'gray') &&
        workbookThemeMatchesPreset(instance.workbook, action);
      button.setAttribute('role', 'menuitemradio');
      button.setAttribute('aria-checked', String(active));
      button.classList.toggle('fc-tb__visual-tile--active', active);
      setMenuControlDisabled(button, workbookThemeUnavailable, unavailableReason);
    }
  };

const buildPrintAreaAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyPrintAreaAction'] =>
  (action) => {
    const sheet = instance.store.getState().data.sheetIndex;
    if (action === 'clear') clearPrintArea(instance.store, sheet, instance.history);
    else if (action === 'add')
      addPrintArea(
        instance.store,
        sheet,
        formatA1Range(normalizedSelectionRange(instance)),
        instance.history,
      );
    else
      setPrintArea(
        instance.store,
        sheet,
        formatA1Range(normalizedSelectionRange(instance)),
        instance.history,
      );
    instance.host.focus();
  };

const updatePrintAreaMenu =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['updatePrintAreaMenu'] =>
  (menu) => {
    const sheet = instance.store.getState().data.sheetIndex;
    const hasPrintArea = !!getPageSetup(instance.store.getState(), sheet).printArea?.trim();
    const disabledReason = instance.i18n.strings.ribbonMenu.printAreaRequiresExisting;
    for (const button of menu.querySelectorAll<HTMLButtonElement>('[data-print-area-action]')) {
      const action = button.dataset.printAreaAction;
      const disabled = (action === 'add' || action === 'clear') && !hasPrintArea;
      setMenuControlDisabled(button, disabled, disabledReason);
    }
  };

const buildPageBreakAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyPageBreakAction'] =>
  (action) => {
    const range = normalizedSelectionRange(instance);
    if (action === 'insert') {
      if (range.r0 > 0)
        insertManualPageBreak(instance.store, range.sheet, 'row', range.r0, instance.history);
      if (range.c0 > 0)
        insertManualPageBreak(instance.store, range.sheet, 'col', range.c0, instance.history);
    } else if (action === 'insert-row')
      insertManualPageBreak(instance.store, range.sheet, 'row', range.r0, instance.history);
    else if (action === 'insert-col')
      insertManualPageBreak(instance.store, range.sheet, 'col', range.c0, instance.history);
    else if (action === 'remove') {
      removeManualPageBreak(instance.store, range.sheet, 'row', range.r0, instance.history);
      removeManualPageBreak(instance.store, range.sheet, 'col', range.c0, instance.history);
    } else if (action === 'remove-row')
      removeManualPageBreak(instance.store, range.sheet, 'row', range.r0, instance.history);
    else if (action === 'remove-col')
      removeManualPageBreak(instance.store, range.sheet, 'col', range.c0, instance.history);
    else if (action === 'reset-all')
      resetManualPageBreaks(instance.store, range.sheet, instance.history);
    instance.host.focus();
  };

const updatePageBreaksMenu =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['updatePageBreaksMenu'] =>
  (menu) => {
    const range = normalizedSelectionRange(instance);
    const setup = getPageSetup(instance.store.getState(), range.sheet);
    const rowBreaks = new Set(setup.manualPageBreakRows ?? []);
    const colBreaks = new Set(setup.manualPageBreakCols ?? []);
    const hasAnyBreak = rowBreaks.size > 0 || colBreaks.size > 0;
    const canInsertAtSelection = range.r0 > 0 || range.c0 > 0;
    const hasBreakAtSelection = rowBreaks.has(range.r0) || colBreaks.has(range.c0);
    const strings = instance.i18n.strings.ribbonMenu;
    for (const button of menu.querySelectorAll<HTMLButtonElement>('[data-page-break-action]')) {
      const action = button.dataset.pageBreakAction;
      const disabled =
        (action === 'insert' && !canInsertAtSelection) ||
        (action === 'insert-row' && range.r0 <= 0) ||
        (action === 'insert-col' && range.c0 <= 0) ||
        (action === 'remove' && !hasBreakAtSelection) ||
        (action === 'remove-row' && !rowBreaks.has(range.r0)) ||
        (action === 'remove-col' && !colBreaks.has(range.c0)) ||
        (action === 'reset-all' && !hasAnyBreak);
      const reason = action?.startsWith('insert')
        ? strings.pageBreakInsertRequiresSelection
        : action === 'reset-all'
          ? strings.pageBreakResetRequiresAny
          : strings.pageBreakRemoveRequiresBreak;
      setMenuControlDisabled(button, disabled, reason);
    }
  };

const buildSheetBackgroundAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applySheetBackgroundAction'] =>
  async (action) => {
    const sheet = instance.store.getState().data.sheetIndex;
    if (action === 'clear') {
      clearSheetBackgroundImage(instance.store, sheet, instance.history);
      instance.host.focus();
      return;
    }
    const picked = await pickImageFileDataUrl();
    if (!picked) {
      instance.host.focus();
      return;
    }
    setSheetBackgroundImage(instance.store, sheet, picked.src, instance.history);
    instance.host.focus();
  };

type PageLayoutDropdownDefaults = Pick<
  DynamicDropdownsCtx,
  | 'applyUiTheme'
  | 'applyWorkbookTheme'
  | 'updatePageThemeMenu'
  | 'applyPrintAreaAction'
  | 'updatePrintAreaMenu'
  | 'applyPageBreakAction'
  | 'updatePageBreaksMenu'
  | 'applySheetBackgroundAction'
>;

export function createPageLayoutDropdownDefaults(
  instance: SpreadsheetInstance,
): PageLayoutDropdownDefaults {
  return {
    applyUiTheme: buildUiTheme(instance),
    applyWorkbookTheme: buildWorkbookTheme(instance),
    updatePageThemeMenu: updatePageThemeMenu(instance),
    applyPrintAreaAction: buildPrintAreaAction(instance),
    updatePrintAreaMenu: updatePrintAreaMenu(instance),
    applyPageBreakAction: buildPageBreakAction(instance),
    updatePageBreaksMenu: updatePageBreaksMenu(instance),
    applySheetBackgroundAction: buildSheetBackgroundAction(instance),
  };
}
