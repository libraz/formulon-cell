import {
  addAllowedEditRange,
  isWorkbookStructureProtected,
  protectedSheetPassword,
  protectedSheetPasswordHash,
  protectedSheetPermissions,
  verifySheetProtectionPasswordHash,
} from '../../commands/protection.js';
import {
  commentAt,
  executeRibbonCommentAction,
  executeRibbonProtectionAction,
  formatA1Range,
  listComments,
  parseA1Range,
  type SpreadsheetInstance,
  setWorkbookStructureProtected,
} from '../../index.js';
import { showMessage } from '../../toolbar/dialogs/prompt.js';
import {
  showAllowEditRangeDialog,
  showProtectSheetDialog,
  showUnprotectSheetDialog,
} from '../../toolbar/dialogs/protection.js';
import type { DynamicDropdownsCtx } from '../../toolbar/ribbon/dynamic-dropdowns.js';
import { createCellsDropdownDefaults } from './cells.js';
import { setMenuControlDisabled, showInstanceReport } from './menu-feedback.js';
import { normalizedSelectionRange } from './selection.js';

const updateReviewCommentsMenu =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['updateReviewCommentsMenu'] =>
  (menu) => {
    const state = instance.store.getState();
    const hasActiveComment = commentAt(state, state.selection.active) !== null;
    const hasAnyComment = listComments(state, state.data.sheetIndex).length > 0;
    const strings = instance.i18n.strings.ribbonMenu;
    for (const button of menu.querySelectorAll<HTMLButtonElement>('[data-comment-action]')) {
      const action = button.dataset.commentAction;
      const disabled =
        (action === 'delete-active' && !hasActiveComment) ||
        (action === 'delete-all' && !hasAnyComment);
      const reason =
        action === 'delete-active' ? strings.commentDeleteRequiresActive : strings.commentNone;
      setMenuControlDisabled(button, disabled, reason);
    }
  };

const buildProtectAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyProtectAction'] =>
  async (action) => {
    if (action === 'protect-sheet') {
      const strings = instance.i18n.strings;
      const protectionStrings = strings.protection as typeof strings.protection & {
        allowAutoFilter: string;
        allowDeleteColumns: string;
        allowDeleteRows: string;
        allowEditObjects: string;
        allowEditScenarios: string;
        allowFormatCells: string;
        allowFormatColumns: string;
        allowFormatRows: string;
        allowInsertColumns: string;
        allowInsertHyperlinks: string;
        allowInsertRows: string;
        allowPivotTables: string;
        allowSelectLockedCells: string;
        allowSelectUnlockedCells: string;
        allowSort: string;
        allowUsersTo: string;
        confirmPassword: string;
        passwordMismatch: string;
      };
      const result = await showProtectSheetDialog({
        strings: {
          title: protectionStrings.protectSheet,
          password: protectionStrings.password,
          passwordPlaceholder: protectionStrings.passwordPlaceholder,
          confirmPassword: protectionStrings.confirmPassword,
          passwordMismatch: protectionStrings.passwordMismatch,
          allowLabel: protectionStrings.allowUsersTo,
          allowSelectLockedCells: protectionStrings.allowSelectLockedCells,
          allowSelectUnlockedCells: protectionStrings.allowSelectUnlockedCells,
          allowFormatCells: protectionStrings.allowFormatCells,
          allowFormatColumns: protectionStrings.allowFormatColumns,
          allowFormatRows: protectionStrings.allowFormatRows,
          allowInsertColumns: protectionStrings.allowInsertColumns,
          allowInsertRows: protectionStrings.allowInsertRows,
          allowInsertHyperlinks: protectionStrings.allowInsertHyperlinks,
          allowDeleteColumns: protectionStrings.allowDeleteColumns,
          allowDeleteRows: protectionStrings.allowDeleteRows,
          allowSort: protectionStrings.allowSort,
          allowAutoFilter: protectionStrings.allowAutoFilter,
          allowPivotTables: protectionStrings.allowPivotTables,
          allowEditObjects: protectionStrings.allowEditObjects,
          allowEditScenarios: protectionStrings.allowEditScenarios,
          ok: strings.pageSetup.ok,
          cancel: strings.pageSetup.cancel,
        },
        initial: protectedSheetPermissions(
          instance.store.getState(),
          instance.store.getState().data.sheetIndex,
        ),
      });
      if (!result) {
        instance.host.focus();
        return;
      }
      (
        instance.setSheetProtected as (
          on: boolean,
          password?: string,
          permissions?: import('../../store/store.js').SheetProtectionPermissions,
        ) => void
      )(true, result.password, result.permissions);
      instance.host.focus();
      return;
    }
    if (action === 'unprotect-sheet') {
      const strings = instance.i18n.strings;
      const sheet = instance.store.getState().data.sheetIndex;
      const currentPassword = protectedSheetPassword(instance.store.getState(), sheet);
      const currentPasswordHash = protectedSheetPasswordHash(instance.store.getState(), sheet);
      if (currentPassword || currentPasswordHash) {
        const password = await showUnprotectSheetDialog({
          title: strings.protection.unprotectSheet,
          password: strings.protection.password,
          ok: strings.pageSetup.ok,
          cancel: strings.pageSetup.cancel,
        });
        if (password === null) {
          instance.host.focus();
          return;
        }
        const matches =
          currentPassword !== undefined
            ? password === currentPassword
            : currentPasswordHash
              ? await verifySheetProtectionPasswordHash(password, currentPasswordHash)
              : false;
        if (!matches) {
          await showMessage({
            title: strings.protection.unprotectSheet,
            message: strings.ribbonMenu.workbookIncorrectPassword,
            okLabel: strings.pageSetup.ok,
          });
          instance.host.focus();
          return;
        }
      }
      instance.setSheetProtected(false);
      instance.host.focus();
      return;
    }
    if (action === 'lock-cell' || action === 'unlock-cell') {
      await createCellsDropdownDefaults(instance, {}).applyCellFormatAction(action);
      return;
    }
    if (action === 'protect-workbook' || action === 'unprotect-workbook') {
      setWorkbookStructureProtected(instance.store, action === 'protect-workbook');
      instance.host.focus();
      return;
    }
    if (action === 'allow-edit-ranges' || action === 'clear-allowed-edit-ranges') {
      const strings = instance.i18n.strings;
      if (action === 'allow-edit-ranges') {
        const selection = normalizedSelectionRange(instance);
        const sheetName = instance.workbook.sheetName(selection.sheet);
        const pivotStrings = strings.pivotTableDialog as typeof strings.pivotTableDialog & {
          rangePickerSelect: string;
        };
        const result = await showAllowEditRangeDialog({
          strings: {
            title: strings.ribbonMenu.allowEditRangesDialogTitle,
            range: strings.ribbonMenu.allowEditRangesDialogRange,
            invalid: strings.ribbonMenu.allowEditRangesDialogInvalid,
            rangePickerLabel: pivotStrings.rangePickerSelect,
            ok: strings.pageSetup.ok,
            cancel: strings.pageSetup.cancel,
          },
          initialRange: formatA1Range(selection),
          pickRange: () => formatA1Range(normalizedSelectionRange(instance)),
          validateRange: (value) => parseA1Range(value, selection.sheet, sheetName) !== null,
          subscribeToRangeChanges: (listener) => instance.store.subscribe(listener),
        });
        if (result === null) {
          instance.host.focus();
          return;
        }
        const range = parseA1Range(result, selection.sheet, sheetName) ?? selection;
        const rangeText = formatA1Range(range);
        addAllowedEditRange(instance.store, range, { title: rangeText });
        const report = {
          title: strings.ribbonMenu.allowEditRangesDialogTitle,
          items: [
            {
              severity: 'info' as const,
              label: strings.ribbonMenu.allowEditRangesCommand,
              detail: strings.ribbonMenu.allowedEditRangeAddedStatus.replace('{range}', rangeText),
            },
          ],
        };
        await showInstanceReport(instance, report.title, report.items);
        instance.host.focus();
        return;
      }
      const report = executeRibbonProtectionAction({
        store: instance.store,
        action: 'clear-allowed-edit-ranges',
        strings: {
          allowEditRangesDialogTitle: strings.ribbonMenu.allowEditRangesDialogTitle,
          allowEditRangesCommand: strings.ribbonMenu.allowEditRangesCommand,
          allowEditRangesClearCommand: strings.ribbonMenu.allowEditRangesClearCommand,
          allowedEditRangeAddedStatus: strings.ribbonMenu.allowedEditRangeAddedStatus,
          allowedEditRangesClearedStatus: strings.ribbonMenu.allowedEditRangesClearedStatus,
        },
      });
      await showInstanceReport(instance, report.title, report.items);
    }
  };

const updateProtectMenu =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['updateProtectMenu'] =>
  (menu) => {
    const state = instance.store.getState();
    const sheetProtected = state.protection.protectedSheets.has(state.data.sheetIndex);
    const workbookProtected = isWorkbookStructureProtected(state);
    const hasAllowedRanges = state.protection.allowedEditRanges.length > 0;
    const strings = instance.i18n.strings.ribbonMenu;
    for (const button of menu.querySelectorAll<HTMLButtonElement>('[data-protect-action]')) {
      const action = button.dataset.protectAction;
      const disabled =
        (action === 'protect-sheet' && sheetProtected) ||
        (action === 'unprotect-sheet' && !sheetProtected) ||
        (action === 'protect-workbook' && workbookProtected) ||
        (action === 'unprotect-workbook' && !workbookProtected) ||
        (action === 'clear-allowed-edit-ranges' && !hasAllowedRanges);
      const reason =
        action === 'protect-sheet'
          ? strings.protectSheetAlreadyProtected
          : action === 'unprotect-sheet'
            ? strings.unprotectSheetRequiresProtected
            : action === 'protect-workbook'
              ? strings.protectWorkbookAlreadyProtected
              : action === 'unprotect-workbook'
                ? strings.unprotectWorkbookRequiresProtected
                : action === 'clear-allowed-edit-ranges'
                  ? strings.allowEditRangesClearRequiresAny
                  : undefined;
      setMenuControlDisabled(button, disabled, reason);
    }
  };

const buildReviewCommentAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyReviewCommentAction'] =>
  (action) => {
    if (action !== 'delete-active' && action !== 'delete-all') return;
    executeRibbonCommentAction({
      store: instance.store,
      workbook: instance.workbook,
      history: instance.history,
      action,
    });
    instance.host.focus();
  };

type ReviewDropdownDefaults = Pick<
  DynamicDropdownsCtx,
  | 'applyReviewCommentAction'
  | 'updateReviewCommentsMenu'
  | 'applyProtectAction'
  | 'updateProtectMenu'
>;

export function createReviewDropdownDefaults(
  instance: SpreadsheetInstance,
): ReviewDropdownDefaults {
  return {
    applyReviewCommentAction: buildReviewCommentAction(instance),
    updateReviewCommentsMenu: updateReviewCommentsMenu(instance),
    applyProtectAction: buildProtectAction(instance),
    updateProtectMenu: updateProtectMenu(instance),
  };
}
