import { commentAt, listComments } from '../../commands/comment.js';
import {
  addAllowedEditRange,
  allowedEditRangeHasPassword,
  allowedEditRangesForSheet,
  isSheetProtected,
  isWorkbookStructureProtected,
  protectedSheetPermissions,
  removeAllowedEditRange,
  setWorkbookStructureProtected,
  sheetHasPassword,
  updateAllowedEditRange,
  verifyAllowedEditRangePassword,
  verifySheetProtectionPassword,
  verifyWorkbookStructurePassword,
  workbookStructureHasPassword,
} from '../../commands/protection.js';
import { executeRibbonCommentAction } from '../../commands/ribbon-comment.js';
import { executeRibbonProtectionAction } from '../../commands/ribbon-protection.js';
import { formatA1Range } from '../../engine/address.js';
import {
  allowedEditRangeAreas,
  areasFromSqref,
  sqrefFromAreas,
} from '../../engine/protection-sync.js';
import type { SpreadsheetInstance } from '../../mount.js';
import { showMessage } from '../../toolbar/dialogs/prompt.js';
import {
  showAllowEditRangeDialog,
  showAllowEditRangesManagerDialog,
  showProtectSheetDialog,
  showProtectWorkbookDialog,
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
    const threads = state.threadedComments.bySheet.get(state.data.sheetIndex) ?? [];
    const hasActiveComment =
      commentAt(state, state.selection.active) !== null ||
      threads.some(
        (entry) =>
          !entry.parentId &&
          entry.row === state.selection.active.row &&
          entry.col === state.selection.active.col,
      );
    const hasAnyComment =
      listComments(state, state.data.sheetIndex).length > 0 || threads.length > 0;
    const strings = instance.i18n.strings.ribbonMenu;
    for (const button of menu.querySelectorAll<HTMLButtonElement>('[data-comment-action]')) {
      const action = button.dataset.commentAction;
      const disabled =
        (action === 'thread' && !instance.workbook.capabilities.threadedComments) ||
        (action === 'delete-active' && !hasActiveComment) ||
        (action === 'delete-all' && !hasAnyComment);
      const reason =
        action === 'delete-active' ? strings.commentDeleteRequiresActive : strings.commentNone;
      setMenuControlDisabled(button, disabled, reason);
    }
  };

/** Ask for the password guarding a protected object; true when it is
 *  unguarded or the typed one verifies, with a message on a wrong password. */
const confirmUnlock = async (
  instance: SpreadsheetInstance,
  title: string,
  guarded: boolean,
  verify: (password: string) => Promise<boolean>,
): Promise<boolean> => {
  if (!guarded) return true;
  const strings = instance.i18n.strings;
  const password = await showUnprotectSheetDialog({
    title,
    password: strings.protection.password,
    ok: strings.pageSetup.ok,
    cancel: strings.pageSetup.cancel,
  });
  if (password === null) return false;
  if (await verify(password)) return true;
  await showMessage({
    title,
    message: strings.ribbonMenu.workbookIncorrectPassword,
    okLabel: strings.pageSetup.ok,
  });
  return false;
};

/** Whether the password of the allowed edit range `id` was entered (or none is set). */
const unlockAllowedRange = (instance: SpreadsheetInstance, id: string): Promise<boolean> =>
  confirmUnlock(
    instance,
    instance.i18n.strings.ribbonMenu.allowEditRangesDialogTitle,
    allowedEditRangeHasPassword(instance.store.getState(), id),
    (password) => verifyAllowedEditRangePassword(instance.store.getState(), id, password),
  );

/** The Allow Users to Edit Ranges list for the active sheet, with New, Modify
 *  and Delete. Changing or deleting a password-guarded range asks for its password. */
const openAllowEditRangesManager = async (instance: SpreadsheetInstance): Promise<void> => {
  const strings = instance.i18n.strings;
  const menu = strings.ribbonMenu;
  const sheet = (): number => instance.store.getState().data.sheetIndex;
  const edit = async (id: string | null): Promise<void> => {
    const state = instance.store.getState();
    const current = id ? state.protection.allowedEditRanges.find((e) => e.id === id) : undefined;
    const selection = normalizedSelectionRange(instance);
    const parseAreas = (text: string) => areasFromSqref(text.replace(/,/g, ' '), sheet());
    const pivotStrings = strings.pivotTableDialog as typeof strings.pivotTableDialog & {
      rangePickerSelect: string;
    };
    const result = await showAllowEditRangeDialog({
      strings: {
        title: menu.allowEditRangesDialogTitle,
        range: menu.allowEditRangesDialogRange,
        invalid: menu.allowEditRangesDialogInvalid,
        rangePickerLabel: pivotStrings.rangePickerSelect,
        titleLabel: menu.allowEditRangesDialogTitleLabel,
        password: strings.protection.password,
        passwordPlaceholder: strings.protection.passwordPlaceholder,
        confirmPassword: strings.protection.confirmPassword,
        passwordMismatch: strings.protection.passwordMismatch,
        ok: strings.pageSetup.ok,
        cancel: strings.pageSetup.cancel,
      },
      initialRange: current
        ? sqrefFromAreas(allowedEditRangeAreas(current))
        : formatA1Range(selection),
      initialTitle: current?.title,
      hasPassword: !!current && allowedEditRangeHasPassword(state, current.id),
      pickRange: () => formatA1Range(normalizedSelectionRange(instance)),
      validateRange: (value) => parseAreas(value) !== null,
      subscribeToRangeChanges: (listener) => instance.store.subscribe(listener),
    });
    if (result === null) return;
    const [range = selection, ...extraAreas] = parseAreas(result.range) ?? [];
    const workbook = instance.workbook;
    if (current) {
      updateAllowedEditRange(
        instance.store,
        current.id,
        {
          title: result.title,
          range,
          extraAreas,
          ...(result.password !== undefined ? { password: result.password } : {}),
        },
        { workbook },
      );
    } else {
      addAllowedEditRange(instance.store, range, {
        title: result.title || sqrefFromAreas([range, ...extraAreas]),
        extraAreas,
        password: result.password ?? undefined,
        workbook,
      });
    }
  };
  const guard = (id: string): Promise<boolean> => unlockAllowedRange(instance, id);
  await showAllowEditRangesManagerDialog({
    strings: {
      title: menu.allowEditRangesDialogTitle,
      newButton: menu.allowEditRangesNewButton,
      modifyButton: menu.allowEditRangesModifyButton,
      deleteButton: menu.allowEditRangesDeleteButton,
      close: strings.pageSetup.ok,
      empty: menu.allowEditRangesEmpty,
    },
    rows: () =>
      allowedEditRangesForSheet(instance.store.getState(), sheet()).map((entry) => ({
        id: entry.id,
        title: entry.title,
        ref: sqrefFromAreas(allowedEditRangeAreas(entry)),
      })),
    onNew: () => edit(null),
    onModify: async (id) => {
      if (await guard(id)) await edit(id);
    },
    onDelete: async (id) => {
      if (await guard(id)) {
        removeAllowedEditRange(instance.store, id, { workbook: instance.workbook });
      }
    },
  });
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
      if (sheetHasPassword(instance.store.getState(), sheet)) {
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
        const matches = await verifySheetProtectionPassword(
          instance.store.getState(),
          sheet,
          password,
        );
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
      const strings = instance.i18n.strings;
      const workbook = instance.workbook;
      if (action === 'protect-workbook') {
        const result = await showProtectWorkbookDialog({
          title: strings.ribbonMenu.protectWorkbookDialogTitle,
          password: strings.protection.password,
          passwordPlaceholder: strings.protection.passwordPlaceholder,
          confirmPassword: strings.protection.confirmPassword,
          passwordMismatch: strings.protection.passwordMismatch,
          ok: strings.pageSetup.ok,
          cancel: strings.pageSetup.cancel,
        });
        if (result) {
          setWorkbookStructureProtected(instance.store, true, {
            password: result.password,
            workbook,
          });
        }
      } else if (
        await confirmUnlock(
          instance,
          strings.ribbonMenu.unprotectWorkbookDialogTitle,
          workbookStructureHasPassword(instance.store.getState()),
          (password) => verifyWorkbookStructurePassword(instance.store.getState(), password),
        )
      ) {
        setWorkbookStructureProtected(instance.store, false, { workbook });
      }
      instance.host.focus();
      return;
    }
    if (action === 'allow-edit-ranges' || action === 'clear-allowed-edit-ranges') {
      const strings = instance.i18n.strings;
      // Allowed ranges are fixed while the sheet is protected, as on desktop.
      if (isSheetProtected(instance.store.getState(), instance.store.getState().data.sheetIndex)) {
        instance.host.focus();
        return;
      }
      if (action === 'allow-edit-ranges') {
        await openAllowEditRangesManager(instance);
        instance.host.focus();
        return;
      }
      // Clearing every range asks for each guarded range's password first.
      const sheetRanges = allowedEditRangesForSheet(
        instance.store.getState(),
        instance.store.getState().data.sheetIndex,
      );
      for (const entry of sheetRanges) {
        const guarded = allowedEditRangeHasPassword(instance.store.getState(), entry.id);
        if (guarded && !(await unlockAllowedRange(instance, entry.id))) {
          instance.host.focus();
          return;
        }
      }
      const report = executeRibbonProtectionAction({
        store: instance.store,
        workbook: instance.workbook,
        action: 'clear-allowed-edit-ranges',
        strings: {
          allowEditRangesDialogTitle: strings.ribbonMenu.allowEditRangesDialogTitle,
          allowEditRangesCommand: strings.ribbonMenu.allowEditRangesCommand,
          allowEditRangesClearCommand: strings.ribbonMenu.allowEditRangesClearCommand,
          allowedEditRangeAddedStatus: strings.ribbonMenu.allowedEditRangeAddedStatus,
          allowedEditRangesClearedStatus: strings.ribbonMenu.allowedEditRangesClearedStatus,
          allowEditRangesRequiresUnprotected: strings.ribbonMenu.allowEditRangesRequiresUnprotected,
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
    const hasAllowedRanges = allowedEditRangesForSheet(state, state.data.sheetIndex).length > 0;
    const strings = instance.i18n.strings.ribbonMenu;
    for (const button of menu.querySelectorAll<HTMLButtonElement>('[data-protect-action]')) {
      const action = button.dataset.protectAction;
      const disabled =
        (action === 'protect-sheet' && sheetProtected) ||
        (action === 'unprotect-sheet' && !sheetProtected) ||
        (action === 'protect-workbook' && workbookProtected) ||
        (action === 'unprotect-workbook' && !workbookProtected) ||
        ((action === 'allow-edit-ranges' || action === 'clear-allowed-edit-ranges') &&
          sheetProtected) ||
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
                : (action === 'allow-edit-ranges' || action === 'clear-allowed-edit-ranges') &&
                    sheetProtected
                  ? strings.allowEditRangesRequiresUnprotected
                  : action === 'clear-allowed-edit-ranges'
                    ? strings.allowEditRangesClearRequiresAny
                    : undefined;
      setMenuControlDisabled(button, disabled, reason);
    }
  };

const buildReviewCommentAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyReviewCommentAction'] =>
  (action) => {
    if (action === 'thread') {
      if (instance.workbook.capabilities.threadedComments)
        instance.openCommentDialog({ mode: 'thread' });
      return;
    }
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
