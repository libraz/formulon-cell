import { canExecuteBuiltInRequest } from '../../commands/built-in-command-policy.js';
import { isMacChrome } from '../../commands/chrome-platform.js';
import { commentAt, listComments } from '../../commands/comment.js';
import {
  allowedEditRangeHasPassword,
  allowedEditRangesForSheet,
  isSheetProtected,
  isWorkbookStructureProtected,
  protectionCheckpoint,
  verifyAllowedEditRangePassword,
  verifyWorkbookStructurePassword,
  workbookStructureHasPassword,
} from '../../commands/protection.js';
import {
  addAllowedEditRangeAction,
  removeAllowedEditRangeAction,
  setWorkbookProtectionAction,
  updateAllowedEditRangeAction,
} from '../../commands/protection-actions.js';
import { executeRibbonCommentAction } from '../../commands/ribbon-comment.js';
import { executeRibbonProtectionAction } from '../../commands/ribbon-protection.js';
import { formatA1Range } from '../../engine/address.js';
import {
  allowedEditRangeAreas,
  areasFromSqref,
  sqrefFromAreas,
} from '../../engine/protection-sync.js';
import { type DialogEntry, openCommentDialogFrom } from '../../interact/dialog-entry.js';
import { createRangePickFormatter } from '../../interact/range-pick-format.js';
import type { SpreadsheetInstance } from '../../mount.js';
import {
  type AllowEditRangesManagerResult,
  showAllowEditRangeDialog,
  showAllowEditRangesManagerDialog,
  showProtectWorkbookDialog,
} from '../../toolbar/dialogs/protection.js';
import type { DynamicDropdownsCtx } from '../../toolbar/ribbon/dynamic-dropdowns.js';
import { createCellsDropdownDefaults } from './cells.js';
import { setMenuControlDisabled, showInstanceReport } from './menu-feedback.js';
import {
  confirmUnlock,
  protectionContext,
  protectSheetFlow,
  unprotectSheetFlow,
} from './protection-flows.js';
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
          entry.cell.row === state.selection.active.row &&
          entry.cell.col === state.selection.active.col,
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

/** Whether the password of the allowed edit range `id` was entered (or none is set). */
const unlockAllowedRange = (instance: SpreadsheetInstance, id: string): Promise<boolean> =>
  confirmUnlock(
    instance,
    instance.i18n.strings.ribbonMenu.allowEditRangesDialogTitle,
    allowedEditRangeHasPassword(instance.store.getState(), id),
    (password) => verifyAllowedEditRangePassword(instance.store.getState(), id, password),
  );

/** The Allow Users to Edit Ranges list for the active sheet, with New, Modify
 *  and Delete. Changing or deleting a password-guarded range asks for its password.
 *  Cancel puts back the ranges as of opening or the last Apply. The Mac desktop has no manager. */
const openAllowEditRangesManager = async (
  instance: SpreadsheetInstance,
  entry: DialogEntry,
): Promise<AllowEditRangesManagerResult> => {
  const ctx = protectionContext(instance, entry);
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
    // Protected areas live on the dialog's sheet, so picks carry no sheet prefix.
    const pickFormat = createRangePickFormatter({ store: instance.store, sheetName: () => '' });
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
      pickRange: () => pickFormat.range({ prefix: 'never' }),
      validateRange: (value) => parseAreas(value) !== null,
      subscribeToRangeChanges: pickFormat.subscribe,
    });
    if (result === null) return;
    const [range = selection, ...extraAreas] = parseAreas(result.range) ?? [];
    if (current) {
      updateAllowedEditRangeAction(ctx, current.id, {
        title: result.title,
        range,
        extraAreas,
        ...(result.password !== undefined ? { password: result.password } : {}),
      });
    } else {
      addAllowedEditRangeAction(ctx, range, {
        title: result.title || sqrefFromAreas([range, ...extraAreas]),
        extraAreas,
        password: result.password ?? undefined,
      });
    }
  };
  const guard = (id: string): Promise<boolean> => unlockAllowedRange(instance, id);
  let restore = protectionCheckpoint(instance.store, instance.workbook);
  const result = await showAllowEditRangesManagerDialog({
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
        removeAllowedEditRangeAction(ctx, id);
      }
    },
    commit: {
      strings: {
        permissions: menu.allowEditRangesPermissionsButton,
        protectSheet: menu.allowEditRangesProtectSheetButton,
        cancel: strings.pageSetup.cancel,
        apply: menu.allowEditRangesApplyButton,
      },
      onApply: () => {
        restore = protectionCheckpoint(instance.store, instance.workbook);
      },
    },
  });
  if (result === 'discard') restore();
  return result;
};

/** The command id a protection menu action runs under when its caller names none. */
const PROTECT_ACTION_COMMAND: Readonly<Record<string, string>> = {
  'protect-sheet': 'protectReview',
  'unprotect-sheet': 'protectReview',
  'protect-workbook': 'protectWorkbookReview',
  'unprotect-workbook': 'protectWorkbookReview',
  'allow-edit-ranges': 'protectionReview',
  'clear-allowed-edit-ranges': 'protectionReview',
};

const protectActionEntry = (action: string): DialogEntry => ({
  origin: 'ribbon',
  commandId: PROTECT_ACTION_COMMAND[action] ?? 'protectReview',
});

const buildProtectAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyProtectAction'] =>
  async (action, entry = protectActionEntry(action)) => {
    if (action === 'protect-sheet') {
      await protectSheetFlow(instance, entry);
      return;
    }
    if (action === 'unprotect-sheet') {
      await unprotectSheetFlow(instance, entry);
      return;
    }
    if (action === 'lock-cell' || action === 'unlock-cell') {
      await createCellsDropdownDefaults(instance, {}).applyCellFormatAction(action);
      return;
    }
    if (action === 'protect-workbook' || action === 'unprotect-workbook') {
      const strings = instance.i18n.strings;
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
        if (result)
          setWorkbookProtectionAction(protectionContext(instance, entry), true, result.password);
      } else if (
        await confirmUnlock(
          instance,
          strings.ribbonMenu.unprotectWorkbookDialogTitle,
          workbookStructureHasPassword(instance.store.getState()),
          (password) => verifyWorkbookStructurePassword(instance.store.getState(), password),
        )
      ) {
        setWorkbookProtectionAction(protectionContext(instance, entry), false);
      }
      instance.host.focus();
      return;
    }
    if (action === 'allow-edit-ranges' || action === 'clear-allowed-edit-ranges') {
      if (!canExecuteBuiltInRequest(instance.store, 'protectionReview', 'ribbon').allowed) return;
      const strings = instance.i18n.strings;
      // Allowed ranges are fixed while the sheet is protected, as on desktop.
      if (isSheetProtected(instance.store.getState(), instance.store.getState().data.sheetIndex)) {
        instance.host.focus();
        return;
      }
      if (action === 'allow-edit-ranges') {
        if ((await openAllowEditRangesManager(instance, entry)) === 'protectSheet') {
          await protectSheetFlow(instance, entry);
          return;
        }
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
        context: protectionContext(instance, entry),
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
      const commandId = action ? PROTECT_ACTION_COMMAND[action] : undefined;
      const denied =
        commandId !== undefined &&
        !canExecuteBuiltInRequest(instance.store, commandId, 'ribbon').allowed;
      const disabled =
        denied ||
        (action === 'protect-sheet' && sheetProtected) ||
        (action === 'unprotect-sheet' && !sheetProtected) ||
        (action === 'protect-workbook' && workbookProtected) ||
        (action === 'unprotect-workbook' && !workbookProtected) ||
        ((action === 'allow-edit-ranges' || action === 'clear-allowed-edit-ranges') &&
          sheetProtected) ||
        (action === 'clear-allowed-edit-ranges' && !hasAllowedRanges);
      const reason = denied
        ? instance.i18n.strings.backstage.commandUnavailable
        : action === 'protect-sheet'
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
      // The Mac desktop has no Allow Edit Ranges: its items are not offered at all.
      if (action === 'allow-edit-ranges' || action === 'clear-allowed-edit-ranges')
        button.hidden = isMacChrome(instance.store);
    }
  };

const buildReviewCommentAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyReviewCommentAction'] =>
  (action) => {
    if (action === 'thread') {
      if (instance.workbook.capabilities.threadedComments)
        openCommentDialogFrom(instance.store, 'thread', {
          origin: 'ribbon',
          commandId: 'newCommentReview',
        });
      return;
    }
    if (action !== 'delete-active' && action !== 'delete-all') return;
    executeRibbonCommentAction({
      store: instance.store,
      workbook: instance.workbook,
      history: instance.history,
      action,
      origin: 'ribbon',
      commandId: 'deleteCommentReview',
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
