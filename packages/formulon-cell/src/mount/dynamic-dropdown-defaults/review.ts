import { canExecuteBuiltInRequest } from '../../commands/built-in-command-policy.js';
import { isMacChrome } from '../../commands/chrome-platform.js';
import {
  commentAt,
  listComments,
  listThreadedComments,
  noteVisibleAt,
} from '../../commands/comment.js';
import { convertNotesToThreads, workbookHasNotes } from '../../commands/convert-notes.js';
import { interactionControllerFor } from '../../commands/interaction-controller.js';
import {
  allowedEditRangeHasPassword,
  allowedEditRangesForSheet,
  isSheetProtected,
  isWorkbookStructureProtected,
  noteChangePermitted,
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
import {
  executeNoteVisibilityAction,
  executeRibbonCommentAction,
} from '../../commands/ribbon-comment.js';
import { executeRibbonProtectionAction } from '../../commands/ribbon-protection.js';
import { formatA1Cell, formatA1Range, MAX_COL } from '../../engine/address.js';
import {
  allowedEditRangeAreas,
  areasFromSqref,
  sqrefFromAreas,
} from '../../engine/protection-sync.js';
import { isMacHost } from '../../extensions/ui-options.js';
import { type DialogEntry, openCommentDialogFrom } from '../../interact/dialog-entry.js';
import { focusSheet } from '../../interact/dialog-focus.js';
import { isNavigationAddrAllowed } from '../../interact/navigation-policy.js';
import { createRangePickFormatter } from '../../interact/range-pick-format.js';
import type { SpreadsheetInstance } from '../../mount.js';
import { mutators } from '../../store/store.js';
import { showConfirm } from '../../toolbar/dialogs/prompt.js';
import {
  type AllowEditRangesManagerResult,
  showAllowEditRangeDialog,
  showAllowEditRangesManagerDialog,
  showProtectWorkbookDialog,
} from '../../toolbar/dialogs/protection.js';
import { projectPressedState } from '../../toolbar/ribbon/button.js';
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

/** Move to an adjacent note or thread in sheet order and open its editor. */
export const selectAdjacentComment = (
  instance: SpreadsheetInstance,
  direction: 1 | -1,
  mode: 'note' | 'thread' | 'all' = 'all',
): void => {
  const state = instance.store.getState();
  if (interactionControllerFor(instance.store)?.canSelect().allowed === false) return;
  const notes = [
    ...(mode === 'thread'
      ? []
      : listComments(state).map((entry) => ({ ...entry, mode: 'note' as const }))),
    ...(mode === 'note'
      ? []
      : (state.threadedComments.bySheet.get(state.data.sheetIndex) ?? [])
          .filter((entry) => !entry.parentId)
          .map((entry) => ({
            addr: { sheet: state.data.sheetIndex, row: entry.cell.row, col: entry.cell.col },
            mode: 'thread' as const,
          }))),
  ]
    .filter((entry) => isNavigationAddrAllowed(instance.store, entry.addr))
    .sort((a, b) => a.addr.row - b.addr.row || a.addr.col - b.addr.col);
  const active = state.selection.active;
  const rank = (addr: { row: number; col: number }): number => addr.row * (MAX_COL + 1) + addr.col;
  const next =
    direction === 1
      ? (notes.find((note) => rank(note.addr) > rank(active)) ?? notes[0])
      : (notes
          .slice()
          .reverse()
          .find((note) => rank(note.addr) < rank(active)) ?? notes.at(-1));
  if (!next) return;
  mutators.setActive(instance.store, next.addr);
  const selected = instance.store.getState().selection.active;
  if (
    selected.sheet !== next.addr.sheet ||
    selected.row !== next.addr.row ||
    selected.col !== next.addr.col
  )
    return;
  instance.openCommentDialog({ mode: next.mode });
};

/** Review > Notes: navigation needs a note on the sheet and Show/Hide one on the active cell; New
 *  Note reads Edit Note on a note cell and is absent while that note is locked; Convert to Comments
 *  needs a note anywhere in the book. The two Show toggles read as pressed while on. */
const updateReviewCommentsMenu =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['updateReviewCommentsMenu'] =>
  (menu) => {
    const state = instance.store.getState();
    const active = state.selection.active;
    const activeNote = commentAt(state, active) !== null;
    const anyNote = listComments(state, state.data.sheetIndex).length > 0;
    const strings = instance.i18n.strings;
    for (const button of menu.querySelectorAll<HTMLButtonElement>('[data-comment-action]')) {
      const action = button.dataset.commentAction;
      if (action === 'new-note') {
        const text = button.querySelector('.fc-tb__menu-item__text') ?? button;
        button.dataset.newNoteLabel ??= text.textContent ?? '';
        const label = activeNote ? strings.ribbon.editNote : button.dataset.newNoteLabel;
        text.textContent = label;
        button.setAttribute('aria-label', label);
        button.hidden = activeNote && !noteChangePermitted(state, active);
        continue;
      }
      if (action === 'toggle-note') projectPressedState(button, noteVisibleAt(state, active));
      if (action === 'toggle-all-notes') projectPressedState(button, state.ui.showAllNotes);
      const reason =
        action === 'convert-notes'
          ? workbookHasNotes(instance.workbook)
            ? null
            : strings.ribbonMenu.commentNone
          : action === 'toggle-all-notes'
            ? null
            : !anyNote
              ? strings.ribbonMenu.commentNone
              : action === 'toggle-note' && !activeNote
                ? strings.ribbonMenu.commentDeleteRequiresActive
                : null;
      setMenuControlDisabled(button, reason !== null, reason ?? undefined);
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
      focusSheet(instance.host);
      return;
    }
    if (action === 'allow-edit-ranges' || action === 'clear-allowed-edit-ranges') {
      if (!canExecuteBuiltInRequest(instance.store, 'protectionReview', 'ribbon').allowed) return;
      const strings = instance.i18n.strings;
      // Allowed ranges are fixed while the sheet is protected, as on desktop.
      if (isSheetProtected(instance.store.getState(), instance.store.getState().data.sheetIndex)) {
        focusSheet(instance.host);
        return;
      }
      if (action === 'allow-edit-ranges') {
        if ((await openAllowEditRangesManager(instance, entry)) === 'protectSheet') {
          await protectSheetFlow(instance, entry);
          return;
        }
        focusSheet(instance.host);
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
          focusSheet(instance.host);
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

/** The Review tab's comment and note actions: Delete Comment, and the Notes menu's items. */
const buildReviewCommentAction =
  (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyReviewCommentAction'] =>
  (action) => {
    const entry = { origin: 'ribbon', commandId: 'notesReview' } as const;
    switch (action) {
      case 'delete-active':
        executeRibbonCommentAction({
          store: instance.store,
          workbook: instance.workbook,
          history: instance.history,
          action,
          origin: 'ribbon',
          commandId: 'deleteCommentReview',
        });
        focusSheet(instance.host);
        return;
      case 'new-note':
        openCommentDialogFrom(instance.store, 'note', entry);
        return;
      case 'previous-note':
      case 'next-note':
        selectAdjacentComment(instance, action === 'next-note' ? 1 : -1, 'note');
        return;
      case 'toggle-note':
      case 'toggle-all-notes':
        executeNoteVisibilityAction({
          store: instance.store,
          workbook: instance.workbook,
          action: action === 'toggle-note' ? 'toggle-active' : 'toggle-all',
        });
        focusSheet(instance.host);
        return;
      case 'convert-notes':
        void convertNotesFrom(instance, 'notesReview');
        return;
    }
  };

/** Show Comments: the active sheet's threads, each under its cell, leaving out cells the
 *  navigation policy keeps out of reach. */
export const showCommentsList = (instance: SpreadsheetInstance, title: string): Promise<void> => {
  const sheet = instance.store.getState().data.sheetIndex;
  return showInstanceReport(
    instance,
    title,
    listThreadedComments(instance.store, sheet, instance.workbook)
      .filter((entry) =>
        isNavigationAddrAllowed(instance.store, {
          sheet,
          row: entry.cell.row,
          col: entry.cell.col,
        }),
      )
      .map((entry) => ({
        severity: 'info' as const,
        label: formatA1Cell(entry.cell.row, entry.cell.col),
        detail: entry.text,
      })),
  );
};

/** Convert to Comments: asks first, then turns every note of the workbook into a thread as one
 *  undo step. Does nothing when the workbook holds no note. */
export async function convertNotesFrom(
  instance: SpreadsheetInstance,
  commandId: string,
): Promise<void> {
  if (!workbookHasNotes(instance.workbook)) return;
  const t = instance.i18n.strings.ribbon;
  const confirmed = await showConfirm({
    title: t.convertNotes,
    message: t.convertNotesMessage,
    okLabel: t.convertNotesConfirm,
    cancelLabel: instance.i18n.strings.pageSetup.cancel,
    alert: isMacHost(instance.host),
  });
  if (confirmed) {
    convertNotesToThreads({
      store: instance.store,
      workbook: instance.workbook,
      history: instance.history,
      origin: 'ribbon',
      commandId,
    });
  }
  focusSheet(instance.host);
}

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
