import {
  protectedSheetPermissions,
  sheetHasPassword,
  verifySheetProtectionPassword,
} from '../../commands/protection.js';
import {
  type ProtectionActionContext,
  setSheetProtectionAction,
} from '../../commands/protection-actions.js';
import type { DialogEntry } from '../../interact/dialog-entry.js';
import type { SpreadsheetInstance } from '../../mount.js';
import { showMessage } from '../../toolbar/dialogs/prompt.js';
import {
  showProtectSheetDialog,
  showUnprotectSheetDialog,
} from '../../toolbar/dialogs/protection.js';

/** A protection command run from `entry` on the mounted instance. */
export const protectionContext = (
  instance: SpreadsheetInstance,
  entry: DialogEntry,
): ProtectionActionContext => ({
  store: instance.store,
  workbook: instance.workbook,
  history: instance.history,
  origin: entry.origin,
  commandId: entry.commandId,
});

const sheetOf = (instance: SpreadsheetInstance): number =>
  instance.store.getState().data.sheetIndex;

/** Ask for the password guarding a protected object; true when it is
 *  unguarded or the typed one verifies, with a message on a wrong password. */
export const confirmUnlock = async (
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

/** The Protect Sheet dialog for the active sheet, then the protection it sets. */
export async function protectSheetFlow(
  instance: SpreadsheetInstance,
  entry: DialogEntry,
): Promise<void> {
  const strings = instance.i18n.strings;
  const mac = instance.host.dataset.fcPlatform === 'mac' ? strings.mac.protectSheet : null;
  const protectionStrings = { ...strings.protection, ...mac } as typeof strings.protection & {
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
      heading: mac?.heading,
      note: mac?.note,
      password: protectionStrings.password,
      // The Mac password field carries no placeholder; (省略可) is in its label.
      passwordPlaceholder: mac ? '' : protectionStrings.passwordPlaceholder,
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
  setSheetProtectionAction(protectionContext(instance, entry), sheetOf(instance), true, {
    password: result.password,
    permissions: result.permissions,
  });
  instance.host.focus();
}

/** Unprotect the active sheet, asking for its password when it has one. */
export async function unprotectSheetFlow(
  instance: SpreadsheetInstance,
  entry: DialogEntry,
): Promise<void> {
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
    const matches = await verifySheetProtectionPassword(instance.store.getState(), sheet, password);
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
  setSheetProtectionAction(protectionContext(instance, entry), sheet, false);
  instance.host.focus();
}

/** Protect Sheet or Unprotect Sheet, whichever the active sheet needs. */
export const toggleSheetProtectionFlow = (
  instance: SpreadsheetInstance,
  entry: DialogEntry,
): Promise<void> =>
  instance.isSheetProtected()
    ? unprotectSheetFlow(instance, entry)
    : protectSheetFlow(instance, entry);
