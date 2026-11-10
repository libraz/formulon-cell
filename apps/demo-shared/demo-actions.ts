import type {
  CommandId,
  DialogId,
  DialogOptions,
  SpreadsheetInstance,
} from '@libraz/formulon-cell';
import { formatA1Cell } from '@libraz/formulon-cell';

/** Opens a dialog when the instance allows it now; a disabled feature or a policy refusal is a no-op. */
export const openDemoDialog = <K extends DialogId>(
  instance: SpreadsheetInstance | null | undefined,
  id: K,
  options?: DialogOptions[K],
): void => {
  if (instance?.dialogs.canOpen(id)) instance.dialogs.open(id, options);
};

/** Runs a ribbon command when the instance allows it now. */
export const runDemoCommand = (
  instance: SpreadsheetInstance | null | undefined,
  id: CommandId,
): void => {
  if (instance?.commands.can(id)) instance.commands.execute(id);
};

/** Draws precedent arrows for the active cell. */
export const traceDemoPrecedents = (instance: SpreadsheetInstance | null | undefined): void => {
  if (!instance) return;
  const { active } = instance.selection.get();
  instance.audit.tracePrecedents(formatA1Cell(active.row, active.col));
};

/** Protects the active sheet, or lifts the protection when it is already on. */
export const toggleDemoSheetProtection = (
  instance: SpreadsheetInstance | null | undefined,
): void => {
  if (!instance) return;
  const sheet = instance.workbook.sheet(instance.selection.get().sheet);
  void (sheet.getSheetProtection() ? sheet.unprotect() : sheet.protect());
};
