import { isMacWorkbook } from '../../engine/spreadsheet-profile.js';
import type { WorkbookHandle } from '../../engine/workbook-handle.js';
import type { Strings } from '../../i18n/strings.js';
import { showConfirm } from './prompt.js';

/**
 * Ask before deleting `sheet`. The Mac desktop warns when the sheet holds
 * data; an empty sheet, or a workbook on another profile, resolves `true`
 * at once. Every sheet-delete entry point goes through this one confirm.
 */
export async function confirmSheetDelete(
  strings: Strings,
  workbook: Pick<WorkbookHandle, 'cells' | 'spreadsheetProfileId'>,
  sheet: number,
): Promise<boolean> {
  if (!isMacWorkbook(workbook)) return true;
  if (workbook.cells(sheet).next().done) return true;
  const t = strings.mac.alerts.deleteSheet;
  return showConfirm({
    title: strings.sheetTabs.deleteSheet,
    message: t.message,
    okLabel: t.confirm,
    cancelLabel: t.cancel,
    alert: true,
  });
}
