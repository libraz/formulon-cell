import { isMacWorkbook } from '../../engine/spreadsheet-profile.js';
import type { Strings } from '../../i18n/strings.js';
import { showConfirm } from './prompt.js';

/** Ask before a drag-move overwrites data. A workbook on the Mac profile gets
 *  the Mac desktop's wording. Resolves true to go ahead. */
export function confirmMoveOverwrite(
  strings: Strings,
  workbook: { spreadsheetProfileId?: () => string | null } | null,
): Promise<boolean> {
  const t = strings.moveCells;
  const mac = isMacWorkbook(workbook) ? strings.mac.alerts.overwriteOnMove : null;
  return showConfirm({
    title: t.title,
    message: mac?.message ?? t.overwriteMessage,
    okLabel: mac?.confirm ?? t.ok,
    cancelLabel: mac?.cancel ?? t.cancel,
  });
}
