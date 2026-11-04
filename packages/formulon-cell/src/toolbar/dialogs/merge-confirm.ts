import { mergeAcrossWillLoseData, mergeWillLoseData } from '../../commands/merge.js';
import { isMacWorkbook } from '../../engine/spreadsheet-profile.js';
import type { Range } from '../../engine/types.js';
import type { Strings } from '../../i18n/strings.js';
import type { State } from '../../store/store.js';
import { showConfirm } from './prompt.js';

/**
 * Ask the user before a merge that would discard data. Resolves `true` when the
 * merge may proceed — either because no data would be lost, or because the user
 * accepted the warning. Resolves `false` when the user cancels. A workbook on
 * the Mac profile gets the Mac desktop's wording.
 */
export async function confirmMergeLoseData(
  strings: Strings,
  state: State,
  range: Range,
  across = false,
  workbook: { spreadsheetProfileId?: () => string | null } | null = null,
): Promise<boolean> {
  if (!(across ? mergeAcrossWillLoseData(state, range) : mergeWillLoseData(state, range)))
    return true;
  const t = strings.ribbon;
  const mac = isMacWorkbook(workbook) ? strings.macAlerts.mergeLoseData : null;
  return showConfirm({
    title: t.mergeLoseDataTitle,
    message: mac?.message ?? t.mergeLoseDataMessage,
    okLabel: mac?.confirm ?? t.mergeLoseDataConfirm,
    cancelLabel: mac?.cancel ?? t.mergeLoseDataCancel,
  });
}
