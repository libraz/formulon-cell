import { mergeWillLoseData } from '../../commands/merge.js';
import type { Range } from '../../engine/types.js';
import type { Strings } from '../../i18n/strings.js';
import type { State } from '../../store/store.js';
import { showConfirm } from './prompt.js';

/**
 * Ask the user before a merge that would discard data. Resolves `true` when the
 * merge may proceed — either because no data would be lost, or because the user
 * accepted the warning. Resolves `false` when the user cancels.
 */
export async function confirmMergeLoseData(
  strings: Strings,
  state: State,
  range: Range,
): Promise<boolean> {
  if (!mergeWillLoseData(state, range)) return true;
  const t = strings.ribbon;
  return showConfirm({
    title: t.mergeLoseDataTitle,
    message: t.mergeLoseDataMessage,
    okLabel: t.mergeLoseDataConfirm,
    cancelLabel: t.mergeLoseDataCancel,
  });
}
