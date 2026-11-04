import type { FormulaDialogsStrings } from '../formula-dialogs/types.js';

type NamedRangeDialogStrings = FormulaDialogsStrings['namedRangeDialog'];

/** Overrides of `namedRangeDialog` that apply on the Mac host only. */
export type MacNamedRangeDialogStrings = Partial<NamedRangeDialogStrings>;

/** The `namedRangeDialog` keys the Mac Name Manager can override. A locale
 *  whose Mac wording is not measured repeats its own shared text. */
const MAC_NAMED_RANGE_DIALOG_KEYS = [
  'title',
  'newButton',
  'formulaHeader',
  'quickRefersLabel',
  'formulaLabel',
] as const satisfies readonly (keyof NamedRangeDialogStrings)[];

export const macNamedRangeDialogFrom = (
  dialog: NamedRangeDialogStrings,
): MacNamedRangeDialogStrings =>
  Object.fromEntries(MAC_NAMED_RANGE_DIALOG_KEYS.map((key) => [key, dialog[key]]));

/** Japanese Mac app wording, measured from Excel for Mac. */
export const macNamedRangeDialogJa: MacNamedRangeDialogStrings = {
  title: 'ネーム マネージャー',
  newButton: '新規...',
  formulaHeader: '参照先',
  quickRefersLabel: '次を参照:',
  formulaLabel: '次を参照:',
};
