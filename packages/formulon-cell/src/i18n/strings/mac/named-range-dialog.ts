import type { FormulaDialogsStrings } from '../formula-dialogs/types.js';

type NamedRangeDialogStrings = FormulaDialogsStrings['namedRangeDialog'];

/** Labels only the Mac 名前の定義 (Define Name) dialog shows. */
export interface MacNamedRangeDialogLabels {
  defineTitle: string;
  /** Caption above the list of names. */
  defineNames: string;
  defineNamePrompt: string;
  defineRangePrompt: string;
}

/** Overrides of `namedRangeDialog`, plus the Mac-only labels, that apply on the Mac host only. */
export type MacNamedRangeDialogStrings = Partial<NamedRangeDialogStrings> &
  MacNamedRangeDialogLabels;

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
  labels: MacNamedRangeDialogLabels,
): MacNamedRangeDialogStrings => ({
  ...Object.fromEntries(MAC_NAMED_RANGE_DIALOG_KEYS.map((key) => [key, dialog[key]])),
  ...labels,
});

/** Japanese Mac app wording, measured from Excel for Mac. */
export const macNamedRangeDialogJa: MacNamedRangeDialogStrings = {
  title: 'ネーム マネージャー',
  newButton: '新規...',
  formulaHeader: '参照先',
  quickRefersLabel: '次を参照:',
  formulaLabel: '次を参照:',
  defineTitle: '名前の定義',
  defineNames: '名前:',
  defineNamePrompt: 'データ範囲の名前を入力してください:',
  defineRangePrompt: 'セルの範囲を選んでください:',
};
