import type { FormulaDialogsStrings } from '../formula-dialogs/types.js';

type IterativeDialogStrings = FormulaDialogsStrings['iterativeDialog'];

/** Labels the Mac Calculation pane words differently from the shared ribbon menu. */
export interface MacIterativeDialogLabels {
  /** The "automatic except for data tables" calculation-mode radio. */
  calcAutoNoTable: string;
}

/** Overrides of `iterativeDialog`, plus the Mac-only labels, that apply on
 *  the Mac host only. */
export type MacIterativeDialogStrings = Partial<IterativeDialogStrings> & MacIterativeDialogLabels;

/** The `iterativeDialog` keys the Mac pane can override. A locale whose Mac
 *  wording is not measured repeats its own shared text. */
const MAC_ITERATIVE_DIALOG_KEYS = [
  'enable',
  'maxIterations',
  'maxChange',
  'workbookSection',
] as const satisfies readonly (keyof IterativeDialogStrings)[];

export const macIterativeDialogFrom = (
  dialog: IterativeDialogStrings,
  labels: MacIterativeDialogLabels,
): MacIterativeDialogStrings => ({
  ...Object.fromEntries(MAC_ITERATIVE_DIALOG_KEYS.map((key) => [key, dialog[key]])),
  ...labels,
});

/** Japanese Mac app wording, measured from Excel for Mac. */
export const macIterativeDialogJa: MacIterativeDialogStrings = {
  enable: '反復計算を使う',
  maxIterations: '最大反復回数:',
  maxChange: '変化の最大値:',
  workbookSection: 'ブックの計算時',
  calcAutoNoTable: 'データ テーブル以外自動',
};
