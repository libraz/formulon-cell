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

/** The `iterativeDialog` keys the Mac pane can override. Each locale starts from its own shared text and overrides the fields the Excel bundle answers. */
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

/** Mac app wording from the Excel bundle; not answered by it: enable, workbookSection, calcAutoNoTable. */
export const macIterativeDialogDe: Partial<MacIterativeDialogStrings> = {
  maxIterations: 'Maximale Iterationszahl:',
  maxChange: 'Maximale Änderung:',
};

/** Mac app wording from the Excel bundle; not answered by it: enable, maxIterations, workbookSection, calcAutoNoTable. */
export const macIterativeDialogFr: Partial<MacIterativeDialogStrings> = {
  maxChange: 'Écart maximal :',
};

/** Mac app wording from the Excel bundle; not answered by it: enable, workbookSection, calcAutoNoTable. */
export const macIterativeDialogKo: Partial<MacIterativeDialogStrings> = {
  maxIterations: '최대 반복 횟수:',
  maxChange: '변화 한도값:',
};

/** Mac app wording from the Excel bundle; not answered by it: enable, workbookSection, calcAutoNoTable. */
export const macIterativeDialogTh: Partial<MacIterativeDialogStrings> = {
  maxIterations: 'รอบการคำนวณซ้ำมากที่สุด:',
  maxChange: 'ค่าแตกต่างสูงสุดที่รับได้:',
};

/** Mac app wording from the Excel bundle; not answered by it: enable, workbookSection, calcAutoNoTable. */
export const macIterativeDialogZh: Partial<MacIterativeDialogStrings> = {
  maxIterations: '最多迭代次数:',
  maxChange: '最大误差:',
};
