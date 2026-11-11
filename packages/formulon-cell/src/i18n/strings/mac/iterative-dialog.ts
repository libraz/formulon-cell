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

/** Mac app wording from the Excel bundle. */
export const macIterativeDialogDe: Partial<MacIterativeDialogStrings> = {
  enable: 'Iterative Berechnung verwenden',
  maxIterations: 'Maximale Iterationszahl:',
  maxChange: 'Maximale Änderung:',
  workbookSection: 'Bei der Berechnung von Arbeitsmappen',
  calcAutoNoTable: 'Automatisch außer bei Datentabellen',
};

/** Mac app wording from the Excel bundle; not answered by it: maxIterations. */
export const macIterativeDialogFr: Partial<MacIterativeDialogStrings> = {
  enable: 'Activer le calcul itératif',
  maxChange: 'Écart maximal :',
  workbookSection: 'Lors du calcul des classeurs',
  calcAutoNoTable: 'Automatique sauf dans les tables de données',
};

/** Mac app wording from the Excel bundle. */
export const macIterativeDialogKo: Partial<MacIterativeDialogStrings> = {
  enable: '반복 계산 사용',
  maxIterations: '최대 반복 횟수:',
  maxChange: '변화 한도값:',
  workbookSection: '통합 문서를 계산하는 경우',
  calcAutoNoTable: '데이터 표만 수동',
};

/** Mac app wording from the Excel bundle. */
export const macIterativeDialogTh: Partial<MacIterativeDialogStrings> = {
  enable: 'ใช้การคำนวณแบบซ้ำ',
  maxIterations: 'รอบการคำนวณซ้ำมากที่สุด:',
  maxChange: 'ค่าแตกต่างสูงสุดที่รับได้:',
  workbookSection: 'เมื่อคำนวณเวิร์กบุ๊ก',
  calcAutoNoTable: 'อัตโนมัติยกเว้นตารางข้อมูล',
};

/** Mac app wording from the Excel bundle. */
export const macIterativeDialogZh: Partial<MacIterativeDialogStrings> = {
  enable: '使用迭代计算',
  maxIterations: '最多迭代次数:',
  maxChange: '最大误差:',
  workbookSection: '计算工作簿时',
  calcAutoNoTable: '除模拟运算表外，自动计算',
};

/** Mac app wording from the Excel bundle. */
export const macIterativeDialogEn: Partial<MacIterativeDialogStrings> = {
  maxIterations: 'Maximum iterations:',
  maxChange: 'Maximum change:',
  enable: 'Use iterative calculation',
  workbookSection: 'When Calculating Workbooks',
  calcAutoNoTable: 'Automatic except for data tables',
};
