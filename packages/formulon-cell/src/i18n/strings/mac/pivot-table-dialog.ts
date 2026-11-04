import type { DataObjectsStrings } from '../data-objects/types.js';

type PivotTableDialogStrings = DataObjectsStrings['pivotTableDialog'];

/** Labels only the Mac Create PivotTable dialog shows. */
export interface MacPivotTableDialogLabels {
  /** Caption of the disabled button under the external-source choice. */
  chooseConnection: string;
  /** Title of the Move PivotTable dialog. */
  moveTitle: string;
  /** Caption of the Move PivotTable dialog's location field. */
  moveLocation: string;
}

/** Overrides of `pivotTableDialog`, plus the Mac-only labels, that apply on
 *  the Mac host only. */
export type MacPivotTableDialogStrings = Partial<PivotTableDialogStrings> &
  MacPivotTableDialogLabels;

/** The `pivotTableDialog` keys the Mac dialog can override. A locale whose Mac
 *  wording is not measured repeats its own shared text. */
const MAC_PIVOT_TABLE_DIALOG_KEYS = [
  'source',
  'destinationSection',
  'destination',
] as const satisfies readonly (keyof PivotTableDialogStrings)[];

export const macPivotTableDialogFrom = (
  dialog: PivotTableDialogStrings,
  labels: MacPivotTableDialogLabels,
): MacPivotTableDialogStrings => ({
  ...Object.fromEntries(MAC_PIVOT_TABLE_DIALOG_KEYS.map((key) => [key, dialog[key]])),
  ...labels,
});

/** Japanese Mac app wording, measured from Excel for Mac. */
export const macPivotTableDialogJa: MacPivotTableDialogStrings = {
  source: 'テーブル/範囲:',
  destinationSection: 'ピボットテーブルの作成先を選択してください。',
  destination: 'テーブル/範囲:',
  chooseConnection: '接続の選択...',
  moveTitle: 'ピボットテーブルの移動',
  moveLocation: '場所:',
};
