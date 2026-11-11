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

/** The `pivotTableDialog` keys the Mac dialog can override. Each locale starts from its own shared text and overrides the fields the Excel bundle answers. */
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

/** Mac app wording from the Excel bundle; not answered by it: destinationSection, destination, moveLocation. */
export const macPivotTableDialogDe: Partial<MacPivotTableDialogStrings> = {
  source: 'Tabelle/Bereich:',
  chooseConnection: 'Verbindung auswählen...',
  moveTitle: 'PivotTable verschieben',
};

/** Mac app wording from the Excel bundle; not answered by it: destinationSection, destination, chooseConnection, moveLocation. */
export const macPivotTableDialogFr: Partial<MacPivotTableDialogStrings> = {
  source: 'Tableau/Plage :',
  moveTitle: 'Déplacer le tableau croisé dynamique',
};

/** Mac app wording from the Excel bundle; not answered by it: destinationSection, destination, moveLocation. */
export const macPivotTableDialogKo: Partial<MacPivotTableDialogStrings> = {
  source: '표/범위:',
  chooseConnection: '연결 선택...',
  moveTitle: '피벗 테이블 이동',
};

/** Mac app wording from the Excel bundle; not answered by it: destinationSection, destination, moveLocation. */
export const macPivotTableDialogTh: Partial<MacPivotTableDialogStrings> = {
  source: 'ตาราง/ช่วง:',
  chooseConnection: 'เลือกการเชื่อมต่อ...',
  moveTitle: 'ย้าย PivotTable',
};

/** Mac app wording from the Excel bundle; not answered by it: destinationSection, destination, moveLocation. */
export const macPivotTableDialogZh: Partial<MacPivotTableDialogStrings> = {
  source: '表/区域:',
  chooseConnection: '选择连接...',
  moveTitle: '移动数据透视表',
};
