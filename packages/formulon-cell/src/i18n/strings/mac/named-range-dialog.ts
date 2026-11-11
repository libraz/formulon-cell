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

/** The `namedRangeDialog` keys the Mac Name Manager can override. Each locale starts from its own shared text and overrides the fields the Excel bundle answers. */
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

/** Mac app wording from the Excel bundle; not answered by it: newButton, defineTitle. */
export const macNamedRangeDialogDe: Partial<MacNamedRangeDialogStrings> = {
  title: 'Namens-Manager',
  formulaHeader: 'Bezieht sich auf',
  quickRefersLabel: 'Bezieht sich auf:',
  formulaLabel: 'Bezieht sich auf:',
  defineNames: 'Namen:',
  defineNamePrompt: 'Geben Sie einen Namen für den Datenbereich ein:',
  defineRangePrompt: 'Wählen Sie den Zellbereich aus:',
};

/** Mac app wording from the Excel bundle; not answered by it: newButton, defineTitle. */
export const macNamedRangeDialogFr: Partial<MacNamedRangeDialogStrings> = {
  title: 'Gestionnaire de noms',
  formulaHeader: 'Fait référence à',
  quickRefersLabel: 'Fait référence à :',
  formulaLabel: 'Fait référence à :',
  defineNames: 'Noms :',
  defineNamePrompt: 'Entrer un nom pour la plage de données :',
  defineRangePrompt: 'Sélectionner la plage de cellules :',
};

/** Mac app wording from the Excel bundle; not answered by it: newButton, defineTitle, defineNames. */
export const macNamedRangeDialogKo: Partial<MacNamedRangeDialogStrings> = {
  title: '이름 관리자',
  formulaHeader: '참조 대상',
  quickRefersLabel: '참조 대상:',
  formulaLabel: '참조 대상:',
  defineNamePrompt: '데이터 범위 이름 입력:',
  defineRangePrompt: '셀 범위 선택:',
};

/** Mac app wording from the Excel bundle; not answered by it: newButton, defineTitle. */
export const macNamedRangeDialogTh: Partial<MacNamedRangeDialogStrings> = {
  title: 'ตัวจัดการชื่อ',
  formulaHeader: 'อ้างอิงไปยัง',
  quickRefersLabel: 'อ้างถึง:',
  formulaLabel: 'อ้างถึง:',
  defineNames: 'ชื่อ:',
  defineNamePrompt: 'ใส่ชื่อสำหรับช่วงข้อมูล:',
  defineRangePrompt: 'เลือกช่วงของเซลล์:',
};

/** Mac app wording from the Excel bundle; not answered by it: newButton, defineTitle, defineNames. */
export const macNamedRangeDialogZh: Partial<MacNamedRangeDialogStrings> = {
  title: '名称管理器',
  formulaHeader: '引用位置',
  quickRefersLabel: '引用:',
  formulaLabel: '引用:',
  defineNamePrompt: '输入数据区域的名称:',
  defineRangePrompt: '选择单元格区域:',
};
