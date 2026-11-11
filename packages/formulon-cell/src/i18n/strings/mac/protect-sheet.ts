import type { EditDialogsStrings } from '../edit-dialogs/types.js';

type ProtectionStrings = EditDialogsStrings['protection'];

/** Lines only the Mac Protect Sheet dialog shows. */
export interface MacProtectSheetLabels {
  /** Bold line above the password fields. */
  heading: string;
  /** Explanatory line under the heading. */
  note: string;
}

/** Overrides of `protection`, plus the Mac-only lines, that apply to the Mac
 *  Protect Sheet dialog only. */
export type MacProtectSheetStrings = Partial<ProtectionStrings> & MacProtectSheetLabels;

/** The `protection` keys the Mac dialog can override. Each locale starts from its own shared text and overrides the fields the Excel bundle answers. */
const MAC_PROTECT_SHEET_KEYS = [
  'protectSheet',
  'password',
  'confirmPassword',
  'allowSelectLockedCells',
  'allowSelectUnlockedCells',
  'allowPivotTables',
] as const satisfies readonly (keyof ProtectionStrings)[];

export const macProtectSheetFrom = (
  protection: ProtectionStrings,
  labels: MacProtectSheetLabels,
): MacProtectSheetStrings => ({
  ...Object.fromEntries(MAC_PROTECT_SHEET_KEYS.map((key) => [key, protection[key]])),
  ...labels,
});

/** Japanese Mac app wording, measured from Excel for Mac. */
export const macProtectSheetJa: MacProtectSheetStrings = {
  protectSheet: 'シートの保護',
  password: 'パスワード (省略可):',
  confirmPassword: '確認:',
  allowSelectLockedCells: 'ロックされたセルの選択',
  allowSelectUnlockedCells: 'ロックされていないセルの選択',
  allowPivotTables: 'ピボットテーブルとピボットグラフを使用する',
  heading: 'シートとロックされたセルの内容を保護する',
  note: '既定ではすべてのセルがロックされていますが、ロックは解除可能です。',
};

/** Mac app wording from the Excel bundle; not answered by it: protectSheet, password, confirmPassword, allowPivotTables, heading, note. */
export const macProtectSheetDe: Partial<MacProtectSheetStrings> = {
  allowSelectLockedCells: 'Gesperrte Zellen auswählen',
  allowSelectUnlockedCells: 'Nicht gesperrte Zellen auswählen',
};

/** Mac app wording from the Excel bundle; not answered by it: protectSheet, password, confirmPassword, allowPivotTables, note. */
export const macProtectSheetFr: Partial<MacProtectSheetStrings> = {
  allowSelectLockedCells: 'Sélectionner les cellules verrouillées',
  allowSelectUnlockedCells: 'Sélectionner les cellules déverrouillées',
  heading: 'Protéger la feuille et le contenu des cellules verrouillées',
};

/** Mac app wording from the Excel bundle; not answered by it: protectSheet, password, confirmPassword, allowPivotTables, note. */
export const macProtectSheetKo: Partial<MacProtectSheetStrings> = {
  allowSelectLockedCells: '잠긴 셀 선택',
  allowSelectUnlockedCells: '잠기지 않은 셀 선택',
  heading: '잠긴 셀의 내용과 워크시트 보호',
};

/** Mac app wording from the Excel bundle; not answered by it: protectSheet, password, confirmPassword, allowPivotTables, note. */
export const macProtectSheetTh: Partial<MacProtectSheetStrings> = {
  allowSelectLockedCells: 'เลือกเซลล์ที่ถูกล็อก',
  allowSelectUnlockedCells: 'เลือกเซลล์ที่ไม่ได้ถูกล็อก',
  heading: 'ป้องกันเวิร์กชีตและเนื้อหาในเซลล์ที่ถูกล็อก',
};

/** Mac app wording from the Excel bundle; not answered by it: protectSheet, password, confirmPassword, allowPivotTables, note. */
export const macProtectSheetZh: Partial<MacProtectSheetStrings> = {
  allowSelectLockedCells: '选择锁定的单元格',
  allowSelectUnlockedCells: '选择未锁定的单元格',
  heading: '保护工作表及锁定的单元格内容',
};
