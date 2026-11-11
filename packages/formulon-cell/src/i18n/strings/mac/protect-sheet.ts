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

/** Mac app wording from the Excel bundle; not answered by it: protectSheet, confirmPassword, allowPivotTables, heading. */
export const macProtectSheetDe: Partial<MacProtectSheetStrings> = {
  password: 'Kennwort (optional):',
  allowSelectLockedCells: 'Gesperrte Zellen auswählen',
  allowSelectUnlockedCells: 'Nicht gesperrte Zellen auswählen',
  note: 'Alle Zellen sind standardmäßig gesperrt, können aber als nicht gesperrt formatiert werden.',
};

/** Mac app wording from the Excel bundle; not answered by it: protectSheet, confirmPassword, allowPivotTables. */
export const macProtectSheetFr: Partial<MacProtectSheetStrings> = {
  password: 'Mot de passe (facultatif) :',
  allowSelectLockedCells: 'Sélectionner les cellules verrouillées',
  allowSelectUnlockedCells: 'Sélectionner les cellules déverrouillées',
  heading: 'Protéger la feuille et le contenu des cellules verrouillées',
  note: 'Toutes les cellules sont verrouillées par défaut, mais peuvent être mises en forme comme déverrouillées.',
};

/** Mac app wording from the Excel bundle; not answered by it: protectSheet, confirmPassword, allowPivotTables. */
export const macProtectSheetKo: Partial<MacProtectSheetStrings> = {
  password: '암호(옵션):',
  allowSelectLockedCells: '잠긴 셀 선택',
  allowSelectUnlockedCells: '잠기지 않은 셀 선택',
  heading: '잠긴 셀의 내용과 워크시트 보호',
  note: '기본적으로 모든 셀은 잠겨 있지만 잠금 해제된 서식으로 지정할 수 있습니다.',
};

/** Mac app wording from the Excel bundle; not answered by it: protectSheet, confirmPassword, allowPivotTables. */
export const macProtectSheetTh: Partial<MacProtectSheetStrings> = {
  password: 'รหัสผ่าน (ใส่หรือไม่ก็ได้):',
  allowSelectLockedCells: 'เลือกเซลล์ที่ถูกล็อก',
  allowSelectUnlockedCells: 'เลือกเซลล์ที่ไม่ได้ถูกล็อก',
  heading: 'ป้องกันเวิร์กชีตและเนื้อหาในเซลล์ที่ถูกล็อก',
  note: 'เซลล์ทั้งหมดถูกล็อกไว้เป็นค่าเริ่มต้น แต่ฟอร์แมตให้ปลดล็อกได้',
};

/** Mac app wording from the Excel bundle; not answered by it: protectSheet, confirmPassword, allowPivotTables. */
export const macProtectSheetZh: Partial<MacProtectSheetStrings> = {
  password: '密码(可选):',
  allowSelectLockedCells: '选择锁定的单元格',
  allowSelectUnlockedCells: '选择未锁定的单元格',
  heading: '保护工作表及锁定的单元格内容',
  note: '默认情况下，所有单元格均锁定，但可设置为未锁定的格式。',
};

/** Mac app wording from the Excel bundle. */
export const macProtectSheetEn: Partial<MacProtectSheetStrings> = {
  password: 'Password (optional):',
  note: 'All cells are locked by default, but can be formatted as unlocked.',
};
