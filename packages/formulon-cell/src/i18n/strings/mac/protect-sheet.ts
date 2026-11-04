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
 *  Protect Sheet dialog and the protected-cell alert only. */
export type MacProtectSheetStrings = Partial<ProtectionStrings> & MacProtectSheetLabels;

/** The `protection` keys the Mac dialog can override. A locale whose Mac
 *  wording is not measured repeats its own shared text. */
const MAC_PROTECT_SHEET_KEYS = [
  'protectSheet',
  'password',
  'confirmPassword',
  'allowSelectLockedCells',
  'allowSelectUnlockedCells',
  'allowPivotTables',
  'protectedCellEdit',
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
  protectedCellEdit:
    '変更しようとしているセルやグラフは保護されているシート上にあります。変更するには、シートの保護を解除してください。パスワードの入力が必要な場合もあります。',
  heading: 'シートとロックされたセルの内容を保護する',
  note: '既定ではすべてのセルがロックされていますが、ロックは解除可能です。',
};
