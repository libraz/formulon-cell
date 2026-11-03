import type { EditDialogsStrings } from '../edit-dialogs/types.js';

/** Overrides of `pasteSpecialDialog` that apply on the Mac host only. */
export type MacPasteSpecialStrings = Partial<EditDialogsStrings['pasteSpecialDialog']>;

/** The `pasteSpecialDialog` keys the Mac host can override. A locale whose Mac
 *  wording is not measured repeats its own shared text for them. */
const MAC_PASTE_SPECIAL_KEYS = [
  'title',
  'sectionPaste',
  'pasteValues',
  'pasteAllUsingSourceTheme',
  'pasteAllMergingCF',
  'skipBlanks',
  'pasteLink',
] as const satisfies readonly (keyof EditDialogsStrings['pasteSpecialDialog'])[];

export const macPasteSpecialFrom = (
  dialog: EditDialogsStrings['pasteSpecialDialog'],
): MacPasteSpecialStrings =>
  Object.fromEntries(MAC_PASTE_SPECIAL_KEYS.map((key) => [key, dialog[key]]));

/** Japanese Mac app wording, measured from Excel for Mac: ペースト where the
 *  shared text says 貼り付け, and its own labels for three options. */
export const macPasteSpecialJa: MacPasteSpecialStrings = {
  title: '形式を選択してペースト',
  sectionPaste: 'ペースト',
  pasteValues: '数値',
  pasteAllUsingSourceTheme: 'コピー元のテーマを使ってすべてペースト',
  pasteAllMergingCF: 'すべて、条件付き書式の結合',
  skipBlanks: '空白セルを無視',
  pasteLink: 'ペースト リンク',
};
