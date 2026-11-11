import type { EditDialogsStrings } from '../edit-dialogs/types.js';

/** Overrides of `pasteSpecialDialog` that apply on the Mac host only. */
export type MacPasteSpecialStrings = Partial<EditDialogsStrings['pasteSpecialDialog']>;

/** The `pasteSpecialDialog` keys the Mac host can override. Each locale starts from its own shared text and overrides the fields the Excel bundle answers. */
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

/** Mac app wording from the Excel bundle; not answered by it: title, sectionPaste, pasteAllMergingCF, pasteLink. */
export const macPasteSpecialDe: Partial<MacPasteSpecialStrings> = {
  pasteValues: 'Werte',
  pasteAllUsingSourceTheme: 'Alles mit Quelldesign',
  skipBlanks: 'Leerzellen überspringen',
};

/** Mac app wording from the Excel bundle; not answered by it: title, sectionPaste, pasteAllMergingCF, pasteLink. */
export const macPasteSpecialFr: Partial<MacPasteSpecialStrings> = {
  pasteValues: 'Valeurs',
  pasteAllUsingSourceTheme: 'Tout, avec le thème source',
  skipBlanks: 'Blancs non compris',
};

/** Mac app wording from the Excel bundle; not answered by it: title, sectionPaste, pasteAllMergingCF, pasteLink. */
export const macPasteSpecialKo: Partial<MacPasteSpecialStrings> = {
  pasteValues: '값',
  pasteAllUsingSourceTheme: '원본 테마 사용',
  skipBlanks: '내용 있는 셀만 붙여넣기',
};

/** Mac app wording from the Excel bundle; not answered by it: title, sectionPaste, pasteAllMergingCF, pasteLink. */
export const macPasteSpecialTh: Partial<MacPasteSpecialStrings> = {
  pasteValues: 'ค่า',
  pasteAllUsingSourceTheme: 'ทั้งหมดที่ใช้ชุดรูปแบบต้นฉบับ',
  skipBlanks: 'ข้ามเซลล์ที่ว่าง',
};

/** Mac app wording from the Excel bundle; not answered by it: title, sectionPaste, pasteAllMergingCF, pasteLink. */
export const macPasteSpecialZh: Partial<MacPasteSpecialStrings> = {
  pasteValues: '值',
  pasteAllUsingSourceTheme: '全部使用源主题',
  skipBlanks: '跳过空单元',
};
