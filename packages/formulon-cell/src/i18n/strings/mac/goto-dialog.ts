import type { EditDialogsStrings } from '../edit-dialogs/types.js';

type GoToDialogStrings = EditDialogsStrings['goToDialog'];

/** Lines only the Mac Go To Special dialog shows. */
export interface MacGoToDialogLabels {
  /** Caption above the option radios. */
  selectCaption: string;
  /** Legend of the box that holds the chosen option's sub-options. */
  optionsGroup: string;
  /** Box text for an option without sub-options. */
  noOptions: string;
}

/** Overrides of `goToDialog`, plus the Mac-only lines, that apply on the Mac host only. */
export type MacGoToDialogStrings = Partial<GoToDialogStrings> & MacGoToDialogLabels;

/** The `goToDialog` keys the Mac dialog can override. Each locale starts from its own shared text and overrides the fields the Excel bundle answers. */
const MAC_GO_TO_DIALOG_KEYS = [
  'kindComments',
  'kindBlanks',
  'reference',
] as const satisfies readonly (keyof GoToDialogStrings)[];

export const macGoToDialogFrom = (
  dialog: GoToDialogStrings,
  labels: MacGoToDialogLabels,
): MacGoToDialogStrings => ({
  ...Object.fromEntries(MAC_GO_TO_DIALOG_KEYS.map((key) => [key, dialog[key]])),
  ...labels,
});

/** Japanese Mac app wording, measured from Excel for Mac. */
export const macGoToDialogJa: MacGoToDialogStrings = {
  kindComments: 'メモ',
  kindBlanks: '空白',
  reference: '参照:',
  selectCaption: '選択',
  optionsGroup: 'オプション',
  noOptions: 'この選択項目にはオプションはありません。',
};

/** Mac app wording from the Excel bundle; not answered by it: kindComments, kindBlanks, reference, selectCaption, optionsGroup. */
export const macGoToDialogDe: Partial<MacGoToDialogStrings> = {
  noOptions: 'Für diese Auswahl sind keine Optionen vorhanden.',
};

/** Mac app wording from the Excel bundle; not answered by it: kindComments, kindBlanks, reference, selectCaption, optionsGroup. */
export const macGoToDialogFr: Partial<MacGoToDialogStrings> = {
  noOptions: 'Il n’y a aucune option pour cette sélection.',
};

/** Mac app wording from the Excel bundle; not answered by it: kindComments, kindBlanks, reference, selectCaption, optionsGroup. */
export const macGoToDialogKo: Partial<MacGoToDialogStrings> = {
  noOptions: '이 선택 사항에 대한 옵션이 없습니다.',
};

/** Mac app wording from the Excel bundle; not answered by it: kindComments, reference, selectCaption, optionsGroup. */
export const macGoToDialogTh: Partial<MacGoToDialogStrings> = {
  kindBlanks: 'ว่าง',
  noOptions: 'ไม่มีตัวเลือกสำหรับการเลือกนี้',
};

/** Mac app wording from the Excel bundle; not answered by it: kindComments, kindBlanks, reference, selectCaption, optionsGroup. */
export const macGoToDialogZh: Partial<MacGoToDialogStrings> = {
  noOptions: '没有适合此选择的选项。',
};

/** Mac app wording from the Excel bundle. */
export const macGoToDialogEn: Partial<MacGoToDialogStrings> = {
  noOptions: 'There are no options for this selection.',
};
