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

/** The `goToDialog` keys the Mac dialog can override. A locale whose Mac
 *  wording is not measured repeats its own shared text. */
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
