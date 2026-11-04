import type { EditDialogsStrings } from '../edit-dialogs/types.js';

type HyperlinkDialogStrings = EditDialogsStrings['hyperlinkDialog'];

/** Labels only the Mac Insert Hyperlink dialog shows. */
export interface MacHyperlinkDialogLabels {
  displayText: string;
  /** Display text shown, disabled, for a multi-cell selection. */
  selectionInDocument: string;
  screenTip: string;
  screenTipTitle: string;
  screenTipLabel: string;
  tabWeb: string;
  tabDocument: string;
  tabEmail: string;
  webIntro: string;
  browse: string;
  /** Why the file chooser is unavailable in a browser host. */
  browseUnavailable: string;
  cellReference: string;
  documentPlace: string;
  emailAddress: string;
  subject: string;
  recentEmails: string;
}

/** Overrides of `hyperlinkDialog`, plus the Mac-only labels, that apply on the Mac host only. */
export type MacHyperlinkDialogStrings = Partial<HyperlinkDialogStrings> & MacHyperlinkDialogLabels;

/** The `hyperlinkDialog` keys the Mac dialog can override. A locale whose Mac
 *  wording is not measured repeats its own shared text. */
const MAC_HYPERLINK_DIALOG_KEYS = [
  'url',
] as const satisfies readonly (keyof HyperlinkDialogStrings)[];

export const macHyperlinkDialogFrom = (
  dialog: HyperlinkDialogStrings,
  labels: MacHyperlinkDialogLabels,
): MacHyperlinkDialogStrings => ({
  ...Object.fromEntries(MAC_HYPERLINK_DIALOG_KEYS.map((key) => [key, dialog[key]])),
  ...labels,
});

/** Japanese Mac app wording, measured from Excel for Mac; the ScreenTip
 *  dialog's title and label and the browse reason are not measured. */
export const macHyperlinkDialogJa: MacHyperlinkDialogStrings = {
  url: 'アドレス:',
  displayText: '表示するテキスト:',
  selectionInDocument: '<<ドキュメント内の選択範囲>>',
  screenTip: 'ヒント設定...',
  screenTipTitle: 'ハイパーリンクのヒントの設定',
  screenTipLabel: 'ヒントのテキスト:',
  tabWeb: 'Web ページまたはファイル',
  tabDocument: 'このドキュメント',
  tabEmail: '電子メール アドレス',
  webIntro: '既存のファイルまたは Web ページにリンクします。',
  browse: '選択...',
  browseUnavailable: 'ブラウザーではファイルを選択できません。アドレスを入力してください。',
  cellReference: 'セル参照を入力してください:',
  documentPlace: 'またはドキュメント内の場所を選択してください:',
  emailAddress: '電子メール アドレス:',
  subject: '件名:',
  recentEmails: '最近使用した電子メール アドレス:',
};
