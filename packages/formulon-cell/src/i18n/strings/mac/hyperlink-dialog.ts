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

/** The `hyperlinkDialog` keys the Mac dialog can override. Each locale starts from its own shared text and overrides the fields the Excel bundle answers. */
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

/** Mac app wording from the Excel bundle; not answered by it: url, selectionInDocument, tabEmail, browse, browseUnavailable, documentPlace, and 2 more. */
export const macHyperlinkDialogDe: Partial<MacHyperlinkDialogStrings> = {
  displayText: 'Anzuzeigender Text:',
  screenTip: 'QuickInfo...',
  screenTipTitle: 'Quickinfo für Link festlegen',
  screenTipLabel: 'QuickInfo-Text:',
  tabWeb: 'Webseite oder Datei',
  tabDocument: 'Dieses Dokument',
  webIntro: 'Verknüpfung zu einer vorhandenen Datei oder Webseite.',
  cellReference: 'Zellbezug eingeben:',
  recentEmails: 'Zuletzt verwendete E-Mail-Adressen:',
};

/** Mac app wording from the Excel bundle; not answered by it: url, selectionInDocument, tabEmail, browse, browseUnavailable, documentPlace, and 2 more. */
export const macHyperlinkDialogFr: Partial<MacHyperlinkDialogStrings> = {
  displayText: 'Texte à afficher :',
  screenTip: 'Info-bulle...',
  screenTipTitle: 'Définir une info-bulle pour le lien hypertexte',
  screenTipLabel: "Texte de l'info-bulle :",
  tabWeb: 'Page web ou fichier',
  tabDocument: 'Ce document',
  webIntro: 'Définir un lien vers une page web ou un fichier existant.',
  cellReference: 'Entrer la référence de la cellule :',
  recentEmails: 'Adresses de courrier récemment utilisées :',
};

/** Mac app wording from the Excel bundle; not answered by it: url, selectionInDocument, tabEmail, browse, browseUnavailable, documentPlace, and 2 more. */
export const macHyperlinkDialogKo: Partial<MacHyperlinkDialogStrings> = {
  displayText: '표시할 텍스트:',
  screenTip: '화면 설명...',
  screenTipTitle: '하이퍼링크 화면 설명 설정',
  screenTipLabel: '화면 설명 텍스트:',
  tabWeb: '웹 페이지 또는 파일',
  tabDocument: '이 문서',
  webIntro: '기존 파일 또는 웹 페이지에 연결합니다.',
  cellReference: '참조할 셀 입력:',
  recentEmails: '최근 사용한 전자 메일 주소:',
};

/** Mac app wording from the Excel bundle; not answered by it: url, selectionInDocument, tabEmail, browse, browseUnavailable, documentPlace, and 2 more. */
export const macHyperlinkDialogTh: Partial<MacHyperlinkDialogStrings> = {
  displayText: 'ข้อความที่จะแสดง:',
  screenTip: 'คำแนะนำบนหน้าจอ...',
  screenTipTitle: 'การตั้งค่าคำแนะนำบนหน้าจอไฮเปอร์ลิงก์',
  screenTipLabel: 'ข้อความคำแนะนำบนหน้าจอ:',
  tabWeb: 'เว็บเพจหรือไฟล์',
  tabDocument: 'เอกสารนี้',
  webIntro: 'ลิงก์ไปยังไฟล์หรือเว็บเพจที่มีอยู่',
  cellReference: 'พิมพ์การอ้างอิงเซลล์:',
  recentEmails: 'อีเมลแอดเดรสที่ใช้ล่าสุด:',
};

/** Mac app wording from the Excel bundle; not answered by it: url, selectionInDocument, tabEmail, browse, browseUnavailable, documentPlace, and 2 more. */
export const macHyperlinkDialogZh: Partial<MacHyperlinkDialogStrings> = {
  displayText: '要显示的文本:',
  screenTip: '屏幕提示...',
  screenTipTitle: '设置超链接屏幕提示',
  screenTipLabel: '屏幕提示文字:',
  tabWeb: '网页或文件',
  tabDocument: '此文档',
  webIntro: '链接到现有文件或网页。',
  cellReference: '输入单元格引用:',
  recentEmails: '最近使用过的电子邮件地址:',
};

/** Mac app wording from the Excel bundle. */
export const macHyperlinkDialogEn: Partial<MacHyperlinkDialogStrings> = {
  displayText: 'Text to Display:',
};
