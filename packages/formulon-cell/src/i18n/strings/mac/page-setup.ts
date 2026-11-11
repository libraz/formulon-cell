import type { PageLayoutStrings } from '../page-layout/types.js';

type PageSetupStrings = PageLayoutStrings['pageSetup'];

/** Overrides of `pageSetup` that apply on the Mac host only. */
export type MacPageSetupStrings = Partial<PageSetupStrings>;

/** The `pageSetup` keys the Mac dialog can override. Each locale starts from its own shared text and overrides the fields the Excel bundle answers. */
const MAC_PAGE_SETUP_KEYS = [
  'marginTop',
  'marginRight',
  'marginBottom',
  'marginLeft',
  'marginHeader',
  'marginFooter',
  'centerOnPage',
  'headerLabel',
  'footerLabel',
  'headerBuiltin',
  'footerBuiltin',
  'headerNone',
  'customHeader',
  'customFooter',
  'alignWithPageMargins',
  'printArea',
  'printTitleRows',
  'printTitleCols',
  'printOptions',
  'showGridlines',
  'showHeadings',
  'comments',
  'pageOrderDownThenOver',
  'pageOrderOverThenDown',
  'adjustTo',
  'percentNormalSize',
  'fitTo',
  'pagesWideBy',
  'tall',
  'firstPageNumber',
] as const satisfies readonly (keyof PageSetupStrings)[];

export const macPageSetupFrom = (dialog: PageSetupStrings): MacPageSetupStrings =>
  Object.fromEntries(MAC_PAGE_SETUP_KEYS.map((key) => [key, dialog[key]]));

/** Japanese Mac app wording, measured from Excel for Mac. */
export const macPageSetupJa: MacPageSetupStrings = {
  marginTop: '上:',
  marginRight: '右:',
  marginBottom: '下:',
  marginLeft: '左:',
  marginHeader: 'ヘッダー:',
  marginFooter: 'フッター:',
  centerOnPage: 'ページ中央:',
  headerLabel: 'ヘッダー:',
  footerLabel: 'フッター:',
  headerBuiltin: 'ヘッダー:',
  footerBuiltin: 'フッター:',
  headerNone: '(指定しない)',
  customHeader: 'ユーザー設定のヘッダー...',
  customFooter: 'ユーザー設定のフッター...',
  alignWithPageMargins: 'ページ余白に合わせて配置',
  printArea: '印刷範囲:',
  printTitleRows: 'タイトル行:',
  printTitleCols: 'タイトル列:',
  printOptions: '印刷オプション',
  showGridlines: '枠線',
  showHeadings: '行列番号',
  comments: 'コメント:',
  pageOrderDownThenOver: '左から右',
  pageOrderOverThenDown: '上から下',
  adjustTo: '拡大/縮小:',
  percentNormalSize: '%',
  fitTo: '次のページ数に合わせる:',
  pagesWideBy: 'ページ (横) ×',
  tall: 'ページ (縦)',
  firstPageNumber: '先頭ページ番号:',
};

/** Mac app wording from the Excel bundle; not answered by it: marginTop, marginRight, marginBottom, marginLeft, marginHeader, marginFooter, and 11 more. */
export const macPageSetupDe: Partial<MacPageSetupStrings> = {
  centerOnPage: 'Auf der Seite zentrieren:',
  headerNone: '(keine)',
  alignWithPageMargins: 'An Seitenrändern ausrichten',
  printTitleCols: 'Wiederholungsspalten links:',
  showGridlines: 'Gitternetzlinien drucken',
  pageOrderDownThenOver: 'Zuerst nach unten',
  pageOrderOverThenDown: 'Zuerst nach rechts',
  adjustTo: 'Verkleinern/Vergrößern:',
  percentNormalSize: '% Normalgröße',
  fitTo: 'Anpassen:',
  pagesWideBy: 'Seiten breit und',
  tall: 'Seiten hoch',
  firstPageNumber: 'Erste Seitenzahl:',
};

/** Mac app wording from the Excel bundle; not answered by it: marginTop, marginRight, marginBottom, marginLeft, marginHeader, marginFooter, and 11 more. */
export const macPageSetupFr: Partial<MacPageSetupStrings> = {
  centerOnPage: 'Centrer sur la page :',
  headerNone: '(aucun)',
  alignWithPageMargins: 'Aligner sur les marges de page',
  printTitleCols: 'Colonnes à répéter à gauche :',
  showGridlines: 'Imprimer le quadrillage',
  pageOrderDownThenOver: 'Vers le bas, puis à droite',
  pageOrderOverThenDown: 'À droite, puis vers le bas',
  adjustTo: 'Réduire/agrandir à :',
  percentNormalSize: '% de la taille normale',
  fitTo: 'Adapter à :',
  pagesWideBy: 'Pages en largeur sur',
  tall: 'En hauteur',
  firstPageNumber: 'Commencer la numérotation à :',
};

/** Mac app wording from the Excel bundle; not answered by it: marginTop, marginRight, marginBottom, marginLeft, marginHeader, marginFooter, and 11 more. */
export const macPageSetupKo: Partial<MacPageSetupStrings> = {
  centerOnPage: '페이지 가운데 맞춤:',
  headerNone: '(없음)',
  alignWithPageMargins: '페이지 여백에 맞추기',
  printTitleCols: '반복할 열:',
  showGridlines: '눈금선 인쇄',
  pageOrderDownThenOver: '행 우선',
  pageOrderOverThenDown: '열 우선',
  adjustTo: '축소/확대 배율:',
  percentNormalSize: '%',
  fitTo: '자동 맞춤:',
  pagesWideBy: '페이지 너비',
  tall: '페이지 높이',
  firstPageNumber: '시작 페이지 번호:',
};

/** Mac app wording from the Excel bundle; not answered by it: marginTop, marginRight, marginBottom, marginLeft, marginHeader, marginFooter, and 11 more. */
export const macPageSetupTh: Partial<MacPageSetupStrings> = {
  centerOnPage: 'กึ่งกลางหน้ากระดาษ:',
  headerNone: '(ไม่มี)',
  alignWithPageMargins: 'จัดชิดระยะขอบกระดาษ',
  printTitleCols: 'คอลัมน์ที่จะพิมพ์ซ้ำทางซ้าย:',
  showGridlines: 'พิมพ์เส้นตาราง',
  pageOrderDownThenOver: 'ลงแล้วซ้ายไปขวา',
  pageOrderOverThenDown: 'ขึ้นแล้วลง',
  adjustTo: 'ปรับเป็น:',
  percentNormalSize: '% จากขนาดปกติ',
  fitTo: 'จัดพอดีกับ:',
  pagesWideBy: 'หน้ากว้าง',
  tall: 'สูง',
  firstPageNumber: 'หมายเลขหน้าแรก:',
};

/** Mac app wording from the Excel bundle; not answered by it: marginTop, marginRight, marginBottom, marginLeft, marginHeader, marginFooter, and 11 more. */
export const macPageSetupZh: Partial<MacPageSetupStrings> = {
  centerOnPage: '居中:',
  headerNone: '(无)',
  alignWithPageMargins: '与页边距对齐',
  printTitleCols: '从左侧重复的列数:',
  showGridlines: '打印网格线',
  pageOrderDownThenOver: '先列后行',
  pageOrderOverThenDown: '先行后列',
  adjustTo: '缩放比例:',
  percentNormalSize: '% 正常大小',
  fitTo: '调整为:',
  pagesWideBy: '页宽 x',
  tall: '页高',
  firstPageNumber: '首页页码:',
};

/** Mac app wording from the Excel bundle. */
export const macPageSetupEn: Partial<MacPageSetupStrings> = {
  centerOnPage: 'Center on page:',
  adjustTo: 'Adjust to:',
  fitTo: 'Fit to:',
  tall: 'Tall',
  firstPageNumber: 'First page number:',
  printTitleCols: 'Columns to repeat at left:',
  pagesWideBy: 'Pages wide by',
};
