import type { PageLayoutStrings } from '../page-layout/types.js';

type PageSetupStrings = PageLayoutStrings['pageSetup'];

/** Overrides of `pageSetup` that apply on the Mac host only. */
export type MacPageSetupStrings = Partial<PageSetupStrings>;

/** The `pageSetup` keys the Mac dialog can override. A locale whose Mac
 *  wording is not measured repeats its own shared text. */
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
