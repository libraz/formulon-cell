import type { PageLayoutStrings } from '../page-layout/types.js';

type PageSetupStrings = PageLayoutStrings['pageSetup'];

/** Overrides of `pageSetup` that apply on the Mac host only. */
export type MacPageSetupStrings = Partial<PageSetupStrings>;

/** The `pageSetup` keys the Mac dialog can override. A locale whose Mac
 *  wording is not measured repeats its own shared text. */
const MAC_PAGE_SETUP_KEYS = [
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
  adjustTo: '拡大/縮小:',
  percentNormalSize: '%',
  fitTo: '次のページ数に合わせる:',
  pagesWideBy: 'ページ (横) ×',
  tall: 'ページ (縦)',
  firstPageNumber: '先頭ページ番号:',
};
