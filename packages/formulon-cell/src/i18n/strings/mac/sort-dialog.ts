import type { RibbonMenusStrings } from '../ribbon-menus/types.js';

type RibbonMenuStrings = RibbonMenusStrings['ribbonMenu'];

/** Labels only the Mac Sort dialog shows: its level-list heading, the extra
 *  table columns, the single sort-on choice and the button captions. */
export interface MacSortDialogLabels {
  levelsHeading: string;
  columnHeading: string;
  sortOnHeading: string;
  sortOnValues: string;
  colorHeading: string;
  copyLevel: string;
  options: string;
}

/** Overrides of the Sort wording in `ribbonMenu`, plus the Mac-only labels,
 *  that apply on the Mac host only. */
export type MacSortDialogStrings = Partial<RibbonMenuStrings> & MacSortDialogLabels;

/** The `ribbonMenu` keys the Mac Sort dialog can override. A locale whose Mac
 *  wording is not measured repeats its own shared text. */
const MAC_SORT_DIALOG_KEYS = [
  'sortDialogHeader',
  'sortDialogAscending',
  'sortDialogDescending',
] as const satisfies readonly (keyof RibbonMenuStrings)[];

export const macSortDialogFrom = (
  menu: RibbonMenuStrings,
  labels: MacSortDialogLabels,
): MacSortDialogStrings => ({
  ...Object.fromEntries(MAC_SORT_DIALOG_KEYS.map((key) => [key, menu[key]])),
  ...labels,
});

/** Japanese Mac app wording, measured from Excel for Mac. */
export const macSortDialogJa: MacSortDialogStrings = {
  sortDialogHeader: '先頭行を列見出しとする',
  sortDialogAscending: '昇順 (A ～ Z)',
  sortDialogDescending: '降順 (Z ～ A)',
  levelsHeading: '並べ替えるレベルを追加:',
  columnHeading: '列',
  sortOnHeading: '並べ替えのキー',
  sortOnValues: '値',
  colorHeading: '色/アイコン',
  copyLevel: 'コピー',
  options: 'オプション...',
};
