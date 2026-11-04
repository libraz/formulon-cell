import type { RibbonMenusStrings } from '../ribbon-menus/types.js';

type RibbonMenuStrings = RibbonMenusStrings['ribbonMenu'];

/** Action radios only the Mac Advanced Filter dialog shows. */
export interface MacAdvancedFilterLabels {
  filterInPlace: string;
  copyToLocation: string;
}

/** Overrides of the Advanced Filter wording in `ribbonMenu`, plus the Mac-only
 *  radios, that apply on the Mac host only. */
export type MacAdvancedFilterStrings = Partial<RibbonMenuStrings> & MacAdvancedFilterLabels;

/** The `ribbonMenu` keys the Mac dialog can override. A locale whose Mac
 *  wording is not measured repeats its own shared text. */
const MAC_ADVANCED_FILTER_KEYS = [
  'advancedFilterListRange',
  'advancedFilterCriteriaRange',
  'advancedFilterCopyTo',
] as const satisfies readonly (keyof RibbonMenuStrings)[];

export const macAdvancedFilterFrom = (
  menu: RibbonMenuStrings,
  labels: MacAdvancedFilterLabels,
): MacAdvancedFilterStrings => ({
  ...Object.fromEntries(MAC_ADVANCED_FILTER_KEYS.map((key) => [key, menu[key]])),
  ...labels,
});

/** Japanese Mac app wording, measured from Excel for Mac. */
export const macAdvancedFilterJa: MacAdvancedFilterStrings = {
  advancedFilterListRange: 'リスト範囲:',
  advancedFilterCriteriaRange: '検索条件範囲:',
  advancedFilterCopyTo: '抽出範囲:',
  filterInPlace: '選択範囲内',
  copyToLocation: '指定した範囲',
};
