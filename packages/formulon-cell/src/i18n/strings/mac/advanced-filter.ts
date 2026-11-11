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

/** The `ribbonMenu` keys the Mac dialog can override. Each locale starts from its own shared text and overrides the fields the Excel bundle answers. */
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

/** Mac app wording from the Excel bundle; not answered by it: advancedFilterListRange, advancedFilterCopyTo. */
export const macAdvancedFilterDe: Partial<MacAdvancedFilterStrings> = {
  advancedFilterCriteriaRange: 'Kriterienbereich:',
  filterInPlace: 'Liste an gleicher Stelle filtern',
  copyToLocation: 'An eine andere Stelle kopieren',
};

/** Mac app wording from the Excel bundle; not answered by it: advancedFilterListRange, advancedFilterCopyTo. */
export const macAdvancedFilterFr: Partial<MacAdvancedFilterStrings> = {
  advancedFilterCriteriaRange: 'Zone de critères :',
  filterInPlace: 'Filtrer la liste sur place',
  copyToLocation: 'Copier à un autre emplacement',
};

/** Mac app wording from the Excel bundle; not answered by it: advancedFilterListRange, advancedFilterCopyTo. */
export const macAdvancedFilterKo: Partial<MacAdvancedFilterStrings> = {
  advancedFilterCriteriaRange: '조건 범위:',
  filterInPlace: '현재 위치에 필터',
  copyToLocation: '다른 장소에 복사',
};

/** Mac app wording from the Excel bundle; not answered by it: advancedFilterListRange, advancedFilterCopyTo. */
export const macAdvancedFilterTh: Partial<MacAdvancedFilterStrings> = {
  advancedFilterCriteriaRange: 'ช่วงเกณฑ์:',
  filterInPlace: 'กรองรายการแล้วใส่ในพื้นที่เดิม',
  copyToLocation: 'คัดลอกไปที่ตำแหน่งอื่น',
};

/** Mac app wording from the Excel bundle; not answered by it: advancedFilterListRange, advancedFilterCopyTo. */
export const macAdvancedFilterZh: Partial<MacAdvancedFilterStrings> = {
  advancedFilterCriteriaRange: '条件区域:',
  filterInPlace: '在原有区域显示筛选结果',
  copyToLocation: '将筛选结果复制到其他位置',
};
