import type { GridPopupsStrings } from '../grid-popups/types.js';

type FilterDropdownStrings = GridPopupsStrings['filterDropdown'];

/** Wording of the Mac column filter panel: the condition list, search and footer. */
export interface MacFilterPanelStrings {
  chooseOne: string;
  equals: string;
  notEquals: string;
  contains: string;
  notContains: string;
  greaterThan: string;
  greaterThanOrEqual: string;
  lessThan: string;
  lessThanOrEqual: string;
  between: string;
  top10: string;
  bottom10: string;
  aboveAverage: string;
  belowAverage: string;
  /** Units of the Top/Bottom count. */
  items: string;
  percent: string;
  search: string;
  selectAll: string;
  autoApply: string;
  apply: string;
  clear: string;
  sortAsc: string;
  sortDesc: string;
  byColor: string;
  colorNone: string;
  cellColor: string;
  fontColor: string;
  cellIcon: string;
  /** Bold section labels above the sort and filter controls, and the close button's name. */
  sortSection: string;
  filterSection: string;
  close: string;
  /** Placeholder row for cells with no value. */
  blanks: string;
  /** Why the conditions the host has no filter for cannot be used. */
  unavailable: string;
}

type MacFilterPanelExtras = Pick<
  MacFilterPanelStrings,
  | 'sortSection'
  | 'filterSection'
  | 'close'
  | 'sortAsc'
  | 'sortDesc'
  | 'byColor'
  | 'colorNone'
  | 'cellColor'
  | 'fontColor'
  | 'cellIcon'
  | 'autoApply'
  | 'between'
  | 'top10'
  | 'bottom10'
  | 'aboveAverage'
  | 'belowAverage'
  | 'items'
  | 'percent'
  | 'unavailable'
>;

/** A locale's panel wording: its shared filter-dropdown text plus the lines only the Mac panel has. */
export const macFilterPanelFrom = (
  t: FilterDropdownStrings,
  extras: MacFilterPanelExtras,
  overrides: Partial<MacFilterPanelStrings> = {},
): MacFilterPanelStrings => ({
  chooseOne: t.conditionNone,
  equals: t.conditionEquals,
  notEquals: t.conditionNotEquals,
  contains: t.conditionContains,
  notContains: t.conditionNotContains,
  greaterThan: t.conditionGreaterThan,
  greaterThanOrEqual: t.conditionGreaterThanOrEqual,
  lessThan: t.conditionLessThan,
  lessThanOrEqual: t.conditionLessThanOrEqual,
  search: t.searchPlaceholder,
  selectAll: t.selectAll,
  apply: t.apply,
  clear: t.clear,
  blanks: t.blanks,
  ...extras,
  ...overrides,
});

/** Mac wording for this locale; not measured. */
const extrasEn: MacFilterPanelExtras = {
  sortSection: 'Sort',
  filterSection: 'Filter',
  close: 'Close',
  sortAsc: 'Ascending',
  sortDesc: 'Descending',
  byColor: 'By color:',
  colorNone: 'None',
  cellColor: 'Cell Color',
  fontColor: 'Font Color',
  cellIcon: 'Cell Icon',
  autoApply: 'Auto Apply',
  between: 'Between',
  top10: 'Top 10',
  bottom10: 'Bottom 10',
  aboveAverage: 'Above Average',
  belowAverage: 'Below Average',
  items: 'Items',
  percent: 'Percent',
  unavailable: 'This feature is not implemented yet.',
};

/** Japanese Mac app wording, measured from Excel for Mac. */
const extrasJa: MacFilterPanelExtras = {
  sortSection: '並べ替え',
  filterSection: 'フィルター',
  close: '閉じる',
  sortAsc: '昇順',
  sortDesc: '降順',
  byColor: '色別:',
  colorNone: 'なし',
  cellColor: 'セルの色',
  fontColor: 'フォントの色',
  cellIcon: 'セルのアイコン',
  autoApply: '自動適用',
  between: '指定の範囲内',
  top10: '上位 10 項目',
  bottom10: '下位 10 項目',
  aboveAverage: '平均より上',
  belowAverage: '平均より下',
  items: '項目',
  percent: 'パーセント',
  unavailable: 'この機能は未実装です。',
};

/** Mac wording for this locale; not measured. */
const extrasDe: MacFilterPanelExtras = {
  sortSection: 'Sortieren',
  filterSection: 'Filtern',
  close: 'Schließen',
  sortAsc: 'Aufsteigend',
  sortDesc: 'Absteigend',
  byColor: 'Nach Farbe:',
  colorNone: 'Keine',
  cellColor: 'Zellenfarbe',
  fontColor: 'Schriftfarbe',
  cellIcon: 'Zellensymbol',
  autoApply: 'Automatisch anwenden',
  between: 'Zwischen',
  top10: 'Top 10',
  bottom10: 'Untere 10',
  aboveAverage: 'Überdurchschnittlich',
  belowAverage: 'Unterdurchschnittlich',
  items: 'Elemente',
  percent: 'Prozent',
  unavailable: 'Diese Funktion ist noch nicht implementiert.',
};

/** Mac wording for this locale; not measured. */
const extrasFr: MacFilterPanelExtras = {
  sortSection: 'Trier',
  filterSection: 'Filtrer',
  close: 'Fermer',
  sortAsc: 'Croissant',
  sortDesc: 'Décroissant',
  byColor: 'Par couleur :',
  colorNone: 'Aucune',
  cellColor: 'Couleur de cellule',
  fontColor: 'Couleur de police',
  cellIcon: 'Icône de cellule',
  autoApply: 'Appliquer automatiquement',
  between: 'Entre',
  top10: '10 premiers',
  bottom10: '10 derniers',
  aboveAverage: 'Au-dessus de la moyenne',
  belowAverage: 'En dessous de la moyenne',
  items: 'Éléments',
  percent: 'Pourcentage',
  unavailable: 'Cette fonctionnalité n’est pas encore implémentée.',
};

/** Mac wording for this locale; not measured. */
const extrasZh: MacFilterPanelExtras = {
  sortSection: '排序',
  filterSection: '筛选',
  close: '关闭',
  sortAsc: '升序',
  sortDesc: '降序',
  byColor: '按颜色:',
  colorNone: '无',
  cellColor: '单元格颜色',
  fontColor: '字体颜色',
  cellIcon: '单元格图标',
  autoApply: '自动应用',
  between: '介于',
  top10: '前 10 项',
  bottom10: '后 10 项',
  aboveAverage: '高于平均值',
  belowAverage: '低于平均值',
  items: '项',
  percent: '百分比',
  unavailable: '此功能尚未实现。',
};

/** Mac wording for this locale; not measured. */
const extrasKo: MacFilterPanelExtras = {
  sortSection: '정렬',
  filterSection: '필터',
  close: '닫기',
  sortAsc: '오름차순',
  sortDesc: '내림차순',
  byColor: '색 기준:',
  colorNone: '없음',
  cellColor: '셀 색',
  fontColor: '글꼴 색',
  cellIcon: '셀 아이콘',
  autoApply: '자동 적용',
  between: '해당 범위',
  top10: '상위 10개',
  bottom10: '하위 10개',
  aboveAverage: '평균 초과',
  belowAverage: '평균 미만',
  items: '항목',
  percent: '백분율',
  unavailable: '이 기능은 아직 구현되지 않았습니다.',
};

/** Mac wording for this locale; not measured. */
const extrasTh: MacFilterPanelExtras = {
  sortSection: 'เรียงลำดับ',
  filterSection: 'ตัวกรอง',
  close: 'ปิด',
  sortAsc: 'จากน้อยไปมาก',
  sortDesc: 'จากมากไปน้อย',
  byColor: 'ตามสี:',
  colorNone: 'ไม่มี',
  cellColor: 'สีเซลล์',
  fontColor: 'สีแบบอักษร',
  cellIcon: 'ไอคอนเซลล์',
  autoApply: 'ใช้อัตโนมัติ',
  between: 'อยู่ระหว่าง',
  top10: '10 อันดับแรก',
  bottom10: '10 อันดับท้าย',
  aboveAverage: 'สูงกว่าค่าเฉลี่ย',
  belowAverage: 'ต่ำกว่าค่าเฉลี่ย',
  items: 'รายการ',
  percent: 'เปอร์เซ็นต์',
  unavailable: 'ฟีเจอร์นี้ยังไม่ได้ใช้งาน',
};

export const macFilterPanelEn = (t: FilterDropdownStrings): MacFilterPanelStrings =>
  macFilterPanelFrom(t, extrasEn);

export const macFilterPanelJa = (t: FilterDropdownStrings): MacFilterPanelStrings =>
  macFilterPanelFrom(t, extrasJa, {
    chooseOne: '1 つ選択してください',
    equals: '指定の値に等しい',
    notEquals: '指定の値に等しくない',
    greaterThan: '指定の値より大きい',
    greaterThanOrEqual: '指定の値以上',
    lessThan: '指定の値より小さい',
    lessThanOrEqual: '指定の値以下',
    search: '検索',
    selectAll: '(すべて選択)',
    apply: 'フィルターの適用',
    clear: 'フィルターのクリア',
    contains: 'を含む',
    notContains: 'を含まない',
    blanks: '(空白セル)',
  });

export const macFilterPanelDe = (t: FilterDropdownStrings): MacFilterPanelStrings =>
  macFilterPanelFrom(t, extrasDe);

export const macFilterPanelFr = (t: FilterDropdownStrings): MacFilterPanelStrings =>
  macFilterPanelFrom(t, extrasFr);

export const macFilterPanelZh = (t: FilterDropdownStrings): MacFilterPanelStrings =>
  macFilterPanelFrom(t, extrasZh);

export const macFilterPanelKo = (t: FilterDropdownStrings): MacFilterPanelStrings =>
  macFilterPanelFrom(t, extrasKo);

export const macFilterPanelTh = (t: FilterDropdownStrings): MacFilterPanelStrings =>
  macFilterPanelFrom(t, extrasTh);
