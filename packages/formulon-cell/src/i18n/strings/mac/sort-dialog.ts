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
  /** Caption of the key list's first column once the sort runs left to right. */
  rowHeading: string;
  orientation: string;
  sortTopToBottom: string;
  sortLeftToRight: string;
  method: string;
  usePhonetic: string;
  noPhonetic: string;
  caseSensitive: string;
}

/** Overrides of the Sort wording in `ribbonMenu`, plus the Mac-only labels,
 *  that apply on the Mac host only. */
export type MacSortDialogStrings = Partial<RibbonMenuStrings> & MacSortDialogLabels;

/** The `ribbonMenu` keys the Mac Sort dialog can override. Each locale starts from its own shared text and overrides the fields the Excel bundle answers. */
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
  rowHeading: '行',
  orientation: '方向',
  sortTopToBottom: '行単位',
  sortLeftToRight: '列単位',
  method: '方法',
  usePhonetic: 'ふりがなを使う',
  noPhonetic: 'ふりがなを使わない',
  caseSensitive: '大文字と小文字を区別する',
};

/** Mac app wording from the Excel bundle; not answered by it: sortDialogHeader, levelsHeading, columnHeading, sortOnValues, copyLevel, options, and 5 more. */
export const macSortDialogDe: Partial<MacSortDialogStrings> = {
  sortDialogAscending: 'A bis Z',
  sortDialogDescending: 'Z bis A',
  sortOnHeading: 'Sortieren nach',
  colorHeading: 'Farbe/Symbol',
  sortTopToBottom: 'Zeilen sortieren',
  sortLeftToRight: 'Spalten sortieren',
  caseSensitive: 'Groß-/Kleinschreibung beachten',
};

/** Mac app wording from the Excel bundle; not answered by it: sortDialogHeader, levelsHeading, columnHeading, sortOnValues, copyLevel, options, and 6 more. */
export const macSortDialogFr: Partial<MacSortDialogStrings> = {
  sortDialogAscending: 'De A à Z',
  sortDialogDescending: 'De Z à A',
  sortOnHeading: 'Trier sur',
  colorHeading: 'Couleur/Icône',
  sortTopToBottom: 'Du haut vers le bas',
  sortLeftToRight: 'De la gauche vers la droite',
};

/** Mac app wording from the Excel bundle; not answered by it: sortDialogHeader, levelsHeading, columnHeading, sortOnValues, copyLevel, options, and 5 more. */
export const macSortDialogKo: Partial<MacSortDialogStrings> = {
  sortDialogAscending: '텍스트 오름차순',
  sortDialogDescending: '텍스트 내림차순',
  sortOnHeading: '정렬 기준',
  colorHeading: '색상/아이콘',
  sortTopToBottom: '위쪽에서 아래쪽',
  sortLeftToRight: '왼쪽에서 오른쪽',
  caseSensitive: '대/소문자 구분',
};

/** Mac app wording from the Excel bundle; not answered by it: sortDialogHeader, levelsHeading, columnHeading, sortOnValues, copyLevel, options, and 5 more. */
export const macSortDialogTh: Partial<MacSortDialogStrings> = {
  sortDialogAscending: 'ก ถึง ฮ',
  sortDialogDescending: 'ฮ ถึง ก',
  sortOnHeading: 'เรียงลำดับ',
  colorHeading: 'สี/ไอคอน',
  sortTopToBottom: 'เรียงลำดับจากบนลงล่าง',
  sortLeftToRight: 'เรียงลำดับจากซ้ายไปขวา',
  caseSensitive: 'ตรงตามตัวพิมพ์ใหญ่-เล็ก',
};

/** Mac app wording from the Excel bundle; not answered by it: sortDialogHeader, sortDialogAscending, sortDialogDescending, levelsHeading, columnHeading, sortOnValues, and 6 more. */
export const macSortDialogZh: Partial<MacSortDialogStrings> = {
  sortOnHeading: '排序依据',
  colorHeading: '颜色/图标',
  copyLevel: '复制',
  sortTopToBottom: '按列排序',
  sortLeftToRight: '按行排序',
  caseSensitive: '区分大小写',
};
