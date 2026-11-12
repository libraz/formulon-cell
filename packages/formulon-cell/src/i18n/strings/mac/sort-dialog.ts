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

/** Mac app wording from the Excel bundle; not answered by it: columnHeading, sortOnValues, copyLevel, options, rowHeading, orientation. */
export const macSortDialogDe: Partial<MacSortDialogStrings> = {
  sortDialogHeader: 'Meine Liste hat Kopfzeilen',
  sortDialogAscending: 'A bis Z',
  sortDialogDescending: 'Z bis A',
  levelsHeading: 'Ebenen für die Sortierung hinzufügen:',
  sortOnHeading: 'Sortieren nach',
  colorHeading: 'Farbe/Symbol',
  sortTopToBottom: 'Zeilen sortieren',
  sortLeftToRight: 'Spalten sortieren',
  usePhonetic: 'Silbentabelle',
  noPhonetic: 'Zeichenfolge',
  caseSensitive: 'Groß-/Kleinschreibung beachten',
};

/** Mac app wording from the Excel bundle; not answered by it: columnHeading, sortOnValues, copyLevel, options, rowHeading, orientation. */
export const macSortDialogFr: Partial<MacSortDialogStrings> = {
  sortDialogHeader: 'Ma liste contient des en-têtes',
  sortDialogAscending: 'De A à Z',
  sortDialogDescending: 'De Z à A',
  levelsHeading: 'Ajouter des niveaux pour trier par :',
  sortOnHeading: 'Trier sur',
  colorHeading: 'Couleur/Icône',
  sortTopToBottom: 'Du haut vers le bas',
  sortLeftToRight: 'De la gauche vers la droite',
  usePhonetic: 'Trier par syllabe',
  noPhonetic: 'Trier par caractère',
  method: 'Méthode ',
};

/** Mac app wording from the Excel bundle; not answered by it: columnHeading, sortOnValues, copyLevel, options, rowHeading, orientation. */
export const macSortDialogKo: Partial<MacSortDialogStrings> = {
  sortDialogHeader: '목록에 머리글 포함',
  sortDialogAscending: '텍스트 오름차순',
  sortDialogDescending: '텍스트 내림차순',
  levelsHeading: '정렬 기준에 수준 추가:',
  sortOnHeading: '정렬 기준',
  colorHeading: '색상/아이콘',
  sortTopToBottom: '위쪽에서 아래쪽',
  sortLeftToRight: '왼쪽에서 오른쪽',
  usePhonetic: '한글 사전순',
  noPhonetic: '획순',
  caseSensitive: '대/소문자 구분',
};

/** Mac app wording from the Excel bundle; not answered by it: columnHeading, sortOnValues, copyLevel, options, rowHeading, orientation. */
export const macSortDialogTh: Partial<MacSortDialogStrings> = {
  sortDialogHeader: 'รายการของฉันมีส่วนหัว',
  sortDialogAscending: 'ก ถึง ฮ',
  sortDialogDescending: 'ฮ ถึง ก',
  levelsHeading: 'เพิ่มระดับเพื่อเรียงลำดับตาม:',
  sortOnHeading: 'เรียงลำดับ',
  colorHeading: 'สี/ไอคอน',
  sortTopToBottom: 'เรียงลำดับจากบนลงล่าง',
  sortLeftToRight: 'เรียงลำดับจากซ้ายไปขวา',
  usePhonetic: 'การเรียงลำดับตามพยางค์',
  noPhonetic: 'การเรียงลำดับตามขีด',
  caseSensitive: 'ตรงตามตัวพิมพ์ใหญ่-เล็ก',
  method: 'วิธี',
};

/** Mac app wording from the Excel bundle; not answered by it: sortDialogAscending, sortDialogDescending, columnHeading, sortOnValues, options, rowHeading, and 1 more. */
export const macSortDialogZh: Partial<MacSortDialogStrings> = {
  sortDialogHeader: '列表包含标题',
  levelsHeading: '添加按以下方式排序的级别:',
  sortOnHeading: '排序依据',
  colorHeading: '颜色/图标',
  copyLevel: '复制',
  sortTopToBottom: '按列排序',
  sortLeftToRight: '按行排序',
  usePhonetic: '字母顺序',
  noPhonetic: '笔划顺序',
  caseSensitive: '区分大小写',
};

/** Mac app wording from the Excel bundle. */
export const macSortDialogEn: Partial<MacSortDialogStrings> = {
  sortDialogHeader: 'My list has headers',
  levelsHeading: 'Add levels to sort by:',
  usePhonetic: 'Syllabary Ordering',
  noPhonetic: 'Stroke Ordering',
};
