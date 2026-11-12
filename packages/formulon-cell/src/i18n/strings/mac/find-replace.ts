import type { EditDialogsStrings } from '../edit-dialogs/types.js';

type FindReplaceStrings = EditDialogsStrings['findReplace'];

/** Overrides of `findReplace` that apply on the Mac host only, plus the
 *  half-width / full-width option the Mac app adds. */
export type MacFindReplaceStrings = Partial<FindReplaceStrings> & { matchWidth: string };

/** The `findReplace` keys the Mac host can override. Each locale starts from its own shared text and overrides the fields the Excel bundle answers. */
const MAC_FIND_REPLACE_KEYS = [
  'optionsMore',
  'optionsLess',
  'search',
  'lookIn',
  'matchCase',
  'matchEntire',
  'cellsFound',
] as const satisfies readonly (keyof FindReplaceStrings)[];

export const macFindReplaceFrom = (
  dialog: FindReplaceStrings,
  matchWidth: string,
): MacFindReplaceStrings => ({
  ...Object.fromEntries(MAC_FIND_REPLACE_KEYS.map((key) => [key, dialog[key]])),
  matchWidth,
});

/** Mac app wording from the Excel bundle; not answered by it: optionsMore, optionsLess, matchEntire, cellsFound. */
export const macFindReplaceDe: Partial<MacFindReplaceStrings> = {
  lookIn: 'Suchen in:',
  matchCase: 'Groß-/Kleinschreibung beachten',
  matchWidth: 'Byte beachten',
  search: 'Suchen:',
};

/** Mac app wording from the Excel bundle; not answered by it: optionsMore, optionsLess, matchEntire, cellsFound. */
export const macFindReplaceFr: Partial<MacFindReplaceStrings> = {
  matchCase: 'Respecter la casse',
  matchWidth: 'Respecter le nombre d’octets',
};

/** Mac app wording from the Excel bundle; not answered by it: optionsMore, optionsLess, matchEntire, cellsFound. */
export const macFindReplaceKo: Partial<MacFindReplaceStrings> = {
  lookIn: '찾는 위치:',
  matchCase: '대/소문자 구분',
  matchWidth: '전자/반자 구분',
};

/** Mac app wording from the Excel bundle; not answered by it: optionsMore, optionsLess, matchEntire, cellsFound. */
export const macFindReplaceTh: Partial<MacFindReplaceStrings> = {
  lookIn: 'มองหาใน:',
  matchCase: 'ตรงตามตัวพิมพ์ใหญ่-เล็ก',
  matchWidth: 'ตรงตามไบต์',
};

/** Mac app wording from the Excel bundle; not answered by it: optionsMore, optionsLess, matchEntire, cellsFound. */
export const macFindReplaceZh: Partial<MacFindReplaceStrings> = {
  lookIn: '查找范围:',
  matchCase: '区分大小写',
  matchWidth: '区分全/半角',
};

/** Mac app wording from the Excel bundle. */
export const macFindReplaceEn: Partial<MacFindReplaceStrings> = {
  matchWidth: 'Match byte',
};

/** The Mac floating find panel (⌘F). */
export interface MacFindPanelStrings {
  title: string;
  /** Scope menu item, and the field placeholder while it is chosen. */
  findInSheet: string;
  findInWorkbook: string;
  /** More Options menu items. */
  advancedSearch: string;
  replace: string;
  previous: string;
  next: string;
  moreOptions: string;
  clear: string;
}

/** Mac app wording from the Excel bundle. */
export const macFindPanelEn: MacFindPanelStrings = {
  title: 'Find',
  findInSheet: 'Find in Sheet',
  findInWorkbook: 'Find in Workbook',
  advancedSearch: 'Advanced Search...',
  replace: 'Replace...',
  previous: 'Previous',
  next: 'Next',
  moreOptions: 'More Options',
  clear: 'Clear',
};

/** Mac app wording from the Excel bundle. */
export const macFindPanelJa: MacFindPanelStrings = {
  title: '検索する',
  findInSheet: 'シート内を検索する',
  findInWorkbook: 'ブック内を検索する',
  advancedSearch: '高度な検索...',
  replace: '置換...',
  previous: '前へ',
  next: '次へ',
  moreOptions: 'その他のオプション',
  clear: 'クリア',
};

/** Mac app wording from the Excel bundle. */
export const macFindPanelDe: MacFindPanelStrings = {
  title: 'Finden',
  findInSheet: 'In Blatt suchen',
  findInWorkbook: 'In Arbeitsmappe suchen',
  advancedSearch: 'Erweiterte Suche...',
  replace: 'Ersetzen...',
  previous: 'Zurück',
  next: 'Weiter',
  moreOptions: 'Weitere Optionen',
  clear: 'Löschen',
};

/** Mac app wording from the Excel bundle. */
export const macFindPanelFr: MacFindPanelStrings = {
  title: 'Rechercher',
  findInSheet: 'Rechercher dans la feuille',
  findInWorkbook: 'Rechercher dans le classeur',
  advancedSearch: 'Recherche avancée...',
  replace: 'Remplacer...',
  previous: 'Précédent',
  next: 'Suivant',
  moreOptions: 'Autres options',
  clear: 'Effacer',
};

/** Mac app wording from the Excel bundle. */
export const macFindPanelKo: MacFindPanelStrings = {
  title: '찾기',
  findInSheet: '시트에서 찾기',
  findInWorkbook: '통합 문서에서 찾기',
  advancedSearch: '고급 검색...',
  replace: '바꾸기...',
  previous: '이전',
  next: '다음',
  moreOptions: '추가 옵션',
  clear: '지우기',
};

/** Mac app wording from the Excel bundle. */
export const macFindPanelTh: MacFindPanelStrings = {
  title: 'ค้นหา',
  findInSheet: 'ค้นหาในแผ่นงาน',
  findInWorkbook: 'ค้นหาในเวิร์กบุ๊ก',
  advancedSearch: 'ค้นหาขั้นสูง...',
  replace: 'แทนที่...',
  previous: 'ก่อนหน้า',
  next: 'ถัดไป',
  moreOptions: 'ตัวเลือกเพิ่มเติม',
  clear: 'ล้าง',
};

/** Mac app wording from the Excel bundle. */
export const macFindPanelZh: MacFindPanelStrings = {
  title: '查找',
  findInSheet: '在工作表中查找',
  findInWorkbook: '在工作簿中查找',
  advancedSearch: '高级搜索...',
  replace: '替换...',
  previous: '上一个',
  next: '下一个',
  moreOptions: '更多选项',
  clear: '清除',
};
