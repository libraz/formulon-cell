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

/** Mac app wording from the Excel bundle; not answered by it: optionsMore, optionsLess, search, matchEntire, cellsFound, matchWidth. */
export const macFindReplaceDe: Partial<MacFindReplaceStrings> = {
  lookIn: 'Suchen in:',
  matchCase: 'Groß-/Kleinschreibung beachten',
};

/** Mac app wording from the Excel bundle; not answered by it: optionsMore, optionsLess, search, lookIn, matchEntire, cellsFound, and 1 more. */
export const macFindReplaceFr: Partial<MacFindReplaceStrings> = {
  matchCase: 'Respecter la casse',
};

/** Mac app wording from the Excel bundle; not answered by it: optionsMore, optionsLess, search, matchEntire, cellsFound, matchWidth. */
export const macFindReplaceKo: Partial<MacFindReplaceStrings> = {
  lookIn: '찾는 위치:',
  matchCase: '대/소문자 구분',
};

/** Mac app wording from the Excel bundle; not answered by it: optionsMore, optionsLess, search, matchEntire, cellsFound, matchWidth. */
export const macFindReplaceTh: Partial<MacFindReplaceStrings> = {
  lookIn: 'มองหาใน:',
  matchCase: 'ตรงตามตัวพิมพ์ใหญ่-เล็ก',
};

/** Mac app wording from the Excel bundle; not answered by it: optionsMore, optionsLess, search, matchEntire, cellsFound, matchWidth. */
export const macFindReplaceZh: Partial<MacFindReplaceStrings> = {
  lookIn: '查找范围:',
  matchCase: '区分大小写',
};
