// Locale-dependent text shared by the function pickers: built-in function
// descriptions, the catalog locale ordinal, and category label keys.

import type {
  CatalogFunctionCategory,
  FunctionCatalogEntry,
} from '../commands/function-categories.js';

/** Concise "spreadsheet-style" descriptions for the most common functions. The
 *  catalog itself (FUNCTION_SIGNATURES) doesn't carry descriptions — anything
 *  not listed here renders without a description blurb. Keep the list small;
 *  exhaustive coverage isn't a goal. */
export const FUNCTION_DESCRIPTIONS: Readonly<Record<string, { en: string; ja: string }>> = {
  SUM: { en: 'Adds its arguments.', ja: '引数の合計を返します。' },
  ACOS: {
    en: 'Returns the arccosine of a number from -1 to 1, in radians.',
    ja: '-1 ～ 1 の数値のアークコサインをラジアンで返します。',
  },
  IF: {
    en: 'Returns one value when a condition is true and another when false.',
    ja: '条件が真のときと偽のときで異なる値を返します。',
  },
  VLOOKUP: {
    en: 'Looks up a value in the leftmost column of a table.',
    ja: '表の左端列で値を検索します。',
  },
  COUNT: { en: 'Counts numeric cells in the range.', ja: '範囲内の数値セルの個数を返します。' },
  COUNTA: {
    en: 'Counts non-empty cells in the range.',
    ja: '範囲内の空白でないセルの個数を返します。',
  },
  COUNTIF: {
    en: 'Counts cells in a range that match a condition.',
    ja: '条件を満たすセルの個数を返します。',
  },
  INDEX: {
    en: 'Returns a value at a given row/column in an array.',
    ja: '配列内の指定された行と列の値を返します。',
  },
  MATCH: {
    en: 'Returns the position of a value in an array.',
    ja: '配列内で一致する値の位置を返します。',
  },
  AVERAGE: { en: 'Returns the arithmetic mean.', ja: '引数の平均値を返します。' },
  MIN: { en: 'Returns the smallest argument.', ja: '引数の最小値を返します。' },
  MAX: { en: 'Returns the largest argument.', ja: '引数の最大値を返します。' },
  ROUND: {
    en: 'Rounds a number to a given precision.',
    ja: '数値を指定した桁数で四捨五入します。',
  },
  IFERROR: {
    en: 'Returns a fallback when the first argument is an error.',
    ja: '式がエラーの場合に代替値を返します。',
  },
  CONCAT: { en: 'Concatenates a list of texts.', ja: '複数の文字列を連結します。' },
  TEXT: {
    en: 'Formats a value as text using a format code.',
    ja: '書式コードに従って数値を文字列に整形します。',
  },
  LEFT: { en: 'Returns the left part of a string.', ja: '文字列の先頭から指定文字数を返します。' },
  RIGHT: {
    en: 'Returns the right part of a string.',
    ja: '文字列の末尾から指定文字数を返します。',
  },
  MID: {
    en: 'Returns characters from the middle of a string.',
    ja: '文字列の中間から指定文字数を返します。',
  },
  LEN: { en: 'Returns the length of a string.', ja: '文字列の文字数を返します。' },
  UPPER: { en: 'Converts a string to upper case.', ja: '文字列を大文字に変換します。' },
  LOWER: { en: 'Converts a string to lower case.', ja: '文字列を小文字に変換します。' },
  AND: { en: 'TRUE only when every argument is TRUE.', ja: 'すべての引数が真のとき真を返します。' },
  OR: { en: 'TRUE when any argument is TRUE.', ja: 'いずれかの引数が真のとき真を返します。' },
  NOT: { en: 'Inverts a boolean.', ja: '論理値を反転します。' },
  ISBLANK: { en: 'Tests whether a value is blank.', ja: '値が空白かどうかを返します。' },
  ISNUMBER: { en: 'Tests whether a value is numeric.', ja: '値が数値かどうかを返します。' },
  NOW: { en: 'Returns the current date and time.', ja: '現在の日付と時刻を返します。' },
  TODAY: { en: "Returns today's date.", ja: '今日の日付を返します。' },
  DATE: {
    en: 'Builds a date from year, month, day.',
    ja: '年・月・日からシリアル値を作成します。',
  },
};

/** `Strings['fxDialog']` key holding each catalog category's title. */
export type CategoryLabelKey =
  | 'categoryLogical'
  | 'categoryLookup'
  | 'categoryText'
  | 'categoryDateTime'
  | 'categoryMath'
  | 'categoryFinancial'
  | 'categoryDynamicArray'
  | 'categoryStatistical'
  | 'categoryEngineering'
  | 'categoryInformation'
  | 'categoryDatabase'
  | 'categoryCompatibility'
  | 'categoryCube'
  | 'categoryWeb';

export const CATEGORY_LABEL_KEY: Readonly<Record<CatalogFunctionCategory, CategoryLabelKey>> = {
  logical: 'categoryLogical',
  lookup: 'categoryLookup',
  text: 'categoryText',
  datetime: 'categoryDateTime',
  math: 'categoryMath',
  financial: 'categoryFinancial',
  dynamicArray: 'categoryDynamicArray',
  statistical: 'categoryStatistical',
  engineering: 'categoryEngineering',
  information: 'categoryInformation',
  database: 'categoryDatabase',
  compatibility: 'categoryCompatibility',
  cube: 'categoryCube',
  web: 'categoryWeb',
};

/** Engine metadata locale for a UI locale tag: 1 for Japanese, 0 otherwise. */
export const catalogLocaleOrdinal = (locale: string): 0 | 1 =>
  locale.trim().toLowerCase().startsWith('ja') ? 1 : 0;

/** Engine-supplied description, else the built-in blurb, else empty. */
export const functionDescription = (
  entry: Pick<FunctionCatalogEntry, 'canonicalName' | 'description'>,
  ordinal: 0 | 1,
): string =>
  entry.description ??
  (ordinal === 1
    ? FUNCTION_DESCRIPTIONS[entry.canonicalName]?.ja
    : FUNCTION_DESCRIPTIONS[entry.canonicalName]?.en) ??
  '';
