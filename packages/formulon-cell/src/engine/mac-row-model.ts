/**
 * @file Mac desktop default-row model.
 *
 * Mac Excel does not keep a loaded sheet's `defaultRowHeight`: unless
 * `<sheetFormatPr customHeight="1">` pins it, the default row is a whole-point
 * height derived from the Normal font. No closed form fits the measurements, so
 * the heights are tabulated per family and size from xlwings reads of Mac
 * Excel 16.x (`Worksheet.StandardHeight`). Font names match after NFKC and case
 * folding; "Yu Gothic" is a different face from 游ゴシック on Mac and is kept apart.
 */

const normalizeFontName = (name: string): string => name.normalize('NFKC').trim().toLowerCase();

/** Default row height in points by family, keyed by Normal font size in points. */
const FAMILY_ROW_PT: ReadonlyArray<readonly [string, Readonly<Record<number, number>>]> = [
  ['Calibri', { 8: 11, 9: 12, 10: 14, 11: 15, 12: 16, 14: 19, 16: 21, 18: 24 }],
  ['Aptos Narrow', { 10: 14, 11: 15, 12: 16, 14: 19 }],
  ['Aptos', { 10: 14, 11: 15, 12: 16, 14: 19 }],
  ['Arial', { 10: 13, 11: 14, 12: 16, 14: 18 }],
  ['Tahoma', { 10: 13, 11: 14, 12: 15, 14: 18 }],
  ['游ゴシック', { 8: 13, 9: 16, 10: 17, 11: 18, 12: 20, 14: 24, 16: 27, 18: 31 }],
  ['Yu Gothic', { 10: 17, 11: 17, 12: 18, 14: 22 }],
  ['ＭＳ Ｐゴシック', { 8: 12, 9: 13, 10: 14, 11: 14, 12: 15, 14: 17, 16: 19, 18: 22 }],
  ['Meiryo UI', { 8: 11, 9: 14, 10: 15, 11: 16, 12: 18, 14: 20, 16: 22, 18: 25 }],
  ['等线', { 10: 13, 11: 15, 12: 16, 14: 18 }],
  ['맑은 고딕', { 10: 15, 11: 17, 12: 18, 14: 20 }],
  ['新細明體', { 10: 13, 11: 14, 12: 15, 14: 20 }],
  ['微軟正黑體', { 10: 16, 11: 17, 12: 18, 14: 21 }],
];

const ROW_PT_BY_FAMILY: ReadonlyMap<string, Readonly<Record<number, number>>> = new Map(
  FAMILY_ROW_PT.map(([name, sizes]) => [normalizeFontName(name), sizes]),
);

/** Alternate spellings, keyed and valued by normalized name. */
const FAMILY_ALIASES: ReadonlyMap<string, string> = new Map(
  (
    [
      ['MS PGothic', 'ＭＳ Ｐゴシック'],
      ['DengXian', '等线'],
      ['Malgun Gothic', '맑은 고딕'],
      ['PMingLiU', '新細明體'],
      ['Microsoft JhengHei', '微軟正黑體'],
    ] as const
  ).map(([alias, name]) => [normalizeFontName(alias), normalizeFontName(name)]),
);

/** Mac Excel's default row height in points for a Normal font, or null when
 *  the family and size are not tabulated. */
export function macDefaultRowPt(fontName: string, sizePt: number): number | null {
  const key = normalizeFontName(fontName);
  return ROW_PT_BY_FAMILY.get(FAMILY_ALIASES.get(key) ?? key)?.[sizePt] ?? null;
}
