/**
 * @file Mac desktop column-width model.
 *
 * The engine's width models are Windows-calibrated; Mac Excel lays a column
 * out as `round(raw × MDW)` points, where `raw` is the stored OOXML width
 * (padding included) and MDW is the Normal font's maximum digit width in
 * whole points. A sheet without `defaultColWidth` uses `baseColWidth × MDW + 5`.
 *
 * MDW is `roundHalfUp(em × sizePt)` with a per-family digit-advance em fitted
 * to xlwings measurements of Mac Excel 16.x; families outside the table use
 * {@link UNKNOWN_FAMILY_EM}. Font names are matched after NFKC and case folding.
 * The padding grows with a large Normal font (see {@link macColumnPaddingPt}).
 */

/** Padding Mac Excel adds to a column's digit cells at the default fonts, in points. */
export const MAC_COLUMN_PADDING_PT = 5;

/** Padding Mac Excel adds around `mdw`-point digits, in points: a quarter
 *  digit each side, rounded up, plus the 1 pt rule. That gives 7 / 9 / 11 for
 *  9-12 / 14 / 17 pt digits (Courier New 20: 8 × 12 + 7 = 103 pt, and the row
 *  number strip). Below the measured 5-12 pt digits it stays 5. */
export const macColumnPaddingPt = (mdw: number): number =>
  Math.max(MAC_COLUMN_PADDING_PT, 2 * Math.ceil(mdw / 4) + 1);

/** OOXML stores column widths in 1/256 character steps. */
const RAW_WIDTH_STEPS = 256;

/** Digit-advance em for a family the table does not list. */
const UNKNOWN_FAMILY_EM = 0.55;

const normalizeFontName = (name: string): string => name.normalize('NFKC').trim().toLowerCase();

const FAMILY_EM: ReadonlyMap<string, number> = new Map(
  (
    [
      ['Calibri', 0.507],
      ['Arial', 0.556],
      ['Helvetica Neue', 0.556],
      ['ＭＳ Ｐゴシック', 0.5],
      ['Aptos Narrow', 0.52],
      ['游ゴシック', 0.56],
      ['Meiryo', 0.6],
      ['Tahoma', 0.546],
      ['等线', 0.5],
      ['맑은 고딕', 0.56],
      ['Courier New', 0.6],
    ] as const
  ).map(([name, em]) => [normalizeFontName(name), em]),
);

/** Alternate spellings, keyed and valued by normalized name. */
const FAMILY_ALIASES: ReadonlyMap<string, string> = new Map(
  (
    [
      ['Yu Gothic', '游ゴシック'],
      ['メイリオ', 'Meiryo'],
      ['MS PGothic', 'ＭＳ Ｐゴシック'],
      ['DengXian', '等线'],
      ['Malgun Gothic', '맑은 고딕'],
    ] as const
  ).map(([alias, name]) => [normalizeFontName(alias), normalizeFontName(name)]),
);

/** Digit advance in points, before Mac Excel rounds it to whole points. */
export function macDigitAdvance(fontName: string, sizePt: number): number {
  const key = normalizeFontName(fontName);
  const em = FAMILY_EM.get(FAMILY_ALIASES.get(key) ?? key) ?? UNKNOWN_FAMILY_EM;
  return em * sizePt;
}

/** Whole-point digit width for a digit advance, rounding half up. */
export const roundDigitAdvance = (advance: number): number => Math.floor(advance + 0.5);

/** Maximum digit width in whole points for a Normal font on Mac. */
export function macMdw(fontName: string, sizePt: number): number {
  return roundDigitAdvance(macDigitAdvance(fontName, sizePt));
}

/** Displayed width in points of a stored raw column width. */
export function macRawToPt(raw: number, mdw: number): number {
  return Math.round(raw * mdw);
}

/** Raw width Mac Excel saves for a column `pt` wide, after snapping to whole points. */
export function macPtToRaw(pt: number, mdw: number): number {
  const width = Math.round(pt);
  if (width <= 0) return 0;
  return Math.trunc((width / mdw) * RAW_WIDTH_STEPS) / RAW_WIDTH_STEPS;
}

/** Width in points of a default column on a sheet that states no `defaultColWidth`. */
export function macDefaultColumnPt(baseColWidth: number, mdw: number): number {
  return baseColWidth * mdw + macColumnPaddingPt(mdw);
}
