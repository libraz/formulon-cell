/**
 * @file Mac desktop digit metrics for the row-number strip.
 *
 * Column widths come from the engine's Mac profile; this table only sizes the
 * row-number strip, which the engine does not lay out. A digit advance is
 * `em × sizePt` with a per-family em fitted to xlwings measurements of Mac
 * Excel 16.x; families outside the table use {@link UNKNOWN_FAMILY_EM}. Font
 * names are matched after NFKC and case folding.
 */

/** Padding Mac Excel adds around `mdw`-point digits, in points: a quarter
 *  digit each side, rounded up, plus the 1 pt rule. That gives 3 / 5 / 7 / 9 /
 *  11 for 3-4 / 5-8 / 9-12 / 14 / 17 pt digits (Arial 6: 8 × 3 + 3 = 27 pt,
 *  Courier New 20: 8 × 12 + 7 = 103 pt, and the row-number strip). */
export const macColumnPaddingPt = (mdw: number): number => 2 * Math.ceil(mdw / 4) + 1;

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
