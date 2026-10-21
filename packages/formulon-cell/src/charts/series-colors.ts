/**
 * Excel's automatic series colours, as the engine computes them: the theme accents
 * in turn, each round of six in the next luminance variation of the default colour
 * style, applied as the engine reads a tinted colour back.
 */

import { cssColorFromArgb, tintCssColor } from '../engine/theme-color.js';

/** accent1..accent6 of the engine's default theme (Office 2023). */
export const DEFAULT_THEME_ACCENTS: readonly string[] = [
  '#156082',
  '#e97132',
  '#196b24',
  '#0f9ed5',
  '#a02b93',
  '#4ea72e',
];

/** The lumMod / lumOff pairs of Excel's default colour style, one round of accents each. */
const COLOR_VARIATIONS: readonly (readonly [mod: number, off: number])[] = [
  [1, 0],
  [0.6, 0],
  [0.8, 0.2],
  [0.8, 0],
  [0.6, 0.4],
  [0.5, 0],
  [0.7, 0.3],
  [0.7, 0],
  [0.5, 0.5],
];

/**
 * The automatic colour of series `idx` in Excel's default colour style: the accents
 * in turn, each round of six in the next luminance variation. Luminance is applied
 * as the engine reads it back, as a tint.
 */
export const autoSeriesColor = (accents: readonly string[], idx: number): string | undefined => {
  if (accents.length < 6) return undefined;
  const [mod, off] = COLOR_VARIATIONS[Math.trunc(idx / 6) % COLOR_VARIATIONS.length] ?? [1, 0];
  const tint = Math.min(1, Math.max(-1, off > 0 ? off : mod - 1));
  return tintCssColor(accents[idx % 6] ?? '', tint) ?? undefined;
};

/** The automatic colours repeat after six accents in each of nine variations. */
const AUTOMATIC_CYCLE = 54;

/**
 * Excel's automatic colours for idx 0..53, which a palette lookup by idx modulo its
 * length repeats exactly; `accents` falls back to the default theme's.
 */
export const automaticPalette = (accents: readonly string[] = DEFAULT_THEME_ACCENTS): string[] => {
  const source = accents.length >= 6 ? accents : DEFAULT_THEME_ACCENTS;
  return Array.from(
    { length: AUTOMATIC_CYCLE },
    (_, idx) => autoSeriesColor(source, idx) ?? DEFAULT_THEME_ACCENTS[idx % 6] ?? '#156082',
  );
};

/** accent1..accent6 as `#rrggbb` from a theme's colours in `clrScheme` order. */
export const accentsFromTheme = (colors: readonly number[] | undefined): string[] | undefined =>
  colors && colors.length >= 10 ? colors.slice(4, 10).map(cssColorFromArgb) : undefined;

/** The palettes Change Colors offers; `office` stands for the workbook theme's accents. */
export const CHART_PALETTE_PRESETS = {
  office: DEFAULT_THEME_ACCENTS,
  blue: ['#1f4e78', '#5b9bd5', '#9dc3e6', '#d9eaf7', '#17365d', '#2f75b5'],
  mono: ['#1f1f1f', '#4d4d4d', '#737373', '#999999', '#bfbfbf', '#d9d9d9'],
} as const satisfies Record<string, readonly string[]>;

export type ChartPalettePreset = keyof typeof CHART_PALETTE_PRESETS;

/** The colours of `preset` under a workbook theme whose accents are `accents`. */
export const chartPalettePreset = (
  preset: ChartPalettePreset,
  accents?: readonly string[],
): string[] =>
  preset === 'office' ? automaticPalette(accents).slice(0, 6) : [...CHART_PALETTE_PRESETS[preset]];
