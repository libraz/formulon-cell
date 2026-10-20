/**
 * Excel's automatic series colours, as the engine computes them: the theme accents
 * in turn, each round of six in the next luminance variation of the default colour
 * style, applied as the engine reads a tinted colour back.
 */

/** accent1..accent6 of the engine's default theme (Office 2023). */
export const DEFAULT_THEME_ACCENTS: readonly string[] = [
  '#156082',
  '#e97132',
  '#196b24',
  '#0f9ed5',
  '#a02b93',
  '#4ea72e',
];

const CSS_HEX_COLOR = /^#([0-9a-f]{3}|[0-9a-f]{6})$/i;

/** `#rrggbb` from an engine ARGB colour; the alpha byte is carried by transparency. */
export const cssColorFromArgb = (argb: number): string =>
  `#${(argb & 0xffffff).toString(16).padStart(6, '0')}`;

/** Opaque ARGB from `#rgb` / `#rrggbb`; null for any other CSS colour. */
export const argbFromCssColor = (color: string): number | null => {
  const digits = color.trim().match(CSS_HEX_COLOR)?.[1];
  if (digits === undefined) return null;
  const hex =
    digits.length === 3 ? [...digits].map((digit) => `${digit}${digit}`).join('') : digits;
  return (0xff000000 | Number.parseInt(hex, 16)) >>> 0;
};

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

const HLS_MAX = 240;
const RGB_MAX = 255;

const hueToRgb = (n1: number, n2: number, hue: number): number => {
  const h = ((hue % HLS_MAX) + HLS_MAX) % HLS_MAX;
  if (h < HLS_MAX / 6) return n1 + Math.trunc(((n2 - n1) * h + HLS_MAX / 12) / (HLS_MAX / 6));
  if (h < HLS_MAX / 2) return n2;
  if (h < (HLS_MAX * 2) / 3) {
    return n1 + Math.trunc(((n2 - n1) * ((HLS_MAX * 2) / 3 - h) + HLS_MAX / 12) / (HLS_MAX / 6));
  }
  return n1;
};

/** `argb` with an Excel tint (-1..1) applied to its luminance, in the engine's integer HLS. */
const tintedArgb = (argb: number, tint: number): number => {
  if (tint === 0) return (argb | 0xff000000) >>> 0;
  const r = (argb >>> 16) & 0xff;
  const g = (argb >>> 8) & 0xff;
  const b = argb & 0xff;
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const lum = Math.trunc(((max + min) * HLS_MAX + RGB_MAX) / (2 * RGB_MAX));
  let sat = 0;
  let hue = 0;
  if (max !== min) {
    sat =
      lum <= HLS_MAX / 2
        ? Math.trunc(((max - min) * HLS_MAX + (max + min) / 2) / (max + min))
        : Math.trunc(
            ((max - min) * HLS_MAX + (2 * RGB_MAX - max - min) / 2) / (2 * RGB_MAX - max - min),
          );
    const span = max - min;
    const delta = (c: number): number => Math.trunc(((max - c) * (HLS_MAX / 6) + span / 2) / span);
    if (r === max) hue = delta(b) - delta(g);
    else if (g === max) hue = HLS_MAX / 3 + delta(r) - delta(b);
    else hue = (2 * HLS_MAX) / 3 + delta(g) - delta(r);
    if (hue < 0) hue += HLS_MAX;
    if (hue > HLS_MAX) hue -= HLS_MAX;
  }
  const keep = tint < 0 ? 1 + tint : 1 - tint;
  const scaled = tint < 0 ? lum * keep : lum * keep + (HLS_MAX - HLS_MAX * keep);
  const newLum = Math.min(HLS_MAX, Math.max(0, Math.floor(scaled)));
  const pack = (red: number, green: number, blue: number): number =>
    (0xff000000 | (red << 16) | (green << 8) | blue) >>> 0;
  if (sat === 0) {
    const v = Math.trunc((newLum * RGB_MAX + HLS_MAX / 2) / HLS_MAX);
    return pack(v, v, v);
  }
  const magic2 =
    newLum <= HLS_MAX / 2
      ? Math.trunc((newLum * (HLS_MAX + sat) + HLS_MAX / 2) / HLS_MAX)
      : newLum + sat - Math.trunc((newLum * sat + HLS_MAX / 2) / HLS_MAX);
  const magic1 = 2 * newLum - magic2;
  const channel = (h: number): number =>
    Math.trunc((hueToRgb(magic1, magic2, h) * RGB_MAX + HLS_MAX / 2) / HLS_MAX);
  return pack(channel(hue + HLS_MAX / 3), channel(hue), channel(hue - HLS_MAX / 3));
};

/**
 * The automatic colour of series `idx` in Excel's default colour style: the accents
 * in turn, each round of six in the next luminance variation. Luminance is applied
 * as the engine reads it back, as a tint.
 */
export const autoSeriesColor = (accents: readonly string[], idx: number): string | undefined => {
  const base = accents.length >= 6 ? argbFromCssColor(accents[idx % 6] ?? '') : null;
  const [mod, off] = COLOR_VARIATIONS[Math.trunc(idx / 6) % COLOR_VARIATIONS.length] ?? [1, 0];
  if (base === null) return undefined;
  const tint = Math.min(1, Math.max(-1, off > 0 ? off : mod - 1));
  return cssColorFromArgb(tintedArgb(base, tint));
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
