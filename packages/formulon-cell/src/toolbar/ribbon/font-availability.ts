import type { Locale } from '../../i18n/strings.js';
import type { FontFamilyOption } from './font-family-options.js';
// Heuristic font probing for the playground's font-family dropdown. The
// canvas measurement trick compares the text width against the generic
// fallbacks — if it shifts, the font is installed. Results are cached so
// each option only probes once per session.

/**
 * Compatibility exports for hosts that used the old value sets.  Grouping in
 * the UI now comes from FontFamilyOption metadata; there is no recent-font
 * state or submenu implementation to expose here.
 */
export const THEME_FONT_VALUES = new Set([
  'Aptos',
  'Aptos Display',
  'Aptos Narrow',
  '游ゴシック Light',
  '游ゴシック',
]);
export const RECENT_FONT_VALUES = new Set<string>();
export const COMMON_FONT_VALUES = new Set([
  'Calibri',
  'Arial',
  'Segoe UI',
  'Times New Roman',
  'Consolas',
]);
export const FONT_SUBMENU_FAMILIES = new Set<string>();

const fontAvailabilityCache = new Map<string, boolean>();

export const isJapaneseFontName = (value: string): boolean => /[　-鿿]/.test(value);

const fontProbeContext = (): CanvasRenderingContext2D | null => {
  const canvas = document.createElement('canvas');
  return canvas.getContext('2d');
};

export const isFontProbablyAvailable = (font: string): boolean => {
  if (THEME_FONT_VALUES.has(font) || COMMON_FONT_VALUES.has(font)) return true;
  const cached = fontAvailabilityCache.get(font);
  if (cached !== undefined) return cached;
  const ctx = fontProbeContext();
  if (!ctx) return true;
  const sample = 'mmmmmmmmmwwwwwiiiiii 0123456789 あいう漢字';
  const available = ['serif', 'sans-serif', 'monospace'].some((fallback) => {
    ctx.font = `16px ${fallback}`;
    const fallbackWidth = ctx.measureText(sample).width;
    ctx.font = `16px "${font}", ${fallback}`;
    return Math.abs(ctx.measureText(sample).width - fallbackWidth) > 0.5;
  });
  fontAvailabilityCache.set(font, available);
  return available;
};

export const isFontOptionProbablyAvailable = (option: FontFamilyOption): boolean => {
  if (option.alwaysVisible) return true;
  const names = [option.value, ...(option.probeNames ?? [])];
  return names.some((name) => isFontProbablyAvailable(name));
};

/** Filter callback for the font-family dropdown: keep the active value
 *  even if probing says it's unavailable, skip Japanese fonts in English
 *  locale, and otherwise fall back to the canvas probe. */
export const shouldShowFontOption = (
  value: string,
  current: string,
  locale: Locale,
  option?: FontFamilyOption,
): boolean => {
  if (value === current) return true;
  if (locale !== 'ja' && isJapaneseFontName(value)) return false;
  if (option?.alwaysVisible) return true;
  return option ? isFontOptionProbablyAvailable(option) : isFontProbablyAvailable(value);
};
