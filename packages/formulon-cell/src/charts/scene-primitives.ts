import type { CategoryChartView, XYChartView } from './filters.js';
import { finite } from './numeric.js';
import type { ChartPrimitive } from './types.js';

/** Excel sets a chart title at 14 pt over 9 or 10 pt chart text. */
const TITLE_FONT_SIZE = 14;

export const TITLE_BASELINE = 16;

/** The top margin of a plot below a title. */
export const TITLE_RESERVE = 30;

const characterWidth = (character: string): number => {
  const code = character.codePointAt(0) ?? 0;
  const wide =
    (code >= 0x1100 && code <= 0x11ff) ||
    (code >= 0x2e80 && code <= 0xa4cf) ||
    (code >= 0xac00 && code <= 0xd7af) ||
    (code >= 0xf900 && code <= 0xfaff) ||
    (code >= 0xff01 && code <= 0xff60) ||
    code >= 0x1f300;
  return code === 0x25a0 ? 9.2 : code === 0x2026 ? 8.2 : wide ? 10 : character === ' ' ? 3.5 : 5.8;
};

/** Rendered width for a summed character width at `fontSize`. */
const scaledWidth = (width: number, fontSize: number): number =>
  Math.max(fontSize === 10 ? 12 : Math.max(1, fontSize * 0.4), width * (finite(fontSize, 10) / 10));

export const textWidth = (text: string, fontSize = 10): number => {
  let width = 0;
  for (const character of Array.from(text)) width += characterWidth(character);
  return scaledWidth(width, fontSize);
};

const ELLIPSIS = '…';

/** Longest prefix of `label` (at most `maxCharacters` code points) that fits `maxWidth` with a trailing ellipsis. */
export const ellipsize = (
  label: string,
  maxWidth: number,
  fontSize = 10,
  maxCharacters = Number.POSITIVE_INFINITY,
): string => {
  const suffixWidth = characterWidth(ELLIPSIS);
  let result = '';
  let width = 0;
  let count = 0;
  for (const character of label) {
    if (count >= maxCharacters) break;
    const next = width + characterWidth(character);
    if (scaledWidth(next + suffixWidth, fontSize) > maxWidth) break;
    result += character;
    width = next;
    count += 1;
  }
  return `${result}${ELLIPSIS}`;
};

export const truncateLabel = (label: string, maxWidth: number, fontSize = 10): string =>
  textWidth(label, fontSize) <= maxWidth ? label : ellipsize(label, maxWidth, fontSize);

export const primitive = (
  tag: ChartPrimitive['tag'],
  attrs: Record<string, string | number>,
  role?: ChartPrimitive['role'],
  seriesIndex?: number,
  pointIndex?: number,
  text?: string,
  identity?: { readonly seriesId?: string; readonly pointId?: string },
): ChartPrimitive => ({
  tag,
  attrs:
    tag === 'text'
      ? {
          fill: 'currentColor',
          'font-family': 'system-ui, sans-serif',
          'font-size': role === 'title' ? TITLE_FONT_SIZE : 10,
          ...attrs,
        }
      : attrs,
  ...(role === undefined ? {} : { role }),
  ...(seriesIndex === undefined ? {} : { seriesIndex }),
  ...(pointIndex === undefined ? {} : { pointIndex }),
  ...(identity?.seriesId === undefined ? {} : { seriesId: identity.seriesId }),
  ...(identity?.pointId === undefined ? {} : { pointId: identity.pointId }),
  ...(text === undefined ? {} : { text }),
});

export const categoryIdAt = (view: CategoryChartView, pointIndex: number): string =>
  view.data.categories[pointIndex]?.id ??
  `category-${view.originalCategoryIndices[pointIndex] ?? pointIndex}`;

export const pointIdAt = (view: XYChartView, seriesIndex: number, pointIndex: number): string => {
  const point = view.data.series[seriesIndex]?.points[pointIndex];
  return point?.id ?? `point-${view.originalPointIndices[seriesIndex]?.[pointIndex] ?? pointIndex}`;
};
