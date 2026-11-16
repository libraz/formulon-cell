import {
  listCustomPivotTableStyles,
  pivotStyleGallerySections,
} from '../../commands/format-as-table.js';
import { ribbonStyleTarget } from '../../commands/table-style-actions.js';
import { BUILTIN_PIVOT_STYLE, builtinPivotStyleElements } from '../../engine/pivot-style.js';
import type { StyleElementFormat, StyleElementFormats } from '../../engine/table-style-elements.js';
import { focusSheet } from '../../interact/dialog-focus.js';
import type { ToolbarLang } from '../ribbon-model.js';
import {
  blockedUnlessPermitted,
  createQuickStyleStrip,
  type QuickStyleEntry,
  type RibbonStyleGalleryOptions,
} from './quick-style-strip.js';

const FAMILIES = [
  { id: 'Light', ja: '淡色', en: 'Light' },
  { id: 'Medium', ja: '中間', en: 'Medium' },
  { id: 'Dark', ja: '濃色', en: 'Dark' },
] as const;

/** Colour names Excel (ja) prefixes to each built-in style's accessible name, by family. */
const JA_COLOUR_NAMES: Readonly<Record<(typeof FAMILIES)[number]['id'], readonly string[]>> = {
  Light: [
    ...Array<string>(8).fill('白'),
    ...[
      'アイスブルー',
      'ローズ',
      '薄い緑',
      'ラベンダー',
      'スカイブルー',
      '薄いオレンジ',
      '薄い灰色',
    ],
    ...['アイスブルー', 'ローズ', '薄い緑', 'ラベンダー', 'スカイブルー', '薄いオレンジ'],
    ...Array<string>(7).fill('白'),
  ],
  Medium: [
    ...['白', '薄い青', 'ローズ', '薄い緑', 'ラベンダー', 'スカイブルー', '薄いオレンジ'],
    ...['白', 'アイスブルー', 'ローズ', '薄い緑', 'ラベンダー', 'スカイブルー', '薄いオレンジ'],
    ...['白', 'アイスブルー', 'ローズ', '薄い緑', 'ラベンダー', 'スカイブルー', '薄いオレンジ'],
    ...[
      '薄い灰色',
      'アイスブルー',
      'ローズ',
      '薄い緑',
      'ラベンダー',
      'スカイブルー',
      '薄いオレンジ',
    ],
  ],
  Dark: [
    ...['薄い灰色', '濃い青', '濃い赤', '濃い緑', '濃い紫', '濃い青緑', '茶'],
    ...Array<string>(7).fill('濃い灰色'),
    ...['灰色', '青', '赤', 'オリーブ', '濃い紫', 'アクア', 'オレンジ'],
    ...['灰色', '青', '赤', 'オリーブ', '濃い紫', 'アクア', 'オレンジ'],
  ],
};

/** The accessible name of a built-in PivotTable style, or None for ''. */
export const pivotStyleLabel = (name: string, lang: ToolbarLang): string => {
  if (name === '') return lang === 'ja' ? 'なし' : 'None';
  const match = BUILTIN_PIVOT_STYLE.exec(name);
  const f = FAMILIES.find((family) => family.id === match?.[1]);
  if (!match || !f) return name;
  const index = Number(match[2]);
  const family = lang === 'ja' ? f.ja : f.en;
  const word = lang === 'ja' ? 'ピボット スタイル' : 'PivotTable Style';
  const colour = lang === 'ja' ? JA_COLOUR_NAMES[f.id][index - 1] : undefined;
  return `${colour ? `${colour}, ` : ''}${word} (${family}) ${index}`;
};

const edgeColor = (edge: unknown): string =>
  edge && typeof edge === 'object' && 'color' in edge && typeof edge.color === 'string'
    ? edge.color
    : '#999';

const paint = (cell: HTMLElement, format: StyleElementFormat | undefined): void => {
  if (!format) return;
  if (format.fill) cell.style.background = format.fill;
  if (format.color) cell.style.color = format.color;
  if (format.bold) cell.style.fontWeight = '700';
  if (format.borders?.top) cell.style.borderTop = `1px solid ${edgeColor(format.borders.top)}`;
  if (format.borders?.bottom) {
    cell.style.borderBottom = `1px solid ${edgeColor(format.borders.bottom)}`;
  }
};

/** A small report drawn in `elements`: header band, plain rows, total band. */
export const pivotStyleThumbnail = (elements: StyleElementFormats | null): HTMLDivElement => {
  const thumb = document.createElement('div');
  thumb.className = 'fc-tb__pivotstyle-thumb';
  const rows = 5;
  for (let r = 0; r < rows; r += 1) {
    const kind = r === 0 ? 'header' : r === rows - 1 ? 'total' : 'body';
    for (let c = 0; c < 3; c += 1) {
      const cell = document.createElement('span');
      cell.className = `fc-tb__pivotstyle-cell fc-tb__pivotstyle-cell--${kind}`;
      paint(
        cell,
        kind === 'header' ? elements?.headerRow : kind === 'total' ? elements?.totalRow : undefined,
      );
      thumb.appendChild(cell);
    }
  }
  return thumb;
};

/** The PivotTable Design Quick Styles: custom styles, None and the built-in styles, a page at a
 *  time, with the Quick Styles button as the more arrow. */
export const createPivotStyleGallery = (opts: RibbonStyleGalleryOptions): HTMLDivElement => {
  const { lang, getInst } = opts;
  const target = () => {
    const inst = getInst();
    return inst ? ribbonStyleTarget(inst, 'pivot', 'pivotQuickStyles') : null;
  };
  const wrap = createQuickStyleStrip({
    lang,
    classPrefix: 'fc-tb__pivotstyle',
    kind: 'pivot',
    mac: opts.mac,
    chipKey: 'pivotStyle',
    entries: () => {
      const inst = getInst();
      if (!inst) return [];
      const theme = inst.workbook.getTheme()?.colors;
      return pivotStyleGallerySections(listCustomPivotTableStyles(inst.store.getState()), {
        none: true,
      }).flatMap((section): QuickStyleEntry[] =>
        section.kind === 'custom'
          ? section.styles.map((style) => ({
              id: style.id,
              label: style.label,
              thumbnail: () => pivotStyleThumbnail(style.elements),
            }))
          : section.names.map((name) => ({
              id: name,
              label: pivotStyleLabel(name, lang),
              thumbnail: () => pivotStyleThumbnail(builtinPivotStyleElements(name, theme)),
            })),
      );
    },
    // No report, or one without a style element, shows None as current.
    current: () => target()?.currentId ?? '',
    blocked: blockedUnlessPermitted(getInst, 'pivotQuickStyles'),
    apply: (id) => {
      const picked = target();
      if (!picked) return;
      picked.apply(id);
      const sheetHost = getInst()?.host;
      if (sheetHost) focusSheet(sheetHost);
    },
    more: { arrow: opts.arrow, menu: opts.menu },
  });
  wrap.dataset.ribbonGallery = 'pivotQuickStyles';
  return wrap;
};
