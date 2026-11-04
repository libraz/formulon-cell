import { updatePivotStyleInfo } from '../../commands/pivot-style-info.js';
import { findPivotTableAtCell } from '../../engine/passthrough-sync.js';
import { builtinPivotStyleElements } from '../../engine/pivot-style.js';
import type { StyleElementFormat } from '../../engine/table-style-elements.js';
import {
  createInteractionButton,
  createInteractionChipButton,
} from '../../interact/chip-button.js';
import type { SpreadsheetInstance } from '../../mount/types.js';
import { projectDisabledState } from '../menu-a11y.js';
import type { ToolbarLang } from '../ribbon-model.js';

const FAMILIES = [
  { id: 'Light', ja: '淡色', en: 'Light' },
  { id: 'Medium', ja: '中間', en: 'Medium' },
  { id: 'Dark', ja: '濃色', en: 'Dark' },
] as const;
const PER_FAMILY = 28;
const PAGE_SIZE = 6;

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

interface GalleryStyle {
  /** The style element's name; empty for no style. */
  name: string;
  label: string;
}

const galleryStyles = (lang: ToolbarLang): GalleryStyle[] => {
  const none = lang === 'ja' ? 'なし' : 'None';
  const family = (f: (typeof FAMILIES)[number]): string => (lang === 'ja' ? f.ja : f.en);
  const word = lang === 'ja' ? 'ピボット スタイル' : 'PivotTable Style';
  return [
    { name: '', label: none },
    ...FAMILIES.flatMap((f) =>
      Array.from({ length: PER_FAMILY }, (_, i) => {
        const colour = lang === 'ja' ? JA_COLOUR_NAMES[f.id][i] : undefined;
        return {
          name: `PivotStyle${f.id}${i + 1}`,
          label: `${colour ? `${colour}, ` : ''}${word} (${family(f)}) ${i + 1}`,
        };
      }),
    ),
  ];
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

/** A small report drawn in `name`: header band, plain rows, total band. */
const thumbnail = (name: string, themeColors: readonly number[] | undefined): HTMLDivElement => {
  const thumb = document.createElement('div');
  thumb.className = 'fc-tb__pivotstyle-thumb';
  const elements = name ? builtinPivotStyleElements(name, themeColors) : null;
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

export interface PivotStyleGalleryOptions {
  lang: ToolbarLang;
  getInst: () => SpreadsheetInstance | null;
}

/** The Design tab's Quick Styles: None and the built-in styles, a page at a time, with arrows. */
export const createPivotStyleGallery = (opts: PivotStyleGalleryOptions): HTMLDivElement => {
  const { lang, getInst } = opts;
  const styles = galleryStyles(lang);
  const wrap = document.createElement('div');
  wrap.className = 'fc-tb__pivotstyle-gallery';
  wrap.dataset.ribbonCommand = 'pivotStyleGallery';
  const strip = document.createElement('div');
  strip.className = 'fc-tb__pivotstyle-strip';
  strip.setAttribute('role', 'group');
  const prev = createInteractionButton({
    className: 'fc-tb__pivotstyle-prev',
    ariaLabel: lang === 'ja' ? '前のクイック スタイル ギャラリー' : 'Previous Quick Styles gallery',
  });
  const next = createInteractionButton({
    className: 'fc-tb__pivotstyle-next',
    ariaLabel: lang === 'ja' ? '次のクイック スタイル ギャラリー' : 'Next Quick Styles gallery',
  });
  wrap.append(prev, strip, next);

  let page = 0;
  const lastPage = Math.ceil(styles.length / PAGE_SIZE) - 1;
  const activePivot = (): { sheetIndex: number; pivotIndex: number } | null => {
    const inst = getInst();
    return inst
      ? findPivotTableAtCell(inst.workbook, inst.store.getState().selection.active)
      : null;
  };
  const currentName = (): string | null => {
    const inst = getInst();
    const pivot = activePivot();
    if (!inst || !pivot) return null;
    return inst.workbook.getPivotStyleInfo(pivot.sheetIndex, pivot.pivotIndex)?.name ?? null;
  };

  const render = (): void => {
    const inst = getInst();
    const theme = inst?.workbook.getTheme()?.colors;
    const current = currentName();
    strip.replaceChildren(
      ...styles.slice(page * PAGE_SIZE, page * PAGE_SIZE + PAGE_SIZE).map((style) => {
        const chip = createInteractionChipButton({
          className: 'fc-tb__pivotstyle-chip',
          label: '',
          dataset: { pivotStyle: style.name },
          pressed: current === style.name || (current === null && style.name === ''),
        });
        chip.setAttribute('aria-label', style.label);
        chip.title = style.label;
        chip.appendChild(thumbnail(style.name, theme));
        chip.addEventListener('click', () => {
          const target = getInst();
          const pivot = activePivot();
          if (!target || !pivot) return;
          updatePivotStyleInfo(
            { store: target.store, workbook: target.workbook, history: target.history },
            pivot,
            { name: style.name },
          );
          target.host.focus();
          render();
        });
        return chip;
      }),
    );
    projectDisabledState(prev, page === 0, null);
    projectDisabledState(next, page === lastPage, null);
  };
  // The ribbon refreshes its controls when the selection or the report changes.
  wrap.addEventListener('fc-ribbon-control-refresh', () => render());
  prev.addEventListener('click', () => {
    page = Math.max(0, page - 1);
    render();
  });
  next.addEventListener('click', () => {
    page = Math.min(lastPage, page + 1);
    render();
  });
  // Open on the page holding the report's current style.
  const currentIndex = styles.findIndex((style) => style.name === currentName());
  if (currentIndex >= 0) page = Math.floor(currentIndex / PAGE_SIZE);
  render();
  return wrap;
};
